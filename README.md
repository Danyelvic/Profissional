# Eu uso o CustomTkinter para construir a interface gráfica com um visual mais moderno.
import customtkinter as ctk
# Também importo o Tkinter, que é a base da interface gráfica usada pelo CustomTkinter.
import tkinter as tk
# O messagebox me permite mostrar avisos, erros e confirmações em pequenas janelas.
from tkinter import messagebox
# Eu uso datetime para trabalhar com datas, horários, turnos e diferenças de tempo.
import datetime as dt
# O subprocess me permite consultar comandos e serviços do Windows pelo Python.
import subprocess
# Com urllib eu faço uma consulta simples à internet para obter uma referência externa de horário.
import urllib.request
# O email.utils ajuda a converter o cabeçalho de data recebido pela internet para um objeto de data e hora.
import email.utils
# O módulo time é usado para medir a duração da consulta de horário pela internet.
import time
# O webbrowser abre o WhatsApp Web usando o navegador padrão do computador.
import webbrowser
# Eu uso Path para trabalhar com caminhos de arquivos de forma mais organizada.
from pathlib import Path

# O openpyxl é a biblioteca responsável pela leitura da planilha do Excel.
import openpyxl

# ============================================================
# CONFIGURAÇÕES E CAMINHOS
# Aqui eu deixo reunidos os caminhos e valores de configuração usados pelo ControlBombas.
# ============================================================
BASE_DIR = Path(__file__).resolve().parent

PASTA_PROG_BOMBAS = Path(
    r"C:\Users\Terra Brasil\Desktop\Program\Bombas"
)

CAMINHO_EXCEL = Path(
    r"C:\Users\Terra Brasil\Desktop\CONTROLE DAS BOMBAS 01 E 02 (2).xlsx"
)

FONTES_HORARIO_WEB = [
    "https://www.google.com",
    "https://www.cloudflare.com",
    "https://www.microsoft.com",
]

DIFERENCA_MAXIMA_ALERTA = 120  # segundos
TIMEOUT_WEB = 4

# ============================================================
# APARÊNCIA DA INTERFACE
# Nesta parte eu defino o tema visual e preparo a janela principal do programa.
# ============================================================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("ControlBombas • Central de Relatórios")
app.after(0, lambda: app.state("zoomed"))

BG = "#101318"
CARD = "#181D24"
CARD_2 = "#202630"
TEXT = "#F2F4F7"
MUTED = "#9AA4B2"
BLUE = "#3478F6"
BLUE_HOVER = "#2865D8"
GREEN = "#20A464"
GREEN_HOVER = "#188650"
RED = "#D64545"
RED_HOVER = "#B93434"
YELLOW = "#D99A19"
YELLOW_HOVER = "#B98113"
BORDER = "#303844"

# ============================================================
# HISTÓRICO DA SESSÃO
# Eu mantenho uma lista em memória para registrar o que acontece enquanto o programa está aberto.
# ============================================================
historico_sessao = []

# Aqui eu centralizo a leitura da data e hora atual do computador para reutilizar o mesmo padrão no programa inteiro.
def agora_local():
    return dt.datetime.now()

# Nesta função eu registro cada ação importante da sessão, adicionando data e hora ao histórico.
def registrar_acao(texto_acao):
    agora = agora_local().strftime("%d/%m/%Y %H:%M:%S")
    historico_sessao.append(f"[{agora}] {texto_acao}")

# Quando o programa é fechado, eu salvo o histórico da sessão em um arquivo de texto e depois encerro a janela.
def salvar_ao_fechar():
    try:
        if historico_sessao:
            PASTA_PROG_BOMBAS.mkdir(parents=True, exist_ok=True)

            data_hora_nome = agora_local().strftime("%d-%m-%Y_%H-%M-%S")
            caminho = PASTA_PROG_BOMBAS / f"sessao_{data_hora_nome}.txt"

            with open(caminho, "w", encoding="utf-8") as arquivo:
                arquivo.write(
                    "=== HISTÓRICO DE ATIVIDADES DA SESSÃO ===\n\n"
                )
                for item in historico_sessao:
                    arquivo.write(item + "\n")
    except Exception as erro:
        print(f"Erro ao salvar histórico: {erro}")

    app.destroy()

app.protocol("WM_DELETE_WINDOW", salvar_ao_fechar)

# ============================================================
# TRATAMENTO DE DATAS, HORÁRIOS E TURNOS
# As próximas funções deixam datas e horários em um formato confiável para os cálculos do relatório.
# ============================================================
# Aqui eu calculo quanto tempo um equipamento ficou ligado, inclusive quando o horário atravessa a meia-noite.
def calcular_tempo_ligado(hora_inicio, hora_fim):
    formato = "%H:%M:%S"

    try:
        if len(hora_inicio.split(":")) == 2:
            hora_inicio += ":00"

        if len(hora_fim.split(":")) == 2:
            hora_fim += ":00"

        inicio = dt.datetime.strptime(hora_inicio, formato)
        fim = dt.datetime.strptime(hora_fim, formato)

        # Se o horário final for menor que o inicial, eu considero que o período passou pela meia-noite.
        if fim < inicio:
            fim += dt.timedelta(days=1)

        diferenca = fim - inicio
        horas, resto = divmod(diferenca.seconds, 3600)
        minutos, _ = divmod(resto, 60)

        return f"{horas:02d}:{minutos:02d}"

    except (ValueError, TypeError):
        return "Erro"

# Nesta função eu padronizo os horários para o formato HH:MM antes de usá-los no relatório.
def formatar_horario(texto):
    try:
        partes = str(texto).split(":")

        if len(partes) >= 2:
            return f"{partes[0]:0>2}:{partes[1]:0>2}"

    except Exception:
        pass

    return str(texto) if texto else ""

# Como o Excel pode devolver horário em formatos diferentes, eu converto tudo para texto no mesmo padrão.
def format_time_val(valor):
    if not valor:
        return ""

    if isinstance(valor, (dt.time, dt.datetime)):
        return valor.strftime("%H:%M")

    return formatar_horario(valor)

# Aqui eu transformo os diferentes formatos de data vindos do Excel em objetos de data que o Python consegue comparar.
def parse_data_excel(valor):
    if not valor:
        return None

    if isinstance(valor, dt.datetime):
        return valor.date()

    if isinstance(valor, dt.date):
        return valor

    if isinstance(valor, str):
        try:
            return dt.datetime.strptime(
                valor[:10], "%d/%m/%Y"
            ).date()
        except ValueError:
            pass

        try:
            return dt.datetime.strptime(
                valor[:10], "%Y-%m-%d"
            ).date()
        except ValueError:
            pass

    return None

# Nesta função eu descubro qual turno está ativo usando os horários fixos de 07h às 19h e de 19h às 07h.
def obter_turno_atual(agora=None):
    """
    Retorna a janela do turno que contém o horário informado.

    Diurno:  07:00 -> 19:00
    Noturno: 19:00 -> 07:00 do dia seguinte
    """
    agora = agora or agora_local()

    # Se a hora estiver entre 07:00 e 18:59, eu considero que estamos no turno diurno.
    if 7 <= agora.hour < 19:
        inicio = agora.replace(
            hour=7, minute=0, second=0, microsecond=0
        )
        fim = agora.replace(
            hour=19, minute=0, second=0, microsecond=0
        )
        nome = "DIURNO (07h às 19h)"
    # A partir das 19:00 eu monto o turno noturno, que termina às 07:00 do dia seguinte.
    elif agora.hour >= 19:
        inicio = agora.replace(
            hour=19, minute=0, second=0, microsecond=0
        )
        fim = (
            agora + dt.timedelta(days=1)
        ).replace(
            hour=7, minute=0, second=0, microsecond=0
        )
        nome = "NOTURNO (19h às 07h)"
    else:
        inicio = (
            agora - dt.timedelta(days=1)
        ).replace(
            hour=19, minute=0, second=0, microsecond=0
        )
        fim = agora.replace(
            hour=7, minute=0, second=0, microsecond=0
        )
        nome = "NOTURNO (19h às 07h)"

    return inicio, fim, nome

# Aqui eu consigo voltar de 12 em 12 horas para gerar também relatórios de turnos anteriores.
def obter_turno_alvo(deslocamento=0, agora=None):
    """
    deslocamento:
      0 = turno atual
      1 = turno anterior
      2 = dois turnos atrás
      3 = três turnos atrás

    Cada turno tem 12 horas.
    """
    inicio_atual, fim_atual, _ = obter_turno_atual(agora)
    delta = dt.timedelta(hours=12 * deslocamento)

    inicio = inicio_atual - delta
    fim = fim_atual - delta

    if inicio.hour == 7:
        nome = "DIURNO (07h às 19h)"
    else:
        nome = "NOTURNO (19h às 07h)"

    return inicio, fim, nome

# ============================================================
# DIAGNÓSTICO DO HORÁRIO DO WINDOWS
# Esta parte verifica o serviço de horário do Windows e compara o relógio do computador com uma referência externa.
# ============================================================
# Eu uso esta função como uma camada segura para executar comandos do Windows e capturar o resultado sem abrir outra janela.
def executar_comando(comando):
    try:
        resultado = subprocess.run(
            comando,
            capture_output=True,
            text=True,
            encoding="cp850",
            errors="replace",
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            timeout=12,
        )

        saida = (resultado.stdout or "").strip()
        erro = (resultado.stderr or "").strip()

        return resultado.returncode, saida, erro

    except Exception as exc:
        return -1, "", str(exc)

# Aqui eu consulto o serviço W32Time para saber se a sincronização de horário do Windows está funcionando.
def obter_status_w32time():
    codigo, saida, erro = executar_comando(
        ["sc", "query", "w32time"]
    )

    if codigo != 0 and not saida:
        return {
            "status": "ERRO",
            "detalhes": erro or "Não foi possível consultar o serviço."
        }

    texto = saida.upper()

    if "RUNNING" in texto or "EXECUTANDO" in texto:
        status = "RUNNING"
    elif "STOPPED" in texto or "PARADO" in texto:
        status = "STOPPED"
    else:
        status = "DESCONHECIDO"

    return {
        "status": status,
        "detalhes": saida
    }

# Nesta função eu consulto os detalhes da sincronização NTP configurada no Windows.
def obter_status_ntp():
    codigo, saida, erro = executar_comando(
        ["w32tm", "/query", "/status"]
    )

    if codigo != 0:
        return {
            "ok": False,
            "texto": erro or saida or "W32Time indisponível."
        }

    return {
        "ok": True,
        "texto": saida
    }

# Aqui eu descubro qual fonte de horário o Windows está usando naquele momento.
def obter_fonte_ntp():
    codigo, saida, erro = executar_comando(
        ["w32tm", "/query", "/source"]
    )

    if codigo != 0:
        return {
            "ok": False,
            "fonte": erro or saida or "Indisponível"
        }

    return {
        "ok": True,
        "fonte": saida.strip()
    }

# Se o serviço de horário estiver parado, esta função tenta iniciá-lo e registra o resultado no histórico.
def iniciar_w32time():
    codigo, saida, erro = executar_comando(
        ["net", "start", "w32time"]
    )

    texto = "\n".join(
        parte for parte in [saida, erro] if parte
    )

    if codigo == 0:
        registrar_acao("Serviço W32Time iniciado.")
        return True, texto

    estado = obter_status_w32time()

    if estado["status"] == "RUNNING":
        registrar_acao("W32Time já estava em execução.")
        return True, texto or "O serviço já está em execução."

    return False, texto or "Não foi possível iniciar o W32Time."

# Aqui eu peço ao Windows para sincronizar o relógio novamente. A elevação de administrador fica por conta do próprio Windows.
def sincronizar_w32time():
    """Solicita ao Windows, com UAC, a sincronização pelo W32Time."""
    import ctypes

    try:
        parametros = '/c "sc start w32time & w32tm /resync /rediscover"'

        retorno = ctypes.windll.shell32.ShellExecuteW(
            None,
            "runas",
            "cmd.exe",
            parametros,
            None,
            1,
        )

        if retorno <= 32:
            registrar_acao(
                f"Windows recusou a elevação para sincronização. Código: {retorno}."
            )
            return False, (
                "Não foi possível obter permissão de administrador.\n\n"
                "Clique em Sim quando o Windows pedir autorização (UAC) e tente novamente."
            )

        registrar_acao(
            "Comando elevado enviado: sc start w32time + w32tm /resync /rediscover."
        )

        messagebox.showinfo(
            "Sincronização solicitada",
            "O Windows foi autorizado a executar a sincronização.\n\n"
            "Comando executado:\n"
            "w32tm /resync /rediscover\n\n"
            "Aguarde alguns segundos e depois clique em 'Atualizar diagnóstico'."
        )

        return True, "Sincronização solicitada com sucesso ao Windows."

    except Exception as erro:
        registrar_acao(f"Erro ao solicitar sincronização W32Time: {erro}")
        return False, f"Erro ao solicitar sincronização: {erro}"

# Nesta função eu consulto fontes HTTPS e uso o cabeçalho Date como referência externa de horário.
def buscar_horario_web():
    erros = []

    for url in FONTES_HORARIO_WEB:
        try:
            requisicao = urllib.request.Request(
                url,
                method="HEAD",
                headers={
                    "User-Agent": "ControlBombas/2.0"
                },
            )

            inicio = time.time()

            with urllib.request.urlopen(
                requisicao,
                timeout=TIMEOUT_WEB
            ) as resposta:

                data_header = resposta.headers.get("Date")

            if not data_header:
                erros.append(
                    f"{url}: resposta sem cabeçalho Date."
                )
                continue

            horario_utc = email.utils.parsedate_to_datetime(
                data_header
            )

            if horario_utc.tzinfo is None:
                horario_utc = horario_utc.replace(
                    tzinfo=dt.timezone.utc
                )

            horario_local = horario_utc.astimezone()

            latencia = time.time() - inicio

            return {
                "ok": True,
                "datetime": horario_local.replace(tzinfo=None),
                "fonte": url,
                "latencia": latencia,
            }

        except Exception as erro:
            erros.append(f"{url}: {erro}")

    return {
        "ok": False,
        "erro": "\n".join(erros)
    }

# Aqui eu comparo o horário do computador com a referência obtida pela internet e calculo a diferença em segundos.
def comparar_horario_sistema():
    horario_sistema = agora_local()
    resultado = buscar_horario_web()

    if not resultado["ok"]:
        return {
            "ok": False,
            "sistema": horario_sistema,
            "erro": resultado["erro"],
        }

    horario_referencia = resultado["datetime"]
    diferenca = (
        horario_sistema - horario_referencia
    ).total_seconds()

    return {
        "ok": True,
        "sistema": horario_sistema,
        "referencia": horario_referencia,
        "diferenca": diferenca,
        "fonte": resultado["fonte"],
        "latencia": resultado["latencia"],
    }

# ============================================================
# LEITURA DA PLANILHA
# Agora eu começo a trabalhar com os dados do Excel, procurando registros válidos das bombas e do poço.
# ============================================================
# Antes de ler a planilha inteira, eu separo somente as linhas que realmente possuem uma data registrada.
def obter_linhas_com_dados(ws, col_data):
    linhas_com_dados = []
    # Eu limito a busca a 10 mil linhas para evitar percorrer uma quantidade desnecessária de células.
    max_busca = min(ws.max_row, 10000)

    for linha in range(2, max_busca + 1):
        valor = ws.cell(
            row=linha,
            column=col_data
        ).value

        if valor is not None and str(valor).strip() != "":
            linhas_com_dados.append(linha)

    return linhas_com_dados

# Esta é uma das partes principais: eu leio os registros de cada equipamento, monto as datas completas e filtro somente o que pertence ao turno escolhido.
def extrair_linhas_bomba(
    ws,
    nome_equipamento,
    col_data,
    col_lig,
    col_des,
    inicio_turno,
    fim_turno,
    col_porteiro=None,
    porteiro_padrao=None,
):
    eventos = []
    linhas_validas = obter_linhas_com_dados(
        ws,
        col_data
    )

    for linha in linhas_validas:
        data_valor = parse_data_excel(
            ws.cell(
                row=linha,
                column=col_data
            ).value
        )

        if not data_valor:
            continue

        str_ligou = format_time_val(
            ws.cell(
                row=linha,
                column=col_lig
            ).value
        )

        str_desligou = format_time_val(
            ws.cell(
                row=linha,
                column=col_des
            ).value
        )

        if col_porteiro is not None:
            porteiro_registro = ws.cell(
                row=linha,
                column=col_porteiro
            ).value

            if (
                porteiro_registro is None
                or str(porteiro_registro).strip() == ""
            ):
                porteiro_registro = "NÃO INFORMADO"
            else:
                porteiro_registro = str(
                    porteiro_registro
                ).strip()

        else:
            porteiro_registro = (
                porteiro_padrao
                or "NÃO INFORMADO"
            )

        if not str_ligou and not str_desligou:
            continue

        dt_ligou = None

        if str_ligou:
            try:
                hora, minuto = map(
                    int,
                    str_ligou.split(":")
                )

                dt_ligou = dt.datetime.combine(
                    data_valor,
                    dt.time(hora, minuto)
                )

            except (ValueError, TypeError):
                pass

        dt_desligou = None

        if str_desligou:
            try:
                hora, minuto = map(
                    int,
                    str_desligou.split(":")
                )

                dt_desligou = dt.datetime.combine(
                    data_valor,
                    dt.time(hora, minuto)
                )

                if (
                    dt_ligou
                    and dt_desligou < dt_ligou
                ):
                    dt_desligou += dt.timedelta(
                        days=1
                    )

                elif (
                    not dt_ligou
                    and hora < 12
                    and fim_turno.hour == 7
                ):
                    dt_desligou += dt.timedelta(
                        days=1
                    )

            except (ValueError, TypeError):
                pass

        dt_ref = (
            dt_ligou
            if dt_ligou
            else (
                dt_desligou
                or inicio_turno
            )
        )

        if dt_ligou and dt_desligou:
            if (
                dt_ligou <= fim_turno
                and dt_desligou >= inicio_turno
            ):
                tempo = calcular_tempo_ligado(
                    str_ligou,
                    str_desligou
                )

                veio_anterior = (
                    dt_ligou < inicio_turno
                )

                eventos.append({
                    "equipamento": nome_equipamento,
                    "porteiro": porteiro_registro,
                    "linha_excel": linha,
                    "ligou": str_ligou,
                    "desligou": str_desligou,
                    "tempo": tempo,
                    "veio_anterior": veio_anterior,
                    "dt_ref": dt_ref,
                    "dt_ligou_full": dt_ligou,
                    "dt_desligou_full": dt_desligou,
                })

        elif not dt_ligou and dt_desligou:
            if inicio_turno <= dt_desligou <= fim_turno:
                eventos.append({
                    "equipamento": nome_equipamento,
                    "porteiro": porteiro_registro,
                    "linha_excel": linha,
                    "ligou": None,
                    "desligou": str_desligou,
                    "tempo": None,
                    "veio_anterior": False,
                    "dt_ref": dt_ref,
                    "dt_ligou_full": None,
                    "dt_desligou_full": dt_desligou,
                })

        elif dt_ligou and not dt_desligou:
            if (
                dt_ligou <= fim_turno
                and dt_ligou >= (
                    inicio_turno
                    - dt.timedelta(hours=12)
                )
            ):
                eventos.append({
                    "equipamento": nome_equipamento,
                    "porteiro": porteiro_registro,
                    "linha_excel": linha,
                    "ligou": str_ligou,
                    "desligou": None,
                    "tempo": None,
                    "veio_anterior": False,
                    "dt_ref": dt_ref,
                    "dt_ligou_full": dt_ligou,
                    "dt_desligou_full": None,
                })

    return eventos

# Depois de coletar os eventos, eu organizo tudo por horário e transformo os dados no texto final do relatório.
def formatar_relatorio_cronologico(todos_eventos):
    if not todos_eventos:
        return (
            "🔹 *Equipamentos:*\n"
            " • Nenhum acionamento registrado no período.\n"
        )

    # Antes de montar o texto, eu ordeno todos os eventos pela data e pelo horário real de referência.
    todos_eventos.sort(
        key=lambda evento: evento["dt_ref"]
    )

    linhas = []

    for evento in todos_eventos:
        equipamento = evento["equipamento"]
        ligou = evento["ligou"]
        desligou = evento["desligou"]
        tempo = evento["tempo"]
        anterior = evento.get(
            "veio_anterior",
            False
        )

        data_ligou = (
            evento["dt_ligou_full"].strftime("%d/%m")
            if evento.get("dt_ligou_full")
            else ""
        )

        data_desligou = (
            evento["dt_desligou_full"].strftime("%d/%m")
            if evento.get("dt_desligou_full")
            else ""
        )

        porteiro = evento.get(
            "porteiro",
            "NÃO INFORMADO"
        )

        linhas.append(
            f"🔹 *{equipamento}:*"
        )

        linhas.append(
            f" • 👮 Registrado por: *{porteiro}*"
        )

        if ligou and desligou:
            if anterior:
                linhas.append(
                    f" • Ligou às {ligou}h ({data_ligou}) "
                    f"*(turno anterior)* e desligou às "
                    f"{desligou}h ({data_desligou}) "
                    f"*(Tempo total: {tempo}h)*\n"
                )
            else:
                linhas.append(
                    f" • Ligou às {ligou}h ({data_ligou}) "
                    f"e desligou às {desligou}h "
                    f"({data_desligou}) "
                    f"*(Tempo: {tempo}h)*\n"
                )

        elif not ligou and desligou:
            linhas.append(
                f" • Ligou: *(horário não registrado)* "
                f"e desligou às {desligou}h "
                f"({data_desligou})\n"
            )

        elif ligou and not desligou:
            linhas.append(
                f" • Ligou às {ligou}h ({data_ligou}) "
                f"e desligou: *(em andamento - passada "
                f"ao próximo turno)*\n"
            )

    return "\n".join(linhas)

# ============================================================
# WHATSAPP
# O programa abre o WhatsApp Web, mas mantém o envio manual para que o texto seja conferido antes.
# ============================================================
# Aqui eu abro o WhatsApp Web no navegador. O envio continua manual para que o relatório possa ser conferido antes.
def abrir_whatsapp():
    webbrowser.open(
        "https://web.whatsapp.com/"
    )

    registrar_acao(
        "WhatsApp Web aberto manualmente."
    )

# ============================================================
# JANELA DO RELATÓRIO
# Aqui eu monto uma janela separada para mostrar o texto pronto e facilitar a cópia.
# ============================================================
# Nesta função eu crio a janela que mostra o relatório pronto, permitindo copiar o texto para a área de transferência.
def mostrar_relatorio(
    titulo,
    mensagem,
):
    janela = ctk.CTkToplevel(app)
    janela.title("Relatório pronto para copiar")
    janela.geometry("700x620")
    janela.minsize(580, 500)

    janela.transient(app)
    janela.grab_set()
    janela.attributes("-topmost", True)
    janela.lift()
    janela.focus_force()

    titulo_frame = ctk.CTkFrame(
        janela,
        fg_color="transparent"
    )
    titulo_frame.pack(
        fill="x",
        padx=24,
        pady=(20, 5)
    )

    ctk.CTkLabel(
        titulo_frame,
        text="📋 Relatório pronto",
        font=("Arial", 21, "bold"),
        text_color=TEXT,
    ).pack(anchor="w")

    ctk.CTkLabel(
        titulo_frame,
        text=(
            "Copie o texto abaixo e cole manualmente no chat do WhatsApp do síndico."
        ),
        font=("Arial", 12),
        text_color=MUTED,
    ).pack(anchor="w", pady=(3, 0))

    texto = ctk.CTkTextbox(
        janela,
        wrap="word",
        font=("Consolas", 13),
        fg_color="#0C0F13",
        border_width=1,
        border_color=BORDER,
        corner_radius=8,
    )

    texto.pack(
        fill="both",
        expand=True,
        padx=24,
        pady=15
    )

    texto.insert("1.0", mensagem)
    texto.configure(state="disabled")

    instrucoes = ctk.CTkLabel(
        janela,
        text=(
            "1️⃣ Clique em COPIAR RELATÓRIO  •  "
            "2️⃣ Abra o WhatsApp do síndico  •  "
            "3️⃣ Pressione Ctrl+V e confira  •  "
            "4️⃣ Pressione Enter para enviar"
        ),
        font=("Arial", 11),
        text_color=MUTED,
        wraplength=640,
        justify="center",
    )
    instrucoes.pack(
        padx=24,
        pady=(0, 10)
    )

    botoes = ctk.CTkFrame(
        janela,
        fg_color="transparent"
    )

    botoes.pack(
        fill="x",
        padx=24,
        pady=(0, 20)
    )

    status_copia = ctk.CTkLabel(
        botoes,
        text="",
        font=("Arial", 11, "bold"),
        text_color=GREEN,
    )
    status_copia.pack(side="left", padx=5)

    # Esta função interna copia o relatório para a área de transferência e informa na tela se a operação deu certo.
    def copiar_relatorio():
        try:
            # Eu limpo a área de transferência antes de colocar nela o novo relatório.
            app.clipboard_clear()
            app.clipboard_append(mensagem)
            app.update()
            registrar_acao("Relatório copiado para a área de transferência.")
            status_copia.configure(text="✓ Relatório copiado!")
        except Exception as erro:
            registrar_acao(f"Erro ao copiar relatório: {erro}")
            messagebox.showerror(
                "Erro ao copiar",
                f"Não foi possível copiar o relatório.\n\n{erro}"
            )

    ctk.CTkButton(
        botoes,
        text="📋  COPIAR RELATÓRIO",
        command=copiar_relatorio,
        width=230,
        height=45,
        corner_radius=22,
        fg_color=BLUE,
        hover_color=BLUE_HOVER,
        font=("Arial", 13, "bold"),
    ).pack(side="left", padx=5)

    ctk.CTkButton(
        botoes,
        text="FECHAR",
        command=janela.destroy,
        width=130,
        height=45,
        corner_radius=22,
        fg_color=CARD_2,
        hover_color="#2B333F",
        font=("Arial", 13, "bold"),
    ).pack(side="right", padx=5)

    registrar_acao("Janela do relatório exibida para cópia manual.")

# ============================================================
# COMPONENTE REUTILIZÁVEL DE BOTÕES
# Em vez de repetir o mesmo código em várias escolhas da interface, eu concentro esse comportamento em uma classe.
# ============================================================
# Eu criei esta classe para reutilizar grupos de botões que funcionam como uma escolha única, evitando repetir a mesma lógica pela interface.
class ButtonGroup:
    # No construtor eu preparo os botões do grupo e guardo as informações necessárias para controlar a seleção.
    def __init__(
        self,
        parent,
        titles,
        callback=None,
    ):
        self.buttons = []
        self.selected_value = None
        self.callback = callback

        frame = ctk.CTkFrame(
            parent,
            fg_color="transparent"
        )

        frame.pack(
            pady=5,
            fill="x"
        )

        sub_frame = ctk.CTkFrame(
            frame,
            fg_color="transparent"
        )

        sub_frame.pack(
            anchor="center"
        )

        for titulo in titles:
            botao = ctk.CTkButton(
                sub_frame,
                text=titulo,
                width=92,
                height=34,
                corner_radius=18,
                fg_color=CARD_2,
                hover_color="#2B333F",
                border_width=1,
                border_color=BORDER,
                command=lambda valor=titulo:
                    self.select(valor),
            )

            botao.pack(
                side="left",
                padx=4
            )

            self.buttons.append(botao)

    # Aqui eu altero visualmente qual botão está selecionado e executo a função associada, quando existir.
    def select(self, texto):
        self.selected_value = texto

        for botao in self.buttons:
            if botao.cget("text") == texto:
                botao.configure(
                    fg_color=GREEN,
                    hover_color=GREEN_HOVER,
                    border_width=0,
                )
            else:
                botao.configure(
                    fg_color=CARD_2,
                    hover_color="#2B333F",
                    border_width=1,
                    border_color=BORDER,
                )

        if self.callback:
            self.callback(texto)

# ============================================================
# CABEÇALHO
# A partir daqui eu começo a montar visualmente a janela principal do ControlBombas.
# ============================================================
header = ctk.CTkFrame(
    app,
    fg_color=CARD,
    corner_radius=0,
    height=54,
)

header.pack(
    fill="x"
)

header.pack_propagate(False)

header_esquerda = ctk.CTkFrame(
    header,
    fg_color="transparent"
)

header_esquerda.pack(
    side="left",
    padx=20
)

ctk.CTkLabel(
    header_esquerda,
    text="💧  ControlBombas",
    font=("Arial", 22, "bold"),
    text_color=TEXT,
).pack(anchor="w")

ctk.CTkLabel(
    header_esquerda,
    text="Central de relatórios • Bombas 01, 02 e Poço",
    font=("Arial", 10),
    text_color=MUTED,
).pack(anchor="w")

# ============================================================
# RELÓGIO DA INTERFACE
# Eu mostro no cabeçalho a hora que o próprio Python está lendo do Windows.
# ============================================================
relogio_label = ctk.CTkLabel(
    header,
    text="--:--:--",
    font=("Consolas", 17, "bold"),
    text_color=TEXT,
)

relogio_label.pack(
    side="right",
    padx=(5, 24)
)

# Eu atualizo o relógio mostrado no cabeçalho a cada segundo usando o próprio agendador do Tkinter.
def atualizar_relogio_tela():
    agora = agora_local()

    relogio_label.configure(
        text=agora.strftime("%H:%M:%S")
    )

    app.after(
        1000,
        atualizar_relogio_tela
    )

atualizar_relogio_tela()

# ============================================================
# ÁREA PRINCIPAL
# Eu separo a interface em abas para deixar o relatório e o diagnóstico de horário independentes.
# ============================================================
tabview = ctk.CTkTabview(
    app,
    fg_color="transparent",
    segmented_button_fg_color=CARD,
    segmented_button_selected_color=BLUE,
    segmented_button_selected_hover_color=BLUE_HOVER,
    segmented_button_unselected_color=CARD,
    segmented_button_unselected_hover_color=CARD_2,
)

tabview.pack(
    padx=10,
    pady=5,
    fill="both",
    expand=True
)

tab_relatorio = tabview.add(
    "📋  RELATÓRIO"
)

tab_diagnostico = tabview.add(
    "🕐  DIAGNÓSTICO DO HORÁRIO"
)

# ============================================================
# ABA DE RELATÓRIO
# Nesta aba ficam as escolhas do porteiro, nível da água, turno e geração do relatório.
# ============================================================
scroll_relatorio = ctk.CTkScrollableFrame(
    tab_relatorio,
    fg_color="transparent"
)

scroll_relatorio.pack(
    fill="both",
    expand=True,
    padx=5,
    pady=5
)

card_config = ctk.CTkFrame(
    scroll_relatorio,
    fg_color=CARD,
    corner_radius=15,
    border_width=1,
    border_color=BORDER,
)

card_config.pack(
    fill="x",
    pady=3
)

ctk.CTkLabel(
    card_config,
    text="👮  Porteiro de plantão",
    font=("Arial", 15, "bold"),
).pack(
    pady=(16, 2)
)

frame_diarista = ctk.CTkFrame(
    card_config,
    fg_color="transparent"
)

entry_diarista = ctk.CTkEntry(
    frame_diarista,
    placeholder_text="Nome do diarista...",
    width=240,
    height=34,
)

entry_diarista.pack(
    pady=5
)

# Quando a opção Diarista é escolhida, eu mostro um campo extra para informar o nome.
def check_diarista(selecionado):
    if selecionado == "Diarista":
        frame_diarista.pack(
            pady=2
        )
    else:
        frame_diarista.pack_forget()

porteiro_group = ButtonGroup(
    card_config,
    [
        "Danyel",
        "Leandro",
        "Alisson",
        "Gustavo",
        "Diarista",
    ],
    callback=check_diarista,
)

porteiro_group.select("Danyel")

ctk.CTkLabel(
    card_config,
    text="💧  Nível da água",
    font=("Arial", 15, "bold"),
).pack(
    pady=(15, 2)
)

nivel_group = ButtonGroup(
    card_config,
    [
        "25%",
        "50%",
        "75%",
        "100%",
    ]
)

nivel_group.select("100%")

ctk.CTkLabel(
    card_config,
    text="🕘  Qual turno deseja gerar?",
    font=("Arial", 15, "bold"),
).pack(
    pady=(15, 2)
)

label_turno_escolhido = ctk.CTkLabel(
    card_config,
    text="",
    font=("Arial", 10),
    text_color=MUTED,
)

# Aqui eu calculo o período correspondente ao botão escolhido e mostro uma prévia antes de gerar o relatório.
def atualizar_previa_turno(selecionado):
    mapa = {
        "Atual": 0,
        "Anterior": 1,
        "2 atrás": 2,
        "3 atrás": 3,
    }

    deslocamento = mapa.get(selecionado, 0)
    inicio, fim, nome = obter_turno_alvo(deslocamento)

    label_turno_escolhido.configure(
        text=(
            f"{nome}  •  "
            f"{inicio.strftime('%d/%m %H:%M')} → "
            f"{fim.strftime('%d/%m %H:%M')}"
        )
    )

turnos_group = ButtonGroup(
    card_config,
    [
        "Atual",
        "Anterior",
        "2 atrás",
        "3 atrás",
    ],
    callback=atualizar_previa_turno,
)

label_turno_escolhido.pack(
    pady=(3, 18)
)

turnos_group.select("Atual")

card_whatsapp = ctk.CTkFrame(
    scroll_relatorio,
    fg_color=CARD,
    corner_radius=15,
    border_width=1,
    border_color=BORDER,
)

card_whatsapp.pack(
    fill="x",
    pady=3
)

ctk.CTkLabel(
    card_whatsapp,
    text="💬  WhatsApp",
    font=("Arial", 15, "bold"),
).pack(
    pady=(15, 2)
)

ctk.CTkLabel(
    card_whatsapp,
    text=(
        "Mantenha o WhatsApp Web aberto no Microsoft Edge "
        "para a automação visual funcionar."
    ),
    font=("Arial", 10),
    text_color=MUTED,
).pack(
    pady=(0, 8)
)

ctk.CTkButton(
    card_whatsapp,
    text="ABRIR WHATSAPP WEB",
    command=abrir_whatsapp,
    width=230,
    height=34,
    corner_radius=19,
    fg_color=BLUE,
    hover_color=BLUE_HOVER,
).pack(
    pady=(0, 15)
)

# ============================================================
# GERAÇÃO DO RELATÓRIO
# Aqui acontece o fluxo principal que transforma os registros da planilha no relatório do turno.
# ============================================================
# Nesta função eu junto as escolhas da interface, leio a planilha, filtro os equipamentos do turno e preparo o relatório final.
def gerar_relatorio_turno():
    """
    Gera UM turno por vez.

    Isso permite, por exemplo:
    - 08:00 da manhã + "Anterior" = relatório da noite anterior,
      das 19:00 até 07:00.
    - 20:00 + "Anterior" = relatório do turno diurno,
      das 07:00 até 19:00.

    A geração não consulta internet nem W32Time de forma síncrona.
    O diagnóstico continua disponível na aba própria, evitando que
    o botão GERAR RELATÓRIO pareça travar por causa de rede/serviço.
    """
    agora = agora_local()

    mapa_turnos = {
        "Atual": 0,
        "Anterior": 1,
        "2 atrás": 2,
        "3 atrás": 3,
    }

    escolha_turno = (
        turnos_group.selected_value
        or "Atual"
    )

    deslocamento = mapa_turnos.get(
        escolha_turno,
        0
    )

    data_inicio, data_fim, turno_nome = obter_turno_alvo(
        deslocamento,
        agora
    )

    registrar_acao(
        f"Turno escolhido: {escolha_turno} | "
        f"{data_inicio.strftime('%d/%m/%Y %H:%M')} até "
        f"{data_fim.strftime('%d/%m/%Y %H:%M')}."
    )

    if not CAMINHO_EXCEL.exists():
        messagebox.showerror(
            "Excel não encontrado",
            (
                "Não encontrei o arquivo:\n\n"
                f"{CAMINHO_EXCEL}\n\n"
                "Confira o caminho configurado."
            ),
        )
        return

    try:
        # Aqui eu abro a planilha com o openpyxl para ler os registros já existentes.
        wb = openpyxl.load_workbook(
            CAMINHO_EXCEL,
            data_only=True
        )

        # Depois de abrir o arquivo, eu trabalho com a planilha que está ativa.
        ws = wb.active

    except Exception as erro:
        messagebox.showerror(
            "Erro de leitura",
            (
                "Não consegui abrir o Excel.\n\n"
                "Certifique-se de que o arquivo "
                "não está travado.\n\n"
                f"Detalhes: {erro}"
            ),
        )
        return

    ev_b1 = extrair_linhas_bomba(
        ws,
        "Bomba 01",
        2,
        3,
        5,
        data_inicio,
        data_fim,
        col_porteiro=6,
    )

    ev_b2 = extrair_linhas_bomba(
        ws,
        "Bomba 02",
        8,
        9,
        11,
        data_inicio,
        data_fim,
        col_porteiro=12,
    )

    ev_poco = extrair_linhas_bomba(
        ws,
        "Poço",
        14,
        15,
        16,
        data_inicio,
        data_fim,
        col_porteiro=None,
    )

    eventos_com_nome = [
        evento
        for evento in (ev_b1 + ev_b2)
        if evento.get("porteiro")
        and evento["porteiro"] != "NÃO INFORMADO"
    ]

    for evento_poco in ev_poco:
        candidatos = []

        for referencia in eventos_com_nome:
            ref_inicio = (
                referencia.get("dt_ligou_full")
                or referencia.get("dt_desligou_full")
            )

            ref_fim = (
                referencia.get("dt_desligou_full")
                or referencia.get("dt_ligou_full")
            )

            poco_inicio = (
                evento_poco.get("dt_ligou_full")
                or evento_poco.get("dt_desligou_full")
            )

            poco_fim = (
                evento_poco.get("dt_desligou_full")
                or evento_poco.get("dt_ligou_full")
            )

            if not ref_inicio or not poco_inicio:
                continue

            sobrepoe = (
                ref_inicio <= poco_fim
                and ref_fim >= poco_inicio
            )

            distancia = abs(
                (
                    ref_inicio
                    - poco_inicio
                ).total_seconds()
            )

            candidatos.append(
                (
                    0 if sobrepoe else 1,
                    distancia,
                    referencia["porteiro"],
                )
            )

        if candidatos:
            candidatos.sort(
                key=lambda item: (
                    item[0],
                    item[1]
                )
            )

            evento_poco["porteiro"] = (
                candidatos[0][2]
            )

        else:
            evento_poco["porteiro"] = (
                "NÃO INFORMADO"
            )

    todos_eventos = (
        ev_b1
        + ev_b2
        + ev_poco
    )

    porteiro_atual = (
        porteiro_group.selected_value
    )

    if porteiro_atual == "Diarista":
        nome = entry_diarista.get().strip()

        porteiro_atual = (
            f"DIARISTA ({nome.upper()})"
            if nome
            else "DIARISTA"
        )

    else:
        porteiro_atual = (
            porteiro_atual.upper()
        )

    nivel_atual = (
        nivel_group.selected_value
    )

    txt_turnos = f"TURNO {turno_nome}"

    data_str_cabecalho = data_inicio.strftime(
        "%d/%m/%Y"
    )

    mensagem = (
        f"📋 *RELATÓRIO DE BOMBAS - {txt_turnos}*\n"
        f"📅 Período: {data_str_cabecalho}\n"
        f"👤 Porteiro que enviou: {porteiro_atual}\n"
        f"💧 Nível d'água final: {nivel_atual}\n\n"
        f"{formatar_relatorio_cronologico(todos_eventos)}"
    )

    registrar_acao(
        f"Relatório gerado: {escolha_turno}, "
        f"{len(todos_eventos)} evento(s)."
    )

    mostrar_relatorio(
        "Relatório do Turno",
        mensagem
    )

ctk.CTkButton(
    scroll_relatorio,
    text="📋  GERAR RELATÓRIO",
    command=gerar_relatorio_turno,
    height=50,
    corner_radius=29,
    fg_color=GREEN,
    hover_color=GREEN_HOVER,
    font=("Arial", 15, "bold"),
).pack(
    fill="x",
    pady=(18, 25),
    padx=20
)

# ============================================================
# ABA DE DIAGNÓSTICO
# Esta aba existe para verificar se o relógio do computador está confiável antes de usar os horários no relatório.
# ============================================================
diag_frame = ctk.CTkScrollableFrame(
    tab_diagnostico,
    fg_color="transparent"
)

diag_frame.pack(
    fill="both",
    expand=True,
    padx=5,
    pady=5
)

ctk.CTkLabel(
    diag_frame,
    text="🕐 Saúde do relógio do computador",
    font=("Arial", 21, "bold"),
    text_color=TEXT,
).pack(
    anchor="w",
    padx=12,
    pady=(8, 2)
)

ctk.CTkLabel(
    diag_frame,
    text=(
        "Esta tela existe justamente para proteger o ControlBombas "
        "de um relógio incorreto no computador."
    ),
    font=("Arial", 12),
    text_color=MUTED,
).pack(
    anchor="w",
    padx=12,
    pady=(0, 15)
)

card_status = ctk.CTkFrame(
    diag_frame,
    fg_color=CARD,
    corner_radius=15,
    border_width=1,
    border_color=BORDER,
)

card_status.pack(
    fill="x",
    pady=3
)

status_w32_label = ctk.CTkLabel(
    card_status,
    text="W32Time: aguardando diagnóstico",
    font=("Arial", 16, "bold"),
)

status_w32_label.pack(
    anchor="w",
    padx=20,
    pady=(18, 5)
)

status_fonte_label = ctk.CTkLabel(
    card_status,
    text="Fonte NTP: —",
    font=("Arial", 12),
    text_color=MUTED,
)

status_fonte_label.pack(
    anchor="w",
    padx=20,
    pady=3
)

status_sync_label = ctk.CTkLabel(
    card_status,
    text="Última sincronização: —",
    font=("Arial", 12),
    text_color=MUTED,
)

status_sync_label.pack(
    anchor="w",
    padx=20,
    pady=(3, 18)
)

card_comparacao = ctk.CTkFrame(
    diag_frame,
    fg_color=CARD,
    corner_radius=15,
    border_width=1,
    border_color=BORDER,
)

card_comparacao.pack(
    fill="x",
    pady=3
)

sistema_label = ctk.CTkLabel(
    card_comparacao,
    text="🖥️  Horário do PC: —",
    font=("Arial", 16, "bold"),
)

sistema_label.pack(
    anchor="w",
    padx=20,
    pady=(18, 5)
)

referencia_label = ctk.CTkLabel(
    card_comparacao,
    text="🌐  Horário de referência: —",
    font=("Arial", 16, "bold"),
)

referencia_label.pack(
    anchor="w",
    padx=20,
    pady=5
)

diferenca_label = ctk.CTkLabel(
    card_comparacao,
    text="Diferença: —",
    font=("Arial", 15, "bold"),
)

diferenca_label.pack(
    anchor="w",
    padx=20,
    pady=5
)

fonte_web_label = ctk.CTkLabel(
    card_comparacao,
    text="Fonte: —",
    font=("Arial", 10),
    text_color=MUTED,
)

fonte_web_label.pack(
    anchor="w",
    padx=20,
    pady=(5, 18)
)

diag_botoes = ctk.CTkFrame(
    diag_frame,
    fg_color="transparent"
)

diag_botoes.pack(
    fill="x",
    pady=12
)

# Aqui eu reúno os testes do relógio do Windows e atualizo a aba de diagnóstico com os resultados.
def executar_diagnostico():
    status_w32 = obter_status_w32time()
    ntp = obter_status_ntp()
    fonte = obter_fonte_ntp()
    # Antes de confiar nos horários, eu comparo o relógio do computador com a referência externa.
    comparacao = comparar_horario_sistema()

    if status_w32["status"] == "RUNNING":
        status_w32_label.configure(
            text="🟢 W32Time: EM EXECUÇÃO",
            text_color=GREEN,
        )

    elif status_w32["status"] == "STOPPED":
        status_w32_label.configure(
            text="🔴 W32Time: PARADO",
            text_color=RED,
        )

    else:
        status_w32_label.configure(
            text="🟡 W32Time: DESCONHECIDO",
            text_color=YELLOW,
        )

    fonte_fim = (
        fonte["fonte"]
        if fonte["ok"]
        else "Indisponível"
    )

    status_fonte_label.configure(
        text=f"Fonte NTP: {fonte_fim}"
    )

    ultima_sync = "Indisponível"

    if ntp["ok"]:
        for linha in ntp["texto"].splitlines():
            if (
                "Última Sincronização" in linha
                or "Last Successful Sync Time" in linha
            ):
                ultima_sync = linha.strip()
                break

    status_sync_label.configure(
        text=f"Última sincronização: {ultima_sync}"
    )

    sistema_label.configure(
        text=(
            "🖥️  Horário do PC: "
            + comparacao["sistema"].strftime(
                "%d/%m/%Y %H:%M:%S"
            )
        )
    )

    if comparacao["ok"]:
        referencia_label.configure(
            text=(
                "🌐  Horário de referência: "
                + comparacao["referencia"].strftime(
                    "%d/%m/%Y %H:%M:%S"
                )
            )
        )

        segundos = comparacao["diferenca"]

        # Se a diferença estiver dentro do limite configurado, eu considero o relógio seguro para uso.
        if abs(segundos) <= DIFERENCA_MAXIMA_ALERTA:
            diferenca_label.configure(
                text=(
                    f"🟢 Diferença: "
                    f"{segundos:+.1f} segundo(s) — OK"
                ),
                text_color=GREEN,
            )

        else:
            diferenca_label.configure(
                text=(
                    f"🔴 Diferença: "
                    f"{segundos / 60:+.1f} minuto(s) — ATENÇÃO"
                ),
                text_color=RED,
            )

        fonte_web_label.configure(
            text=(
                f"Fonte: {comparacao['fonte']}  •  "
                f"Latência aproximada: "
                f"{comparacao['latencia']:.2f}s"
            )
        )

        registrar_acao(
            "Diagnóstico do relógio executado."
        )

    else:
        referencia_label.configure(
            text="🌐  Horário de referência: indisponível"
        )

        diferenca_label.configure(
            text="🟡 Diferença: não foi possível comparar",
            text_color=YELLOW,
        )

        fonte_web_label.configure(
            text=(
                "Não foi possível consultar as fontes "
                "de horário pela internet."
            )
        )

        registrar_acao(
            "Diagnóstico: fontes de horário web indisponíveis."
        )

# Esta função é chamada pelo botão de sincronização e mostra ao usuário se o comando foi aceito pelo Windows.
def acao_sincronizar():
    ok, mensagem = sincronizar_w32time()

    if ok:
        messagebox.showinfo(
            "Sincronização",
            (
                "O Windows recebeu o comando para "
                "redescobrir a fonte e sincronizar o horário.\n\n"
                "Agora execute o diagnóstico novamente."
            ),
        )

        executar_diagnostico()

    else:
        messagebox.showerror(
            "Erro na sincronização",
            (
                "Não foi possível sincronizar.\n\n"
                f"{mensagem}\n\n"
                "Se o Windows solicitar privilégios, "
                "abra o ControlBombas como administrador."
            ),
        )

ctk.CTkButton(
    diag_botoes,
    text="🔎  VERIFICAR AGORA",
    command=executar_diagnostico,
    height=45,
    fg_color=BLUE,
    hover_color=BLUE_HOVER,
    font=("Arial", 13, "bold"),
).pack(
    side="left",
    expand=True,
    fill="x",
    padx=5
)

ctk.CTkButton(
    diag_botoes,
    text="🔄  SINCRONIZAR WINDOWS",
    command=acao_sincronizar,
    height=45,
    fg_color=GREEN,
    hover_color=GREEN_HOVER,
    font=("Arial", 13, "bold"),
).pack(
    side="left",
    expand=True,
    fill="x",
    padx=5
)

card_tecnico = ctk.CTkFrame(
    diag_frame,
    fg_color=CARD,
    corner_radius=15,
    border_width=1,
    border_color=BORDER,
)

card_tecnico.pack(
    fill="x",
    pady=5
)

ctk.CTkLabel(
    card_tecnico,
    text="🛡️  Proteção do ControlBombas",
    font=("Arial", 15, "bold"),
).pack(
    anchor="w",
    padx=20,
    pady=(18, 6)
)

ctk.CTkLabel(
    card_tecnico,
    text=(
        "Antes de gerar um relatório, o programa tenta comparar "
        "o relógio do Windows com uma referência de horário pela "
        "internet. Se a diferença for maior que 2 minutos, o "
        "relatório é bloqueado para evitar horários incorretos.\n\n"
        "O Python não altera o relógio do computador. Ele apenas "
        "lê a hora do Windows. Quando você usa 'Sincronizar Windows', "
        "o comando é enviado ao W32Time, que é o serviço responsável "
        "pela sincronização do sistema."
    ),
    justify="left",
    anchor="w",
    wraplength=820,
    font=("Arial", 10),
    text_color=MUTED,
).pack(
    anchor="w",
    padx=20,
    pady=(0, 18)
)

# Pouco depois de abrir o programa, eu executo automaticamente um primeiro diagnóstico do horário.
def diagnostico_inicial():
    try:
        executar_diagnostico()
    except Exception as erro:
        registrar_acao(
            f"Erro no diagnóstico inicial: {erro}"
        )

app.after(
    500,
    diagnostico_inicial
)

registrar_acao(
    "ControlBombas V2 iniciado."
)

app.mainloop()
