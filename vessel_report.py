import anthropic
import imaplib
import smtplib
import email
import email.header
import os
import io
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta

import openpyxl

GMAIL_USER = os.environ.get("GMAIL_USER", "joao@lyrashipping.com.br")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
REPORT_RECIPIENT = os.environ.get("REPORT_RECIPIENT", GMAIL_USER)
CRISTIANO_EMAIL = "cristiano@lyrashipping.com.br"

FLEET = ["EKATERINA", "CALLIO", "DAHLIA", "PARAGON", "HERAKLITOS", "MARCOS DIAS", "LEFTERIS T"]


def previous_business_day():
    today = datetime.now()
    offset = 1
    if today.weekday() == 0:   # Monday -> Friday
        offset = 3
    elif today.weekday() == 6: # Sunday -> Friday
        offset = 2
    return today - timedelta(days=offset)


def decode_header_value(value):
    parts = email.header.decode_header(value)
    decoded = []
    for part, charset in parts:
        if isinstance(part, bytes):
            decoded.append(part.decode(charset or "utf-8", errors="ignore"))
        else:
            decoded.append(part)
    return "".join(decoded)


def excel_to_text(data: bytes) -> str:
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    lines = []
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        lines.append(f"[Aba: {sheet}]")
        for row in ws.iter_rows(values_only=True):
            row_values = [str(v) if v is not None else "" for v in row]
            if any(v.strip() for v in row_values):
                lines.append("\t".join(row_values))
    return "\n".join(lines)


def fetch_cristiano_positions(mail_conn):
    """Busca o Excel de posições do Cristiano do dia útil anterior."""
    pbd = previous_business_day()
    since = pbd.strftime("%d-%b-%Y")

    _, ids = mail_conn.search(None, f'(FROM "{CRISTIANO_EMAIL}" SINCE {since})')
    if not ids[0]:
        return None, None

    # varre todos os emails do mais recente até achar um com Excel
    for msg_id in reversed(ids[0].split()):
        _, msg_data = mail_conn.fetch(msg_id, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])
        subject = decode_header_value(msg.get("Subject", ""))
        date_str = msg.get("Date", "")

        for part in msg.walk():
            fn = part.get_filename()
            if fn:
                fn_decoded = decode_header_value(fn)
                ext = fn_decoded.lower().split(".")[-1]
                if ext in ("xlsx", "xls"):
                    payload = part.get_payload(decode=True)
                    if payload:
                        text = excel_to_text(payload)
                        return text, f"De: {CRISTIANO_EMAIL} | Data: {date_str} | Assunto: {subject}"
    return None, None


def fetch_vessel_emails():
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_USER, GMAIL_APP_PASSWORD)
    mail.select("inbox")

    positions_text, positions_meta = fetch_cristiano_positions(mail)

    since_date = (datetime.now() - timedelta(days=1)).strftime("%d-%b-%Y")
    _, message_ids = mail.search(None, f"(SINCE {since_date})")

    emails = []
    ids = message_ids[0].split()
    for msg_id in ids[-50:]:
        _, msg_data = mail.fetch(msg_id, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])

        subject = decode_header_value(msg.get("Subject", "(sem assunto)"))
        sender = decode_header_value(msg.get("From", ""))
        date_str = msg.get("Date", "")
        body = ""

        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    payload = part.get_payload(decode=True)
                    if payload:
                        body = payload.decode("utf-8", errors="ignore")
                        break
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                body = payload.decode("utf-8", errors="ignore")

        body = body[:3000]
        emails.append({"subject": subject, "from": sender, "date": date_str, "body": body})

    mail.close()
    mail.logout()
    return emails, positions_text, positions_meta


def generate_report(emails, positions_text, positions_meta):
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    today = datetime.now().strftime("%d/%m/%Y")
    fleet_list = ", ".join(FLEET)

    email_content = "\n\n---\n\n".join([
        f"De: {e['from']}\nData: {e['date']}\nAssunto: {e['subject']}\n\n{e['body']}"
        for e in emails
    ]) if emails else "Nenhum email recebido nas últimas 24 horas."

    if positions_text:
        positions_section = f"""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PLANILHA DE POSIÇÕES (Cristiano — dia útil anterior)
{positions_meta}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{positions_text[:4000]}

Com base nessa planilha:
- Identifique quais navios da frota estão em viagem ativa (em navegação, em porto, fundeados)
- Identifique quais navios possivelmente estão sem viagem ativa (laid up, sem rota, etc.)
- Use essa informação para enriquecer o status de cada navio no relatório
- Se um navio da frota não aparecer na planilha, indique "não consta na planilha de posições"
"""
    else:
        positions_section = "\nPlanilha de posições do Cristiano: não encontrada para o dia útil anterior.\n"

    prompt = f"""Voce e um assistente especializado em operacoes maritimas da empresa Lyra Shipping.

REGRAS:
- NAO use markdown (sem ##, sem **, sem _).
- Use APENAS texto simples com os separadores abaixo.
- Seja conciso. Sem frases longas. Prefira bullet points curtos.

A frota tem EXATAMENTE estes 7 navios — inclua TODOS, sem excecao:
{fleet_list}

{positions_section}

Gere o relatorio usando EXATAMENTE este formato:

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
LYRA SHIPPING  ·  RELATORIO DIARIO  ·  {today}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[Para cada navio, use este bloco:]

[NIVEL]  M/V [NOME]
  Status    [navegando / em porto / fundeado / retido / inativo]
  Rota      [Porto de origem] -> [Porto de destino]  |  ETA [data hora]
  Viagem    [VOY]  |  [Carga e quantidade]

  URGENTE                          <- so se houver algo critico
  - [item critico curto e claro, sem siglas desnecessarias]

  OPERACIONAL
  - [apenas o mais relevante: posicao, ROB, status da operacao, agente, PDA]
  - maximo 4 bullets

  PROXIMA VIAGEM
  - Carga: [tipo e quantidade]  |  Cliente: [se disponivel]
  - [Porto de carga] -> [Porto de descarga]
  (omitir se nao houver informacao)

  PROXIMOS PASSOS
  - [acao necessaria, clara e direta]

──────────────────────────────────────────────

REGRAS DE ESCRITA:
- Escreva os nomes dos portos por extenso, sem siglas (ex: Santos, nao STS).
- Evite siglas tecnicas que um leitor nao maritimo nao entenderia. Quando usar, explique brevemente (ex: ETB = previsao de atracacao).
- Nao liste todas as escalas futuras — apenas a proxima viagem imediata.
- Seja direto. Frases curtas. Sem repeticoes.

[NIVEL]:
  🔴  retencao, danos, disputa ativa, prazo critico
  🟡  pendencia financeira, ETB incerto, decisao pendente
  🟢  operacao normal sem pendencias
  ⚫  sem viagem ativa ou sem informacao

ORDENACAO: 🔴 primeiro, 🟢 por ultimo.
Se nao houver URGENTE, omita essa secao inteira.

---
Emails das ultimas 24 horas:
{email_content}"""

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=8000,
        messages=[{"role": "user", "content": prompt}],
    )

    return response.content[0].text


def send_report(report_text):
    today = datetime.now().strftime("%d/%m/%Y")
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Relatorio Diario de Navios — {today}"
    msg["From"] = GMAIL_USER
    msg["To"] = REPORT_RECIPIENT

    html = f"<html><body><pre style='font-family:monospace;font-size:14px'>{report_text}</pre></body></html>"
    msg.attach(MIMEText(report_text, "plain", "utf-8"))
    msg.attach(MIMEText(html, "html", "utf-8"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_USER, REPORT_RECIPIENT, msg.as_string())

    print(f"[OK] Relatorio enviado para {REPORT_RECIPIENT}")


if __name__ == "__main__":
    print(f"[{datetime.now().strftime('%d/%m/%Y %H:%M')}] Iniciando geracao do relatorio diario...")

    if not GMAIL_APP_PASSWORD:
        print("[ERRO] GMAIL_APP_PASSWORD nao configurado.")
        exit(1)
    if not ANTHROPIC_API_KEY:
        print("[ERRO] ANTHROPIC_API_KEY nao configurado.")
        exit(1)

    emails, positions_text, positions_meta = fetch_vessel_emails()
    print(f"[INFO] {len(emails)} email(s) encontrado(s) nas ultimas 24h.")
    if positions_text:
        print(f"[INFO] Planilha de posicoes encontrada: {positions_meta}")
    else:
        print("[AVISO] Planilha de posicoes do Cristiano nao encontrada.")

    report = generate_report(emails, positions_text, positions_meta)
    print("\n--- RELATORIO GERADO ---")
    print(report.encode("cp1252", errors="replace").decode("cp1252"))
    print("------------------------\n")

    send_report(report)
