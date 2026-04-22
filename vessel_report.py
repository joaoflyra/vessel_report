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
    if today.weekday() == 0:
        offset = 3
    elif today.weekday() == 6:
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
    pbd = previous_business_day()
    since = pbd.strftime("%d-%b-%Y")

    _, ids = mail_conn.search(None, f'(FROM "{CRISTIANO_EMAIL}" SINCE {since})')
    if not ids[0]:
        return None, None

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


def fetch_last_report(mail_conn):
    """Busca o último relatório diário enviado para contextualização."""
    since = (datetime.now() - timedelta(days=7)).strftime("%d-%b-%Y")

    sent_folders = ['"[Gmail]/Sent Mail"', '"[Gmail]/Enviados"', "Sent", "Enviados"]
    selected = False
    for folder in sent_folders:
        try:
            status, _ = mail_conn.select(folder)
            if status == "OK":
                selected = True
                break
        except Exception:
            continue

    if not selected:
        return None, None

    _, ids = mail_conn.search(None, f'(SUBJECT "Relatorio Diario" SINCE {since})')
    mail_conn.select("inbox")

    if not ids[0]:
        return None, None

    msg_id = ids[0].split()[-1]

    # re-seleciona a pasta enviados para buscar o email
    for folder in sent_folders:
        try:
            status, _ = mail_conn.select(folder)
            if status == "OK":
                break
        except Exception:
            continue

    _, msg_data = mail_conn.fetch(msg_id, "(RFC822)")
    msg = email.message_from_bytes(msg_data[0][1])
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

    mail_conn.select("inbox")
    return body[:5000], date_str


def fetch_vessel_emails():
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_USER, GMAIL_APP_PASSWORD)
    mail.select("inbox")

    positions_text, positions_meta = fetch_cristiano_positions(mail)
    last_report, last_report_date = fetch_last_report(mail)

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
    return emails, positions_text, positions_meta, last_report, last_report_date


def generate_report(emails, positions_text, positions_meta, last_report, last_report_date):
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    today = datetime.now().strftime("%d/%m/%Y")
    fleet_list = ", ".join(FLEET)

    email_content = "\n\n---\n\n".join([
        f"De: {e['from']}\nData: {e['date']}\nAssunto: {e['subject']}\n\n{e['body']}"
        for e in emails
    ]) if emails else "Nenhum email recebido nas ultimas 24 horas."

    if positions_text:
        positions_section = f"""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PLANILHA DE POSICOES (Cristiano — dia util anterior)
{positions_meta}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{positions_text[:4000]}
"""
    else:
        positions_section = "\nPlanilha de posicoes do Cristiano: nao encontrada para o dia util anterior.\n"

    if last_report:
        last_report_section = f"""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ULTIMO RELATORIO ENVIADO (para contextualizacao)
Data: {last_report_date}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{last_report}
"""
    else:
        last_report_section = "\nUltimo relatorio: nao encontrado.\n"

    prompt = f"""Voce e um assistente especializado em operacoes maritimas da empresa Lyra Shipping.

REGRA FUNDAMENTAL — SEM DEDUCOES:
- Inclua APENAS informacoes confirmadas explicitamente nos emails ou na planilha de posicoes.
- Se uma informacao nao estiver clara nos dados fornecidos, escreva "sem informacao" — jamais deduza, infira ou suponha.
- Nao complete lacunas com estimativas proprias. Se o ETA nao estiver no email, escreva "ETA nao informado".
- Em caso de duvida sobre qualquer dado, omita-o ou indique "dado nao confirmado".

OUTRAS REGRAS:
- NAO use markdown (sem ##, sem **, sem _).
- Use APENAS texto simples com os separadores abaixo.
- Seja conciso. Sem frases longas. Prefira bullet points curtos.
- Escreva os nomes dos portos por extenso, sem siglas.
- Evite siglas tecnicas sem explicacao.
- Nao liste escalas futuras — apenas a proxima viagem imediata.
- Nao repita informacoes do ultimo relatorio se nao houver atualizacao confirmada.

A frota tem EXATAMENTE estes 7 navios — inclua TODOS, sem excecao:
{fleet_list}

{positions_section}

{last_report_section}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
FORMATO DO RELATORIO — use EXATAMENTE:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
LYRA SHIPPING  ·  RELATORIO DIARIO  ·  {today}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

[Para cada navio:]

[NIVEL]  M/V [NOME]
  Status    [navegando / em porto / fundeado / retido / inativo / sem informacao]
  Rota      [Porto de origem] -> [Porto de destino]  |  ETA [data hora ou "nao informado"]
  Viagem    [VOY]  |  [Carga e quantidade]

  URGENTE                          <- so se houver algo critico confirmado
  - [item critico curto e claro]

  OPERACIONAL
  - [apenas o confirmado nos emails: posicao, ROB, status da operacao, agente, PDA]
  - maximo 4 bullets
  - se nao houver informacao nova confirmada, escreva "sem atualizacao confirmada"

  PROXIMA VIAGEM
  - Tipo: [Voyage Charter / Time Charter]  |  Cliente: [se disponivel]
  - Carga: [tipo e quantidade, se voyage charter]
  - [Porto de carga] -> [Porto de descarga]
  (omitir se nao houver informacao confirmada)

  PROXIMOS PASSOS
  - [acao necessaria confirmada pelos emails]

──────────────────────────────────────────────

[NIVEL]:
  🔴  retencao, danos, disputa ativa, prazo critico
  🟡  pendencia financeira, ETB incerto, decisao pendente
  🟢  operacao normal sem pendencias
  ⚫  sem viagem ativa ou sem informacao

ORDENACAO: 🔴 primeiro, 🟢 por ultimo.
Se nao houver URGENTE, omita essa secao inteira.

Apos o relatorio em Portugues, adicione exatamente este separador e repita o relatorio completo em Ingles:

════════════════════════════════════════════════════════
ENGLISH VERSION  ·  DAILY VESSEL REPORT  ·  {today}
════════════════════════════════════════════════════════

[Same format and same rules in English. Only confirmed data.]

---
Emails das ultimas 24 horas:
{email_content}"""

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=10000,
        messages=[{"role": "user", "content": prompt}],
    )

    return response.content[0].text


def send_report(report_text):
    today = datetime.now().strftime("%d/%m/%Y")
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Relatorio Diario de Navios — {today}"
    msg["From"] = GMAIL_USER
    msg["To"] = REPORT_RECIPIENT
    msg["Cc"] = CRISTIANO_EMAIL

    html = f"<html><body><pre style='font-family:monospace;font-size:14px'>{report_text}</pre></body></html>"
    msg.attach(MIMEText(report_text, "plain", "utf-8"))
    msg.attach(MIMEText(html, "html", "utf-8"))

    recipients = list({REPORT_RECIPIENT, CRISTIANO_EMAIL})
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_USER, recipients, msg.as_string())

    print(f"[OK] Relatorio enviado para {', '.join(recipients)}")


if __name__ == "__main__":
    print(f"[{datetime.now().strftime('%d/%m/%Y %H:%M')}] Iniciando geracao do relatorio diario...")

    if not GMAIL_APP_PASSWORD:
        print("[ERRO] GMAIL_APP_PASSWORD nao configurado.")
        exit(1)
    if not ANTHROPIC_API_KEY:
        print("[ERRO] ANTHROPIC_API_KEY nao configurado.")
        exit(1)

    emails, positions_text, positions_meta, last_report, last_report_date = fetch_vessel_emails()
    print(f"[INFO] {len(emails)} email(s) encontrado(s) nas ultimas 24h.")
    if positions_text:
        print(f"[INFO] Planilha de posicoes encontrada: {positions_meta}")
    else:
        print("[AVISO] Planilha de posicoes do Cristiano nao encontrada.")
    if last_report:
        print(f"[INFO] Ultimo relatorio encontrado: {last_report_date}")
    else:
        print("[AVISO] Ultimo relatorio nao encontrado.")

    report = generate_report(emails, positions_text, positions_meta, last_report, last_report_date)
    print("\n--- RELATORIO GERADO ---")
    print(report.encode("cp1252", errors="replace").decode("cp1252"))
    print("------------------------\n")

    send_report(report)
