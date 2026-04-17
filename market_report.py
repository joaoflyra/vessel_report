import anthropic
import imaplib
import smtplib
import email
import email.header
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta

GMAIL_USER = os.environ.get("GMAIL_USER", "joao@lyrashipping.com.br")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
REPORT_RECIPIENT = os.environ.get("REPORT_RECIPIENT", GMAIL_USER)


def decode_header_value(value):
    parts = email.header.decode_header(value)
    decoded = []
    for part, charset in parts:
        if isinstance(part, bytes):
            decoded.append(part.decode(charset or "utf-8", errors="ignore"))
        else:
            decoded.append(part)
    return "".join(decoded)


def fetch_broker_emails():
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_USER, GMAIL_APP_PASSWORD)
    mail.select("inbox")

    since_date = (datetime.now() - timedelta(days=1)).strftime("%d-%b-%Y")
    _, message_ids = mail.search(None, f"(SINCE {since_date})")

    emails = []
    ids = message_ids[0].split()
    for msg_id in ids:
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

        body = body[:2000]
        emails.append({"subject": subject, "from": sender, "date": date_str, "body": body})

    mail.close()
    mail.logout()
    return emails


def generate_market_report(emails):
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    today = datetime.now().strftime("%d/%m/%Y")

    email_content = "\n\n---\n\n".join([
        f"De: {e['from']}\nAssunto: {e['subject']}\n\n{e['body']}"
        for e in emails
    ]) if emails else "Nenhum email recebido."

    prompt = f"""Você é um analista especializado em mercado de fretes maritimos (dry bulk).

Analise os emails abaixo recebidos nas últimas 24 horas e gere um relatório de mercado.

Primeiro, identifique quais emails são de brokers maritimos. Use critério AMPLO — qualquer email que contenha ofertas de navio, ordens de carga, fixturamentos ou circulares de mercado deve ser considerado broker.

Brokers conhecidos (mas não limitado a estes):
Clarksons, SSY, Fearnleys, Braemar, Howe Robinson, BRS, Banchero Costa, Poten, Barry Rogliano, ACM, Maersk Broker, Simpson Spence Young, Intermodal, Compass Maritime, Ifchor, Galbraiths, DHP Maritime, North Harbour Shipping, Prime Shipping & Chartering (chartering@primeshipping.org), Niavigrains Chartering, Skyhi Shipping, Guanabara Shipping, e qualquer outro remetente com conteúdo típico de circular de mercado maritimo (tonnage, orders, laycan, DWT, L/D rates, etc.).

Em caso de dúvida sobre se um email é de broker, INCLUA — é melhor incluir do que ignorar.

De cada email de broker, extraia APENAS estes 3 tipos de itens:
- Ofertas de navio (tonnage): navio disponivel com porto/data de abertura
- Procuras de navio (orders): cargo buscando navio
- Ofertas de carga (cargo offer): carga disponivel buscando navio

NAO inclua "procuras de carga" (cargo wanted / backhaul). Ignore completamente esses itens.

Classifique cada navio por porte:
- Capesize: 100.000+ DWT
- Panamax / Kamsarmax: 65.000-99.999 DWT
- Supramax / Ultramax: 45.000-64.999 DWT
- Handysize: abaixo de 45.000 DWT

Classifique por região:
- Américas (ECSA = Costa Leste América do Sul, WCSA, USG, USEC)
- Europa / Mediterrâneo
- Ásia (Far East, Southeast Asia)
- Oceania (Austrália, Nova Zelândia)
- Oriente Médio / Índia
- África

Gere o relatório seguindo EXATAMENTE este formato (texto simples, sem markdown):

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
LYRA SHIPPING  ·  RELATORIO DE MERCADO  ·  {today}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

RESUMO EXECUTIVO
  E-mails de brokers analisados: [N]
  Itens individuais extraídos:   [N]
    Ofertas de navio (tonnage):       [N]
    Procuras de navio (orders):       [N]
    Ofertas de carga (cargo offer):   [N]

══════════════════════════════════════════════
OFERTAS DE NAVIO (TONNAGE)
══════════════════════════════════════════════

[Por porte, depois por região, apenas onde houver itens:]

CAPESIZE
  Americas (ECSA / USG)
    - [Nome navio] / [DWT] / [ano construção] — Aberto [porto] [data] — Broker: [nome]
  Asia (Far East)
    - ...

PANAMAX / KAMSARMAX
  ...

SUPRAMAX / ULTRAMAX
  ...

HANDYSIZE
  ...

══════════════════════════════════════════════
PROCURAS DE NAVIO (ORDERS)
══════════════════════════════════════════════

[Por porte, depois por região:]

CAPESIZE
  Americas (ECSA)
    - [quantidade] [tipo de carga] [porto carga]/[porto descarga] — laycan [datas] — Charterer: [se conhecido] — Broker: [nome]
  ...

PANAMAX / KAMSARMAX
  ...

══════════════════════════════════════════════
OFERTAS DE CARGA (CARGO OFFER)
══════════════════════════════════════════════

[Por região de carregamento:]

  Americas (ECSA / USG)
    - [quantidade] [tipo carga] [porto carga]/[porto descarga] — laycan [datas] — Shipper: [se conhecido] — Broker: [nome]
  ...

══════════════════════════════════════════════
DESTAQUES DO DIA
══════════════════════════════════════════════
- [3 a 5 bullets com movimentos mais relevantes: concentração de ordens, niveis de frete mencionados, tendências, padrões]

══════════════════════════════════════════════
AUDITORIA
══════════════════════════════════════════════
  E-mails ignorados (nao-broker): [N]
  E-mails de broker sem item extraível: [N]

Se não houver emails de brokers suficientes, indique claramente e faça um resumo do que foi encontrado.

---
Emails recebidos nas últimas 24 horas ({len(emails)} emails):
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
    msg["Subject"] = f"Relatorio de Mercado — {today}"
    msg["From"] = GMAIL_USER
    msg["To"] = REPORT_RECIPIENT

    html = f"<html><body><pre style='font-family:monospace;font-size:14px'>{report_text}</pre></body></html>"
    msg.attach(MIMEText(report_text, "plain", "utf-8"))
    msg.attach(MIMEText(html, "html", "utf-8"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_USER, REPORT_RECIPIENT, msg.as_string())

    print(f"[OK] Relatorio de mercado enviado para {REPORT_RECIPIENT}")


if __name__ == "__main__":
    print(f"[{datetime.now().strftime('%d/%m/%Y %H:%M')}] Gerando relatorio de mercado...")

    if not GMAIL_APP_PASSWORD:
        print("[ERRO] GMAIL_APP_PASSWORD nao configurado.")
        exit(1)
    if not ANTHROPIC_API_KEY:
        print("[ERRO] ANTHROPIC_API_KEY nao configurado.")
        exit(1)

    emails = fetch_broker_emails()
    print(f"[INFO] {len(emails)} email(s) encontrado(s) nas ultimas 24h.")

    report = generate_market_report(emails)
    print("\n--- RELATORIO GERADO ---")
    print(report.encode("cp1252", errors="replace").decode("cp1252"))
    print("------------------------\n")

    send_report(report)
