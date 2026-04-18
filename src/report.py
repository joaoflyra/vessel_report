import os
import imaplib
import smtplib
import email as email_lib
from datetime import datetime, timezone, timedelta
from email.mime.text import MIMEText
import anthropic

GMAIL_USER = os.environ["GMAIL_USER"]
GMAIL_TO = os.environ["GMAIL_TO"]
GMAIL_APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"].replace(" ", "").replace("\xa0", "").strip()
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]

BRASILIA = timezone(timedelta(hours=-3))
TODAY_BR = datetime.now(BRASILIA).strftime("%d/%m/%Y")
TODAY_IMAP = datetime.now(BRASILIA).strftime("%d-%b-%Y")

VESSELS = ["HERAKLITOS", "PARAGON", "EKATERINA", "LEFTERIS T", "DAHLIA", "MARCOS DIAS", "CALLIO"]


def fetch_emails():
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_USER, GMAIL_APP_PASSWORD)
    mail.select("inbox")
    emails_by_vessel = {}
    for vessel in VESSELS:
        ids = set()
        for criteria in [f'SUBJECT "{vessel}"', f'BODY "{vessel}"']:
            _, data = mail.search(None, f'(SINCE "{TODAY_IMAP}" {criteria})')
            ids.update(data[0].split())
        vessel_emails = []
        for num in list(ids)[:8]:
            _, msg_data = mail.fetch(num, "(RFC822)")
            msg = email_lib.message_from_bytes(msg_data[0][1])
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode("utf-8", errors="ignore")[:3000]
                        break
            else:
                body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")[:3000]
            vessel_emails.append({"subject": str(msg["subject"] or ""), "from": str(msg["from"] or ""), "body": body})
        emails_by_vessel[vessel] = vessel_emails
        print(f"{vessel}: {len(vessel_emails)} emails encontrados")
    mail.logout()
    return emails_by_vessel


def generate_report(emails_by_vessel):
    emails_text = ""
    for vessel, emails in emails_by_vessel.items():
        emails_text += f"\n\n=== {vessel} ===\n"
        if not emails:
            emails_text += "Sem emails hoje.\n"
        for e in emails:
            emails_text += f"\nAssunto: {e['subject']}\nDe: {e['from']}\n{e['body']}\n---\n"
    prompt = f"""Voce e um assistente de operacoes maritimas da Lyra Shipping.
Com base nos emails de hoje ({TODAY_BR}), gere o relatorio diario.

EMAILS DE HOJE:
{emails_text}

FORMATO:
LYRA SHIPPING . RELATORIO DIARIO . {TODAY_BR}

[EMOJI] M/V [NOME]
  Status    [descricao]
  Rota      [origem -> destino] | [viagem] | [carga]

  URGENTE (somente se necessario)
  - [item]

  OPERACIONAL
  - [posicao, bunkers, eventos]

  PROXIMA VIAGEM
  - [detalhes]

  PROXIMOS PASSOS
  - [acoes]

Emoji: vermelho critico, amarelo pendencias, verde normal. Ordem: critico primeiro, pendencias segundo, normal terceiro.
Se sem emails: indique "Sem atualizacoes recebidas hoje"."""
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        messages=[{"role": "user", "content": prompt}]
    )
    return message.content[0].text


def send_email(subject, body):
    msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"] = subject
    msg["From"] = GMAIL_USER
    msg["To"] = GMAIL_TO
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_USER, GMAIL_TO, msg.as_string())
    print(f"Email enviado para {GMAIL_TO}")


if __name__ == "__main__":
    emails_by_vessel = fetch_emails()
    report = generate_report(emails_by_vessel)
    subject = f"LYRA SHIPPING . RELATORIO DIARIO . {TODAY_BR}"
    send_email(subject, report)
