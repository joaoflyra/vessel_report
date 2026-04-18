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

FALLBACK_VESSELS = ["HERAKLITOS", "PARAGON", "EKATERINA", "LEFTERIS T", "DAHLIA", "MARCOS DIAS", "CALLIO"]


def fetch_emails_imap():
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_USER, GMAIL_APP_PASSWORD)
    return mail


def get_vessels_from_position_list(mail):
    """Busca a position list mais recente de Cristiano para identificar navios ativos."""
    mail.select("inbox")
    ids = set()
    for criteria in [
        '(FROM "cristiano@lyrashipping.com.br" SUBJECT "posicao")',
        '(FROM "cristiano@lyrashipping.com.br" SUBJECT "position")',
        '(FROM "cristiano@lyrashipping.com.br" SUBJECT "daily")',
    ]:
        _, data = mail.search(None, criteria)
        ids.update(data[0].split())

    if not ids:
        print("Position list nao encontrada, usando lista de fallback")
        return FALLBACK_VESSELS, ""

    # Pega o email mais recente
    latest = sorted(ids)[-1]
    _, msg_data = mail.fetch(latest, "(RFC822)")
    msg = email_lib.message_from_bytes(msg_data[0][1])
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                break
    else:
        body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")

    # Extrai nomes de navios conhecidos que aparecem na position list
    known_vessels = ["HERAKLITOS", "PARAGON", "EKATERINA", "LEFTERIS T", "LEFTERIS",
                     "DAHLIA", "MARCOS DIAS", "CALLIO", "RAYS", "APOGEE", "AGIA ELENI",
                     "SEPETIBA", "VANTAGE ROSE", "CHINTANA", "DISCOVERER", "HANZE",
                     "ATAYAL", "PIO GRANDE", "ADVENTURER", "LOYALTY", "SEAHEAVEN",
                     "LOWLANDS", "REVENGER", "EKATERINA", "PARANA WARRIOR"]
    found = [v for v in known_vessels if v in body.upper()]
    if len(found) < 3:
        print(f"Poucos navios identificados ({found}), usando lista completa de fallback")
        return FALLBACK_VESSELS, body

    print(f"Navios na position list: {found}")
    return found, body


def fetch_all_emails(mail, vessels):
    """Busca emails de hoje sobre cada navio E emails gerais de operacao."""
    mail.select("inbox")
    emails_by_vessel = {}

    for vessel in vessels:
        ids = set()
        for criteria in [f'SUBJECT "{vessel}"', f'BODY "{vessel}"']:
            _, data = mail.search(None, f'(SINCE "{TODAY_IMAP}" {criteria})')
            ids.update(data[0].split())

        vessel_emails = []
        for num in sorted(ids)[-10:]:
            _, msg_data = mail.fetch(num, "(RFC822)")
            msg = email_lib.message_from_bytes(msg_data[0][1])
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode("utf-8", errors="ignore")[:4000]
                        break
            else:
                body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")[:4000]
            vessel_emails.append({
                "subject": str(msg["subject"] or ""),
                "from": str(msg["from"] or ""),
                "date": str(msg["date"] or ""),
                "body": body,
            })
        emails_by_vessel[vessel] = vessel_emails
        print(f"{vessel}: {len(vessel_emails)} emails encontrados")

    # Busca emails de viagens antigas (nao relacionados aos navios ativos de hoje)
    _, data = mail.search(None, f'(SINCE "{TODAY_IMAP}")')
    all_today_ids = set(data[0].split())
    vessel_ids = set()
    for v_emails in emails_by_vessel.values():
        pass  # ids ja processados acima

    return emails_by_vessel


def generate_report(emails_by_vessel, position_list_body=""):
    emails_text = ""
    for vessel, emails in emails_by_vessel.items():
        emails_text += f"\n\n=== {vessel} ===\n"
        if not emails:
            emails_text += "Sem emails hoje.\n"
        for e in emails:
            emails_text += f"\nAssunto: {e['subject']}\nDe: {e['from']}\nData: {e['date']}\n{e['body']}\n---\n"

    position_list_section = ""
    if position_list_body:
        position_list_section = f"\n\nPOSITION LIST MAIS RECENTE:\n{position_list_body[:3000]}"

    prompt = f"""Voce e um assistente experiente de operacoes maritimas da Lyra Shipping (Rio de Janeiro).
Gere o relatorio diario de navios com base nos emails de hoje ({TODAY_BR}).

{position_list_section}

EMAILS DE HOJE POR NAVIO:
{emails_text}

INSTRUCOES:

1. ORDENACAO POR IMPORTANCIA:
   - 🔴 PRIMEIRO: Navios com problemas criticos (retido, reparos obrigatorios, P&I, demurrage correndo, situacoes legais urgentes)
   - 🟡 SEGUNDO: Navios com pendencias importantes (fundeado aguardando berco, descarregando, pendencias financeiras abertas, documentos pendentes, decisoes necessarias)
   - 🟢 TERCEIRO: Navios navegando normalmente ou carregando sem problemas — seja BREVE e direto para esses

2. NIVEL DE DETALHE:
   - Navios criticos (🔴): detalhe completo com todos os pontos urgentes
   - Navios com pendencias (🟡): detalhe relevante, foco nas pendencias
   - Navios tranquilos (🟢): apenas status, rota, ETA e proximos passos essenciais — sem blocos vazios

3. FORMATO POR NAVIO:
[EMOJI]  M/V [NOME]
  Status    [descricao]
  Rota      [origem -> destino]  |  [viagem]  |  [carga]

  URGENTE (somente se houver)
  - [item]

  OPERACIONAL
  - [info relevante]

  PROXIMA VIAGEM (somente se relevante)
  - [detalhes]

  PROXIMOS PASSOS
  - [acoes concretas pendentes]

4. VIAGENS ANTIGAS: Se houver emails sobre viagens ja encerradas (ex: demonstrativos de frete, saldos pendentes de escalas anteriores), liste-os ao final em uma secao separada:

  PENDENCIAS DE VIAGENS ANTERIORES
  - M/V [NOME] voy [X]: [descricao da pendencia]

5. Se um navio nao tiver emails hoje mas estiver na position list, inclua com "Sem atualizacoes recebidas hoje" de forma breve.

Cabecalho do relatorio:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
LYRA SHIPPING  ·  RELATORIO DIARIO  ·  {TODAY_BR}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Separador entre navios: ──────────────────────────────────────"""

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=8096,
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
    mail = fetch_emails_imap()
    vessels, position_list_body = get_vessels_from_position_list(mail)
    emails_by_vessel = fetch_all_emails(mail, vessels)
    mail.logout()
    report = generate_report(emails_by_vessel, position_list_body)
    subject = f"LYRA SHIPPING . RELATORIO DIARIO . {TODAY_BR}"
    send_email(subject, report)
