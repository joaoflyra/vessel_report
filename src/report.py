import os
import imaplib
import smtplib
import email as email_lib
import time
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
SINCE_48H_IMAP = (datetime.now(BRASILIA) - timedelta(hours=48)).strftime("%d-%b-%Y")

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
    """Busca emails das ultimas 48h sobre cada navio."""
    mail.select("inbox")
    emails_by_vessel = {}

    for vessel in vessels:
        ids = set()
        for criteria in [f'SUBJECT "{vessel}"', f'BODY "{vessel}"']:
            _, data = mail.search(None, f'(SINCE "{SINCE_48H_IMAP}" {criteria})')
            ids.update(data[0].split())

        vessel_emails = []
        for num in sorted(ids)[-5:]:
            _, msg_data = mail.fetch(num, "(RFC822)")
            msg = email_lib.message_from_bytes(msg_data[0][1])
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode("utf-8", errors="ignore")[:800]
                        break
            else:
                body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")[:800]
            vessel_emails.append({
                "subject": str(msg["subject"] or ""),
                "from": str(msg["from"] or ""),
                "date": str(msg["date"] or ""),
                "body": body,
            })
        emails_by_vessel[vessel] = vessel_emails
        print(f"{vessel}: {len(vessel_emails)} emails encontrados")

    return emails_by_vessel


def fetch_previous_report(mail):
    """Busca o relatorio do dia anterior na caixa de enviados."""
    try:
        mail.select('"[Gmail]/Sent Mail"')
        _, data = mail.search(None, 'SUBJECT "LYRA SHIPPING . RELATORIO DIARIO"')
        ids = data[0].split()
        if not ids:
            return ""
        latest = sorted(ids)[-1]
        _, msg_data = mail.fetch(latest, "(RFC822)")
        msg = email_lib.message_from_bytes(msg_data[0][1])
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode("utf-8", errors="ignore")[:5000]
                    break
        else:
            body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")[:5000]
        print(f"Relatorio anterior encontrado: {msg['subject']}")
        return body
    except Exception as e:
        print(f"Nao foi possivel buscar relatorio anterior: {e}")
        return ""


def generate_report(emails_by_vessel, position_list_body="", previous_report=""):
    emails_text = ""
    for vessel, emails in emails_by_vessel.items():
        emails_text += f"\n\n=== {vessel} ===\n"
        if not emails:
            emails_text += "Sem emails nas ultimas 48h.\n"
        for e in emails:
            emails_text += f"\nAssunto: {e['subject']}\nDe: {e['from']}\nData: {e['date']}\n{e['body']}\n---\n"

    position_list_section = ""
    if position_list_body:
        position_list_section = f"\n\nPOSITION LIST MAIS RECENTE:\n{position_list_body[:3000]}"

    previous_report_section = ""
    if previous_report:
        previous_report_section = f"\n\nRELATORIO DO DIA ANTERIOR (use como contexto para assuntos em andamento):\n{previous_report}"

    prompt = f"""Voce e um assistente experiente de operacoes maritimas da Lyra Shipping (Rio de Janeiro).
Gere o relatorio diario de navios com base nos emails das ultimas 48 horas (ate {TODAY_BR}).

{position_list_section}
{previous_report_section}

EMAILS DAS ULTIMAS 48H POR NAVIO:
{emails_text}

INSTRUCOES:

1. ORDENACAO POR IMPORTANCIA:
   - 🔴 PRIMEIRO: Navios com problemas criticos (retido, reparos obrigatorios, P&I, demurrage correndo, situacoes legais urgentes)
   - 🟡 SEGUNDO: Navios com pendencias importantes (fundeado aguardando berco, descarregando, pendencias financeiras abertas, documentos pendentes, decisoes necessarias)
   - 🟢 TERCEIRO: Navios navegando normalmente ou carregando sem problemas — seja BREVE e direto para esses

2. NIVEL DE DETALHE:
   - Navios criticos (🔴): detalhe completo com todos os pontos urgentes
   - Navios com pendencias (🟡): detalhe relevante, foco nas pendencias
   - Navios tranquilos (🟢): apenas status, rota, ETA proxima escala — sem blocos vazios

3. FOCO DO RELATORIO:
   - Priorize o andamento operacional e o progresso de cada viagem (onde esta, o que esta fazendo, o que vem a seguir)
   - Destaque POSSIVEIS PENDENCIAS que a empresa deve acompanhar, mesmo que nao confirmadas
   - NAO inclua detalhes de combustivel/bunker, EXCETO se houver previsao concreta de necessidade de abastecimento proxima
   - Nao inclua horarios exatos desnecessarios — mencione apenas datas/periodos relevantes para decisoes
   - Se um assunto apareceu no relatorio anterior e continua em andamento, mantenha-o sem repetir todo o contexto

4. PESO DAS INFORMACOES:
   - Emails mais recentes tem PRIORIDADE sobre emails mais antigos — se houver contradição, prevalece o mais recente
   - Em caso de duvida sobre posicao, status ou proxima escala do navio, use a POSITION LIST como referencia definitiva
   - Se a position list contradiz os emails, destaque a contradição e mencione ambas as fontes

5. FORMATO POR NAVIO (siga rigorosamente este modelo):

[EMOJI]  M/V [NOME] · Voy. [NUMERO]
  Status    [descricao detalhada — porto, berco, operacao em curso]
  Rota      [origem → destino] | [carga e quantidade] | [viagem]

  URGENTE (somente se houver — use ⚠️ para cada item critico)
  ⚠️ [item urgente com contexto suficiente para acao imediata]

  OPERACIONAL
  - [progresso da operacao com dados concretos: quantidades, percentuais, datas]
  - [use ✔ para itens confirmados/concluidos]

  POSSIVEIS PENDENCIAS (somente se houver suspeita)
  - [item que merece atencao mesmo sem confirmacao]

  PROXIMA VIAGEM (somente se relevante)
  - [carga, rota, laydays, agente]

  PROXIMOS PASSOS
  - [acoes concretas e objetivas, com responsavel quando conhecido]

6. VIAGENS ANTERIORES: Se houver emails sobre viagens ja encerradas (demonstrativos de frete, saldos pendentes), liste ao final:

  PENDENCIAS DE VIAGENS ANTERIORES
  - M/V [NOME] voy [X]: [descricao]

7. Se um navio nao tiver emails nas ultimas 48h, inclua brevemente com status do relatorio anterior se disponivel.

Cabecalho do relatorio:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
LYRA SHIPPING  ·  RELATORIO DIARIO  ·  {TODAY_BR}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Separador entre navios: ──────────────────────────────────────"""

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    for attempt in range(3):
        try:
            message = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=8096,
                messages=[{"role": "user", "content": prompt}]
            )
            return message.content[0].text
        except anthropic.RateLimitError:
            if attempt < 2:
                wait = 60 * (attempt + 1)
                print(f"Rate limit atingido, aguardando {wait}s antes de tentar novamente...")
                time.sleep(wait)
            else:
                raise


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
    previous_report = fetch_previous_report(mail)
    mail.logout()
    report = generate_report(emails_by_vessel, position_list_body, previous_report)
    subject = f"LYRA SHIPPING . RELATORIO DIARIO . {TODAY_BR}"
    send_email(subject, report)
