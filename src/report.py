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
NOW_BR = datetime.now(BRASILIA)
TODAY_BR = NOW_BR.strftime("%d/%m/%Y")

# Segunda-feira = janela de 72h (cobre o fim de semana); demais dias = 24h
EMAIL_WINDOW_HOURS = 72 if NOW_BR.weekday() == 0 else 24
SINCE_IMAP = (NOW_BR - timedelta(hours=EMAIL_WINDOW_HOURS)).strftime("%d-%b-%Y")

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

    known_vessels = ["HERAKLITOS", "PARAGON", "EKATERINA", "LEFTERIS T", "LEFTERIS",
                     "DAHLIA", "MARCOS DIAS", "CALLIO", "RAYS", "APOGEE", "AGIA ELENI",
                     "SEPETIBA", "VANTAGE ROSE", "CHINTANA", "DISCOVERER", "HANZE",
                     "ATAYAL", "PIO GRANDE", "ADVENTURER", "LOYALTY", "SEAHEAVEN",
                     "LOWLANDS", "REVENGER", "PARANA WARRIOR"]
    found = [v for v in known_vessels if v in body.upper()]
    if len(found) < 3:
        print(f"Poucos navios identificados ({found}), usando lista completa de fallback")
        return FALLBACK_VESSELS, body

    print(f"Navios na position list: {found}")
    return found, body


def fetch_all_emails(mail, vessels):
    """Busca emails da janela definida (24h normais, 72h nas segundas) sobre cada navio."""
    mail.select("inbox")
    emails_by_vessel = {}

    for vessel in vessels:
        ids = set()
        for criteria in [f'SUBJECT "{vessel}"', f'BODY "{vessel}"']:
            _, data = mail.search(None, f'(SINCE "{SINCE_IMAP}" {criteria})')
            ids.update(data[0].split())

        vessel_emails = []
        for num in sorted(ids)[-15:]:
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

    return emails_by_vessel


def fetch_previous_briefing(mail):
    """Busca o briefing do dia anterior na caixa de enviados."""
    try:
        mail.select('"[Gmail]/Sent Mail"')
        _, data = mail.search(None, 'SUBJECT "Briefing da Frota"')
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
                    body = part.get_payload(decode=True).decode("utf-8", errors="ignore")[:3000]
                    break
        else:
            body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")[:3000]
        print(f"Briefing anterior encontrado: {msg['subject']}")
        return body
    except Exception as e:
        print(f"Nao foi possivel buscar briefing anterior: {e}")
        return ""


def generate_briefing(emails_by_vessel, position_list_body="", previous_briefing=""):
    emails_text = ""
    for vessel, emails in emails_by_vessel.items():
        emails_text += f"\n\n=== {vessel} ===\n"
        if not emails:
            emails_text += f"Sem emails nas ultimas {EMAIL_WINDOW_HOURS}h.\n"
        for e in emails:
            emails_text += f"\nAssunto: {e['subject']}\nDe: {e['from']}\nData: {e['date']}\n{e['body']}\n---\n"

    position_list_section = ""
    if position_list_body:
        position_list_section = f"\n\nPOSITION LIST MAIS RECENTE (use como referencia de status em caso de duvida):\n{position_list_body[:3000]}"

    previous_section = ""
    if previous_briefing:
        previous_section = f"\n\nBRIEFING DO DIA ANTERIOR (use para identificar o que mudou e o que continua em aberto):\n{previous_briefing}"

    prompt = f"""Voce e o redator do briefing diario de operacoes maritimas da Lyra Shipping (Rio de Janeiro).
Hoje e {TODAY_BR}. Os emails abaixo cobrem as ultimas {EMAIL_WINDOW_HOURS} horas.

{position_list_section}
{previous_section}

EMAILS POR NAVIO:
{emails_text}

---

PAPEL
Voce envia uma unica mensagem de email, no proprio corpo, para todos os membros da empresa.
O objetivo e contextualizar o time e lembrar o que pode exigir atencao hoje.
Pense em "leitura de 2 minutos para comecar o dia sabendo onde a frota esta e o que esta pegando".

PUBLICO E TOM
- Publico: toda a empresa (operacoes, comercial, financeiro, juridico, administrativo, diretoria). Nem todo mundo e tecnico.
- Tom: objetivo, claro, profissional, levemente informal. Frases curtas. Sem jargao desnecessario.
- Extensao: o email inteiro deve caber em 1 tela — idealmente 200-400 palavras. Se tiver muita coisa, priorize.

O QUE INCLUIR
Para cada navio em viagem ativa, transmita em 1-3 frases:
- Onde o navio esta (em termos gerais: "navegando para X", "no porto de Y carregando", "fundeado em Z aguardando berco")
- Como a viagem esta indo (no prazo, atrasando, adiantando, parado)
- O que o time precisa saber ou fazer, se houver algo que envolva outras areas

O QUE NAO INCLUIR
- Quantidade de combustivel / ROB
- Horarios em formato UTC ou notacao tecnica — use linguagem natural ("hoje a tarde", "amanha", "fim de semana", "proxima semana")
- Coordenadas, velocidade em nos, milhas restantes
- Numeros de NOR, clausulas de CP, IMOs
- Dados de tripulacao
- Listas exaustivas de emails analisados

REGRAS DE REDACAO
- Uma ideia por bullet. Se precisar explicar mais, e detalhe demais para este formato.
- Nome do navio em negrito (**NOME**) para facilitar escaneamento visual.
- Datas em linguagem natural. Se precisar de data exata, use DD/MM sem ano.
- Destaque o que o time precisa FAZER, nao o que o sistema ja registrou.
  Ex.: em vez de "ETA atualizada para 22/04", escreva "chegada em Santos deve atrasar para o fim de semana — pode impactar agenda comercial da semana que vem."
- Nao use adjetivos subjetivos sem base ("tudo correndo bem"). Prefira fatos: "sem intercorrencias nas ultimas 24h".
- Severidade sinalizada por emoji (🟢 normal, 🟡 atencao, 🔴 urgente).
- Emails mais recentes tem PRIORIDADE sobre emails mais antigos — se houver contradicao, prevalece o mais recente.
- Em caso de duvida sobre posicao ou status, use a POSITION LIST como referencia definitiva.
- Jamais invente. Se a informacao nao esta nos emails, nao inclua, ou escreva "sem atualizacao recente sobre [navio]".
- Se um assunto do briefing anterior continua em aberto, mantenha-o sem repetir todo o contexto.

ESTRUTURA DO EMAIL (siga este template):
Assunto: Briefing da Frota — {TODAY_BR} | [resumo em 5 palavras do destaque do dia]

Bom dia a todos,

[1 frase abrindo o panorama geral]

🚢 Em destaque hoje
- [**Navio**] — [o que esta acontecendo em 1 linha, e o que isso pede do time, se algo].

🟢 Andamento normal
- [**Navio X**], [**Navio Y**] — [status geral em 1 linha cada, bem enxuto].

⚠️ Atencao / pendencias
- [Item que pode virar problema, com area sugerida entre parenteses].

📅 Proximos marcos
- [Chegadas, saidas, operacoes relevantes previstas para os proximos 2-3 dias, em linguagem natural].

Qualquer duvida, respondam a este email ou falem com Operacoes.

— Briefing gerado automaticamente a partir dos emails operacionais.

CASOS ESPECIAIS
- Dia sem novidades: o email ainda sai, bem curto, confirmando que a frota segue sem intercorrencias e listando apenas proximos marcos.
- Informacao ambigua: registrar como "a confirmar" — nao expor confusao para a empresa.
- Fim de semana / feriado: tom ainda mais enxuto.
- Omita secoes que nao tiverem conteudo (nao escreva "nada a reportar").

EXEMPLO DE BRIEFING (referencia de tom e tamanho):
Assunto: Briefing da Frota — 18/04 | Heraklitos com contradicao de posicao

Bom dia a todos,

Dia com atencao para o Heraklitos, cuja posicao apresenta contradicao entre fontes — equipe de operacoes ja esta apurando.

🚢 Em destaque hoje
- 🔴 **Heraklitos** — Pre Notice indica zarpe no dia 17, mas Noon Report de hoje ainda posiciona o navio no berco em Brake. Contradicao sendo investigada com Narval e agente J.Muller. Alem disso, saldo de 20% do frete so deve entrar na quarta por conta do feriado — Financeiro acompanhar.
- 🟡 **Lefteris T** — em descarga em Salvador, com zarpe previsto amanha. Atencao para documento de agua de lastro ainda pendente com agente em Areia Branca — prazo apertado.

🟢 Andamento normal
- **Marcos Dias** — fundeado em San Lorenzo aguardando berco para carregar trigo. Dentro do previsto.
- **Ekaterina** — fundeado em Imbituba aguardando vez no berco. ETB confirmado para o fim de semana.
- **Dahlia** — chegando a Vitoria hoje a tarde para inicio da descarga de malte.
- **Paragon** — chegando a Nueva Palmira amanha para carregamento de cevada.
- **Callio** — navegando pelo Mediterraneo, chega a Ortona na proxima terca.

⚠️ Atencao / pendencias
- Demonstrativo final de frete do Paragon (viagem anterior) ainda nao enviado para contraparte — Operacoes verificar com urgencia.

📅 Proximos marcos
- **Lefteris T** zarpa de Salvador amanha e chega em Areia Branca na proxima segunda.
- **Paragon** inicia carregamento em Nueva Palmira amanha.
- **Callio** chega em Ortona na proxima terca — documentacao ao agente deve sair hoje.

Qualquer duvida, respondam a este email ou falem com Operacoes.

— Briefing gerado automaticamente a partir dos emails operacionais.

Agora gere o briefing de hoje ({TODAY_BR}) seguindo exatamente este modelo. Comece com a linha do Assunto."""

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
    mail = fetch_emails_imap()
    vessels, position_list_body = get_vessels_from_position_list(mail)
    emails_by_vessel = fetch_all_emails(mail, vessels)
    previous_briefing = fetch_previous_briefing(mail)
    mail.logout()
    briefing = generate_briefing(emails_by_vessel, position_list_body, previous_briefing)

    # Extrai assunto da primeira linha gerada pelo Claude
    lines = briefing.strip().splitlines()
    subject = lines[0].replace("Assunto:", "").strip() if lines[0].startswith("Assunto:") else f"Briefing da Frota — {TODAY_BR}"
    body = "\n".join(lines[1:]).strip() if lines[0].startswith("Assunto:") else briefing

    send_email(subject, body)
