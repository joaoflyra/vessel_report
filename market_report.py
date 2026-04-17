import anthropic
import imaplib
import smtplib
import email
import email.header
import os
import json
import io
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime, timedelta

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT

GMAIL_USER = os.environ.get("GMAIL_USER", "joao@lyrashipping.com.br")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
REPORT_RECIPIENT = os.environ.get("REPORT_RECIPIENT", GMAIL_USER)

NAVY = colors.HexColor("#1a2e4a")
BLUE = colors.HexColor("#2c5f8a")
LIGHT = colors.HexColor("#eaf2fb")
ACCENT = colors.HexColor("#e8f0e8")
WHITE = colors.white
GRAY = colors.HexColor("#666666")
RED = colors.HexColor("#c0392b")


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

        subject = decode_header_value(msg.get("Subject", ""))
        sender = decode_header_value(msg.get("From", ""))
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

        emails.append({"subject": subject, "from": sender, "body": body[:600]})

    mail.close()
    mail.logout()

    # pre-filtro local: manter apenas emails com conteudo de broker
    broker_keywords = [
        "dwt", "laycan", "tonnage", "open ", "aberto", "charter",
        "vessel", "cargo", "freight", "bulk", "moloo", "molco",
        "shinc", "fhex", "pwwd", "demurrage", "dispatch",
        "eta ", "ets ", "etb ", "load port", "disch", "l/d",
        "bss", "tct", "t/c ", "time charter", "voyage charter",
        "kamsarmax", "panamax", "supramax", "ultramax", "handysize",
        "capesize", "geared", "grabber", "scrubber",
    ]

    filtered = []
    for e in emails:
        combined = (e["subject"] + " " + e["body"]).lower()
        if any(kw in combined for kw in broker_keywords):
            filtered.append(e)

    print(f"[INFO] Pre-filtro: {len(filtered)} emails com conteudo de broker (de {len(emails)} total)")
    return filtered


def analyze_market(emails):
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    today = datetime.now().strftime("%d/%m/%Y")

    email_content = "\n\n---\n\n".join([
        f"De: {e['from']}\nAssunto: {e['subject']}\n\n{e['body']}"
        for e in emails
    ]) if emails else "Nenhum email."

    prompt = f"""Você é um analista de mercado de fretes maritimos (dry bulk).

Analise os {len(emails)} emails abaixo e retorne um JSON estruturado com os dados de mercado.

REGRAS:
- Identifique emails de brokers com critério AMPLO. Brokers incluem: Clarksons, SSY, Fearnleys, Braemar, Howe Robinson, BRS, Banchero Costa, DHP Maritime, North Harbour Shipping, Prime Shipping & Chartering (chartering@primeshipping.org), Niavigrains, Skyhi Shipping, Guanabara Shipping, Teomare, Smarship, Ifchor Galbraiths, G.Moundreas, PNA Shipbrokers, United Seas Corp, Grain Compass, Laasco Dry, e qualquer remetente com conteúdo de circular maritima (tonnage, laycan, DWT, L/D, etc.).
- Em caso de dúvida, INCLUA.
- Extraia APENAS: ofertas de navio (tonnage), procuras de navio (orders), ofertas de carga (cargo offer).
- NAO inclua procuras de carga (cargo wanted/backhaul/TCT seeking cargo).

Porte dos navios:
- Capesize: 100.000+ DWT
- Panamax/Kamsarmax: 65.000-99.999 DWT
- Supramax/Ultramax: 45.000-64.999 DWT
- Handysize: abaixo de 45.000 DWT

Regiões:
- Americas (ECSA)  — Costa Leste América do Sul (Brasil, Argentina, Uruguai)
- Americas (USG/USEC) — Golfo e Costa Leste EUA
- Americas (Outros) — Caribe, WCSA, América Central
- Europa / Mediterrâneo
- Ásia (Far East)
- Ásia (Outros) — SE Asia, Sul da Asia
- Oriente Médio / Índia
- África
- Oceania

Retorne APENAS um JSON válido, sem texto antes ou depois, seguindo este schema:

{{
  "summary": {{
    "total_emails": {len(emails)},
    "broker_emails": 0,
    "ignored_emails": 0,
    "tonnage_count": 0,
    "orders_count": 0,
    "cargo_offers_count": 0
  }},
  "brokers_seen": ["broker1", "broker2"],
  "tonnage": [
    {{
      "vessel": "MV NOME",
      "dwt": 57000,
      "year": 2013,
      "size_class": "Supramax/Ultramax",
      "open_port": "Paranagua",
      "open_date": "20-25 Abr",
      "region": "Americas (ECSA)",
      "broker": "Clarksons Hellas",
      "notes": "graos, limpo"
    }}
  ],
  "orders": [
    {{
      "cargo": "Iron Ore",
      "quantity": "170.000 MT",
      "load_port": "Dampier",
      "discharge_port": "Qingdao",
      "laycan": "01-03 Mai",
      "charterer": "Rio Tinto",
      "size_class": "Capesize",
      "region": "Oceania",
      "broker": "Prime S&C",
      "type": "voyage",
      "notes": ""
    }}
  ],
  "cargo_offers": [
    {{
      "cargo": "Soja em bags",
      "quantity": "25.000 MT",
      "load_port": "Paranagua",
      "discharge_port": "Jeddah",
      "laycan": "26 Abr-05 Mai",
      "shipper": "n/i",
      "region": "Americas (ECSA)",
      "broker": "Prime S&C",
      "notes": ""
    }}
  ],
  "highlights": []
}}

---
Emails ({len(emails)} total):
{email_content}"""

    response = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=16000,
        messages=[{"role": "user", "content": prompt}],
    )

    text = response.content[0].text.strip()
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
    data = json.loads(text)

    # gera destaques com Sonnet (melhor analise)
    data["highlights"] = generate_highlights(data, email_content)
    return data


def generate_highlights(data, email_content):
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    s = data["summary"]

    tonnage_summary = "\n".join([
        f"- {v['vessel']} / {v['dwt']} DWT / {v['size_class']} — Aberto {v['open_port']} {v['open_date']} ({v['region']})"
        for v in data.get("tonnage", [])
    ])
    orders_summary = "\n".join([
        f"- {o['quantity']} {o['cargo']} {o['load_port']}->{o['discharge_port']} laycan {o['laycan']} ({o['size_class']})"
        for o in data.get("orders", [])
    ])
    cargo_summary = "\n".join([
        f"- {c['quantity']} {c['cargo']} {c['load_port']}->{c['discharge_port']} laycan {c['laycan']}"
        for c in data.get("cargo_offers", [])
    ])

    prompt = f"""Você é um analista sênior de mercado de fretes maritimos dry bulk.

Com base nos dados extraídos abaixo, escreva de 4 a 6 destaques analíticos do dia.
Seja perspicaz — identifique tendências, concentrações, desequilíbrios oferta/demanda, rotas em destaque, níveis de frete mencionados, comportamento de brokers específicos.

Dados extraídos:
- Total emails broker: {s['broker_emails']}
- Tonelagem ofertada: {s['tonnage_count']} navios
- Ordens de frete: {s['orders_count']}
- Ofertas de carga: {s['cargo_offers_count']}

TONNAGE:
{tonnage_summary or 'nenhuma'}

ORDERS:
{orders_summary or 'nenhuma'}

CARGO OFFERS:
{cargo_summary or 'nenhuma'}

Retorne APENAS uma lista JSON de strings, sem texto adicional:
["destaque 1", "destaque 2", "destaque 3"]"""

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1000,
        messages=[{"role": "user", "content": prompt}],
    )

    text = response.content[0].text.strip()
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
    return json.loads(text)


def build_pdf(data):
    today = datetime.now().strftime("%d/%m/%Y")
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=1.8*cm, rightMargin=1.8*cm,
        topMargin=1.5*cm, bottomMargin=1.5*cm
    )

    styles = getSampleStyleSheet()
    s_title = ParagraphStyle("title", fontSize=18, textColor=WHITE, alignment=TA_CENTER, fontName="Helvetica-Bold", spaceAfter=2)
    s_sub   = ParagraphStyle("sub", fontSize=10, textColor=LIGHT, alignment=TA_CENTER, fontName="Helvetica", spaceAfter=0)
    s_sec   = ParagraphStyle("sec", fontSize=11, textColor=WHITE, fontName="Helvetica-Bold", spaceBefore=4, spaceAfter=4, leftIndent=4)
    s_sub2  = ParagraphStyle("sub2", fontSize=9.5, textColor=NAVY, fontName="Helvetica-Bold", spaceBefore=6, spaceAfter=2)
    s_body  = ParagraphStyle("body", fontSize=8.5, textColor=colors.black, fontName="Helvetica", spaceAfter=2, leading=12)
    s_hl    = ParagraphStyle("hl", fontSize=9, textColor=colors.black, fontName="Helvetica", leftIndent=10, spaceAfter=3, leading=13)
    s_note  = ParagraphStyle("note", fontSize=8, textColor=GRAY, fontName="Helvetica-Oblique", spaceAfter=2)

    story = []

    # Header
    header_data = [[Paragraph(f"LYRA SHIPPING", s_title)], [Paragraph(f"Relatório de Mercado — {today}", s_sub)]]
    header_table = Table(header_data, colWidths=[doc.width])
    header_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), NAVY),
        ("TOPPADDING", (0,0), (-1,-1), 10),
        ("BOTTOMPADDING", (0,0), (-1,-1), 10),
        ("LEFTPADDING", (0,0), (-1,-1), 10),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 0.4*cm))

    # Summary boxes
    s = data["summary"]
    boxes = [
        ["Emails\nAnalisados", str(s["total_emails"])],
        ["Brokers\nIdentificados", str(s["broker_emails"])],
        ["Ofertas\nde Navio", str(s["tonnage_count"])],
        ["Procuras\nde Navio", str(s["orders_count"])],
        ["Ofertas\nde Carga", str(s["cargo_offers_count"])],
    ]
    box_style = ParagraphStyle("bk", fontSize=8, textColor=GRAY, alignment=TA_CENTER, fontName="Helvetica")
    box_num   = ParagraphStyle("bn", fontSize=20, textColor=NAVY, alignment=TA_CENTER, fontName="Helvetica-Bold")

    box_cells = [[
        [Paragraph(b[1], box_num), Paragraph(b[0], box_style)]
        for b in boxes
    ]]
    col_w = doc.width / len(boxes)
    summary_table = Table(box_cells, colWidths=[col_w]*len(boxes), rowHeights=[1.4*cm])
    summary_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), LIGHT),
        ("BOX", (0,0), (-1,-1), 0.5, BLUE),
        ("INNERGRID", (0,0), (-1,-1), 0.5, WHITE),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("TOPPADDING", (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 0.4*cm))

    def section_header(title):
        t = Table([[Paragraph(title, s_sec)]], colWidths=[doc.width])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), BLUE),
            ("TOPPADDING", (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
            ("LEFTPADDING", (0,0), (-1,-1), 8),
        ]))
        return t

    def count_table(items, key):
        from collections import Counter
        counts = Counter(i[key] for i in items)
        if not counts:
            return None
        rows = [["", "Quantidade"]]
        for k, v in sorted(counts.items(), key=lambda x: -x[1]):
            rows.append([k, str(v)])
        t = Table(rows, colWidths=[doc.width * 0.75, doc.width * 0.25])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), NAVY),
            ("TEXTCOLOR", (0,0), (-1,0), WHITE),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE", (0,0), (-1,-1), 8),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [WHITE, LIGHT]),
            ("GRID", (0,0), (-1,-1), 0.3, colors.HexColor("#cccccc")),
            ("ALIGN", (1,0), (1,-1), "CENTER"),
            ("TOPPADDING", (0,0), (-1,-1), 3),
            ("BOTTOMPADDING", (0,0), (-1,-1), 3),
            ("LEFTPADDING", (0,0), (-1,-1), 6),
        ]))
        return t

    # ── TONNAGE ──────────────────────────────────
    tonnage = data.get("tonnage", [])
    if tonnage:
        story.append(section_header(f"OFERTAS DE NAVIO — TONNAGE   ({len(tonnage)} navios)"))
        story.append(Spacer(1, 0.2*cm))

        # count by size
        story.append(Paragraph("Por porte:", s_sub2))
        ct = count_table(tonnage, "size_class")
        if ct:
            story.append(ct)
        story.append(Spacer(1, 0.2*cm))

        # count by region
        story.append(Paragraph("Por região:", s_sub2))
        cr = count_table(tonnage, "region")
        if cr:
            story.append(cr)
        story.append(Spacer(1, 0.3*cm))

        # listing by size + region
        from collections import defaultdict
        by_size = defaultdict(lambda: defaultdict(list))
        size_order = ["Capesize", "Panamax/Kamsarmax", "Supramax/Ultramax", "Handysize"]
        for v in tonnage:
            by_size[v["size_class"]][v["region"]].append(v)

        for sz in size_order:
            if sz not in by_size:
                continue
            story.append(Paragraph(f"▸  {sz}", s_sub2))
            for region, vessels in sorted(by_size[sz].items()):
                story.append(Paragraph(f"{region}  ({len(vessels)})", ParagraphStyle("rg", fontSize=8.5, textColor=BLUE, fontName="Helvetica-Bold", leftIndent=12, spaceAfter=1)))
                for v in vessels:
                    yr = f" / {v['year']}" if v.get("year") else ""
                    note = f"  —  {v['notes']}" if v.get("notes") else ""
                    line = f"<b>{v['vessel']}</b>  /  {v['dwt']:,} DWT{yr}  —  Aberto {v['open_port']} {v['open_date']}{note}  —  <i>Broker: {v['broker']}</i>"
                    story.append(Paragraph(line, ParagraphStyle("vl", fontSize=8, fontName="Helvetica", leftIndent=24, spaceAfter=2, leading=11)))
        story.append(Spacer(1, 0.4*cm))

    # ── ORDERS ───────────────────────────────────
    orders = data.get("orders", [])
    if orders:
        story.append(section_header(f"PROCURAS DE NAVIO — ORDERS   ({len(orders)} ordens)"))
        story.append(Spacer(1, 0.2*cm))

        story.append(Paragraph("Por porte:", s_sub2))
        ct = count_table(orders, "size_class")
        if ct: story.append(ct)
        story.append(Spacer(1, 0.2*cm))

        story.append(Paragraph("Por região:", s_sub2))
        cr = count_table(orders, "region")
        if cr: story.append(cr)
        story.append(Spacer(1, 0.3*cm))

        by_size = defaultdict(list)
        for o in orders:
            by_size[o["size_class"]].append(o)

        for sz in size_order:
            if sz not in by_size:
                continue
            story.append(Paragraph(f"▸  {sz}", s_sub2))
            for o in by_size[sz]:
                tp = " (T/C)" if o.get("type") == "tct" else ""
                ch = f"  —  Charterer: {o['charterer']}" if o.get("charterer") and o["charterer"] not in ("n/i", "n/d", "") else ""
                note = f"  —  {o['notes']}" if o.get("notes") else ""
                line = (f"<b>{o['quantity']}</b> {o['cargo']}{tp}  —  "
                        f"{o['load_port']} → {o['discharge_port']}  —  "
                        f"Laycan: {o['laycan']}{ch}{note}  —  "
                        f"<i>Broker: {o['broker']}</i>")
                story.append(Paragraph(line, ParagraphStyle("ol", fontSize=8, fontName="Helvetica", leftIndent=16, spaceAfter=3, leading=11)))
        story.append(Spacer(1, 0.4*cm))

    # ── CARGO OFFERS ─────────────────────────────
    cargo = data.get("cargo_offers", [])
    if cargo:
        story.append(section_header(f"OFERTAS DE CARGA — CARGO OFFER   ({len(cargo)} itens)"))
        story.append(Spacer(1, 0.2*cm))

        story.append(Paragraph("Por região de carregamento:", s_sub2))
        cr = count_table(cargo, "region")
        if cr: story.append(cr)
        story.append(Spacer(1, 0.3*cm))

        by_region = defaultdict(list)
        for c in cargo:
            by_region[c["region"]].append(c)

        for region, items in sorted(by_region.items()):
            story.append(Paragraph(f"▸  {region}  ({len(items)})", s_sub2))
            for c in items:
                sh = f"  —  Shipper: {c['shipper']}" if c.get("shipper") and c["shipper"] not in ("n/i", "n/d", "") else ""
                note = f"  —  {c['notes']}" if c.get("notes") else ""
                line = (f"<b>{c['quantity']}</b> {c['cargo']}  —  "
                        f"{c['load_port']} → {c['discharge_port']}  —  "
                        f"Laycan: {c['laycan']}{sh}{note}  —  "
                        f"<i>Broker: {c['broker']}</i>")
                story.append(Paragraph(line, ParagraphStyle("cl", fontSize=8, fontName="Helvetica", leftIndent=16, spaceAfter=3, leading=11)))
        story.append(Spacer(1, 0.4*cm))

    # ── HIGHLIGHTS ───────────────────────────────
    highlights = data.get("highlights", [])
    if highlights:
        story.append(section_header("DESTAQUES DO DIA"))
        story.append(Spacer(1, 0.2*cm))
        for h in highlights:
            story.append(Paragraph(f"•  {h}", s_hl))
        story.append(Spacer(1, 0.4*cm))

    # ── BROKERS ──────────────────────────────────
    brokers = data.get("brokers_seen", [])
    if brokers:
        story.append(section_header("BROKERS IDENTIFICADOS"))
        story.append(Spacer(1, 0.2*cm))
        story.append(Paragraph(", ".join(sorted(brokers)), s_note))

    doc.build(story)
    buf.seek(0)
    return buf.read()


def send_report(pdf_bytes):
    today = datetime.now().strftime("%d/%m/%Y")
    today_file = datetime.now().strftime("%Y-%m-%d")

    msg = MIMEMultipart()
    msg["Subject"] = f"Relatório de Mercado — {today}"
    msg["From"] = GMAIL_USER
    msg["To"] = REPORT_RECIPIENT

    body = MIMEText(f"Relatório de mercado em anexo — {today}", "plain", "utf-8")
    msg.attach(body)

    pdf_part = MIMEApplication(pdf_bytes, _subtype="pdf")
    pdf_part.add_header("Content-Disposition", "attachment", filename=f"mercado_{today_file}.pdf")
    msg.attach(pdf_part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_USER, REPORT_RECIPIENT, msg.as_string())

    print(f"[OK] PDF enviado para {REPORT_RECIPIENT}")


if __name__ == "__main__":
    print(f"[{datetime.now().strftime('%d/%m/%Y %H:%M')}] Gerando relatorio de mercado...")

    if not GMAIL_APP_PASSWORD:
        print("[ERRO] GMAIL_APP_PASSWORD nao configurado.")
        exit(1)
    if not ANTHROPIC_API_KEY:
        print("[ERRO] ANTHROPIC_API_KEY nao configurado.")
        exit(1)

    emails = fetch_broker_emails()
    print(f"[INFO] {len(emails)} emails encontrados.")

    data = analyze_market(emails)
    s = data["summary"]
    print(f"[INFO] Brokers: {s['broker_emails']} | Tonnage: {s['tonnage_count']} | Orders: {s['orders_count']} | Cargo: {s['cargo_offers_count']}")

    pdf = build_pdf(data)
    print(f"[INFO] PDF gerado ({len(pdf):,} bytes)")

    send_report(pdf)
