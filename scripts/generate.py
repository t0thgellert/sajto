#!/usr/bin/env python3
"""
Sajtófigyelő automatikus feldolgozó
Outlook-ból letölti a legújabb Word mellékletet, kinyeri a struktúrált adatokat,
majd frissíti a GitHub Pages HTML oldalt.
"""

import os, re, json, base64, datetime, requests
from pathlib import Path
from urllib.parse import urlparse

SENDER_EMAIL      = os.environ["PR_SENDER_EMAIL"]
GRAPH_CLIENT_ID   = os.environ["MS_CLIENT_ID"]
GRAPH_CLIENT_SEC  = os.environ["MS_CLIENT_SECRET"]
GRAPH_TENANT_ID   = os.environ["MS_TENANT_ID"]
GRAPH_USER_EMAIL  = os.environ["MS_USER_EMAIL"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
OUTPUT_FILE       = Path("docs/index.html")

# ── Microsoft Graph ──────────────────────────────────────────────────────────
def get_graph_token():
    r = requests.post(
        f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}/oauth2/v2.0/token",
        data={"grant_type":"client_credentials","client_id":GRAPH_CLIENT_ID,
              "client_secret":GRAPH_CLIENT_SEC,"scope":"https://graph.microsoft.com/.default"},
        timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]

def graph_get(token, path, **kw):
    r = requests.get(f"https://graph.microsoft.com/v1.0{path}",
        headers={"Authorization":f"Bearer {token}"}, timeout=30, **kw)
    r.raise_for_status()
    return r.json()

def fetch_latest_word_attachment(token):
    msgs = graph_get(token, f"/users/{GRAPH_USER_EMAIL}/messages", params={
        "$filter": f"from/emailAddress/address eq '{SENDER_EMAIL}'",
        "$orderby":"receivedDateTime desc","$top":"5",
        "$select":"id,subject,receivedDateTime,hasAttachments"
    }).get("value", [])
    for msg in msgs:
        if not msg.get("hasAttachments"): continue
        for att in graph_get(token, f"/users/{GRAPH_USER_EMAIL}/messages/{msg['id']}/attachments").get("value",[]):
            if att.get("name","").lower().endswith((".docx",".doc")):
                print(f"✓ Melléklet: {att['name']}")
                return msg["subject"], base64.b64decode(att["contentBytes"])
    raise RuntimeError("Nem találtam .docx mellékletet.")

# ── Word kinyerés ─────────────────────────────────────────────────────────────
def extract_text_from_docx(docx_bytes):
    import io
    from docx import Document
    doc = Document(io.BytesIO(docx_bytes))

    # Kinyerjük a paragrafusokat és a hozzájuk tartozó linkeket (hyperlink)
    paragraphs = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        # Hyperlink kinyerése a paragrafusból
        url = None
        for rel in para.part.rels.values():
            if "hyperlink" in rel.reltype:
                # csak az aktuális paragrafushoz tartozót
                pass
        # Egyszerűbb: összegyűjtjük a teljes szöveget
        paragraphs.append(text)

    return "\n\n".join(paragraphs)

def extract_docx_with_links(docx_bytes):
    """Kinyeri a szöveget és az URL-eket a Word dokumentumból."""
    import io
    from docx import Document
    from lxml import etree

    doc = Document(io.BytesIO(docx_bytes))
    lines = []

    # Namespace
    W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    R  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    # Relationships: id -> url
    rels = {}
    for rel_id, rel in doc.part.rels.items():
        if "hyperlink" in rel.reltype:
            rels[rel_id] = rel._target

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # Keressük a hyperlinkeket az XML-ben
        urls_in_para = []
        for hl in para._element.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hyperlink"):
            r_id = hl.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            if r_id and r_id in rels:
                urls_in_para.append(rels[r_id])

        if urls_in_para:
            lines.append(f"{text} [URL: {urls_in_para[0]}]")
        else:
            lines.append(text)

    return "\n\n".join(lines)

# ── Claude API feldolgozás ────────────────────────────────────────────────────
SYSTEM_PROMPT = """Te egy sajtófigyelő feldolgozó vagy. A PR ügynökség által küldött
Word dokumentum már strukturált: szekciókra és alkategóriákra van osztva.

A dokumentum felépítése:
- Főszekciók: "Prohumanra vonatkozó megjelenések", "Munkaerőpiaci hírek és elemzések",
  "Versenytársakra vonatkozó megjelenések"
- Alkategóriák: pl. "Trenkwalder", "Randstad", "WHC", "Munkaerő-kölcsönzés" stb.
- Minden cikknél: Cím (néha linkként [Cím](URL) formátumban, vagy [URL: ...] jelöléssel),
  forrás és dátum zárójelben: "(Médium, 2026.03.20.)", majd összefoglaló szöveg

A fokategória meghatározása:
- "prohuman": ha a "Prohumanra vonatkozó megjelenések" szekcióban van
- "versenytars": ha a "Versenytársakra vonatkozó megjelenések" szekcióban van
- "piac": ha a "Munkaerőpiaci hírek és elemzések" szekcióban van

Nyerd ki az összes cikket és add vissza KIZÁRÓLAG érvényes JSON-ban:
{"generated":"YYYY-MM-DD","total":42,"cikkek":[{"cim":"...","url":"https://...","forras":"...","datum":"2026-03-20","fokategoria":"piac","alkategoria":"...","osszefoglalo":"...","kulcsszavak":["..."],"het":"2026-03-16","id":"ART-0001"}]}

SZABÁLYOK:
1. cim és osszefoglalo: KIZÁRÓLAG sima szöveg, semmi markdown (#, *, [, ])
2. Ha nincs URL, értéke null
3. het: az adott dátum hetének hétfője YYYY-MM-DD formátumban
4. Ha "Nem volt releváns megjelenés" szerepel egy alkategóriánál, azt hagyd ki
5. Kulcsszavak: 2-5 releváns kulcsszó""".strip()

def classify_with_claude(text):
    today = datetime.date.today().isoformat()
    r = requests.post("https://api.anthropic.com/v1/messages",
        headers={"x-api-key":ANTHROPIC_API_KEY,"anthropic-version":"2023-06-01","content-type":"application/json"},
        json={"model":"claude-opus-4-5","max_tokens":8192,"system":SYSTEM_PROMPT,
              "messages":[{"role":"user","content":f"Mai dátum: {today}\n\n{text[:40000]}"]},
        timeout=180)
    r.raise_for_status()
    raw = r.json()["content"][0]["text"]
    m = re.search(r"\{[\s\S]+\}", raw)
    if not m: raise ValueError(f"Nincs JSON:\n{raw[:300]}")
    return json.loads(m.group())

# ── Adattisztítás ─────────────────────────────────────────────────────────────
def clean_text(s):
    if not s: return s
    s = re.sub(r'^#+\s*', '', s.strip())
    s = re.sub(r'^\[', '', s)
    return s.strip()

def clean_url(url):
    if not url: return ""
    url = url.strip()
    return url if url.startswith(("http://","https://")) else ""

def source_from_url(url):
    """Forrás neve domain alapján."""
    domain_map = {
        'hrportal.hu': 'HR Portál', 'portfolio.hu': 'Portfolio',
        'penzcentrum.hu': 'Pénzcentrum', 'vg.hu': 'Világgazdaság',
        'economx.hu': 'Economx', 'origo.hu': 'Origo',
        'hvg.hu': 'HVG', 'index.hu': 'Index',
        'autopro.hu': 'Autopro', 'hrpwr.hu': 'HRPWR',
        'blikk.hu': 'Blikk', 'forbes.hu': 'Forbes',
        'magyarnemzet.hu': 'Magyar Nemzet', 'telex.hu': 'Telex',
        'mmonline.hu': 'Magyar Munkáltatók Online',
        'dehir.hu': 'Dehir', 'profitline.hu': 'Profitline',
        'turizmus.com': 'Turizmus.com', 'uzletem.hu': 'Üzletem',
        'baon.hu': 'Baon', 'vehir.hu': 'Vehír',
        'trademagazin.hu': 'Trademagazin', 'storeinsider.hu': 'Store Insider',
        'behaviour.hu': 'Behaviour', 'piacesprofit.hu': 'Piac és Profit',
    }
    try:
        domain = urlparse(url).netloc.replace('www.', '')
        return domain_map.get(domain, domain)
    except Exception:
        return ""

def sanitize_data(data):
    for i, c in enumerate(data.get("cikkek", []), 1):
        c["cim"]          = clean_text(c.get("cim", ""))
        c["osszefoglalo"] = clean_text(c.get("osszefoglalo", ""))
        c["url"]          = clean_url(c.get("url", ""))
        # Ha nincs forrás vagy Ismeretlen, próbáljuk URL-ből
        if (not c.get("forras") or c.get("forras") == "Ismeretlen") and c.get("url"):
            c["forras"] = source_from_url(c["url"]) or "Ismeretlen"
        if not c.get("id"):
            c["id"] = f"ART-{i:04d}"
    data["total"] = len(data.get("cikkek", []))
    if not data.get("generated"):
        data["generated"] = datetime.date.today().isoformat()
    return data

# ── HTML frissítése ───────────────────────────────────────────────────────────
HU_MONTHS = {1:"január",2:"február",3:"március",4:"április",5:"május",6:"június",
             7:"július",8:"augusztus",9:"szeptember",10:"október",11:"november",12:"december"}

def week_label(date_str):
    try:
        d = datetime.date.fromisoformat(date_str)
        mon = d - datetime.timedelta(days=d.weekday())
        sun = mon + datetime.timedelta(days=6)
        wn  = mon.isocalendar()[1]
        def f(dt): return f"{dt.year}. {HU_MONTHS[dt.month]} {dt.day}."
        return f"{wn}. hét ({f(mon)} – {f(sun)})"
    except Exception:
        return date_str

def update_html(data):
    if not OUTPUT_FILE.exists():
        raise FileNotFoundError("Nem találom a docs/index.html template-et.")
    html = OUTPUT_FILE.read_text(encoding="utf-8")

    db_json = json.dumps(data, ensure_ascii=False, separators=(',', ':'))

    # DB konstans cseréje
    start_idx = html.find('const DB = {')
    pos = start_idx + len('const DB = ')
    depth, end_idx = 0, pos
    for i, ch in enumerate(html[pos:]):
        if ch == '{': depth += 1
        elif ch == '}':
            depth -= 1
            if depth == 0:
                end_idx = pos + i + 2
                break
    html = html[:start_idx] + f'const DB = {db_json};' + html[end_idx:]

    # Hét legördülő frissítése
    wl = {w: week_label(w) for w in sorted(
        set(c["het"] for c in data["cikkek"] if c.get("het")), reverse=True)}
    opts = '\n      '.join(f'<option value="{w}">{wl[w]}</option>' for w in wl)
    html = re.sub(
        r'(<option value="">Összes hét</option>).*?(</select>)',
        f'\\1\n      {opts}\n    \\2', html, flags=re.DOTALL)

    OUTPUT_FILE.write_text(html, encoding="utf-8")

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    print("🔐 Graph token...")
    token = get_graph_token()

    print("📧 Email letöltése...")
    subject, docx = fetch_latest_word_attachment(token)

    print("📄 Word kinyerése (linkekkel)...")
    text = extract_docx_with_links(docx)
    print(f"   {len(text)} karakter")

    print("🤖 Claude feldolgozás...")
    data = classify_with_claude(text)

    print("🧹 Adattisztítás...")
    data = sanitize_data(data)
    no_url = sum(1 for c in data["cikkek"] if not c.get("url"))
    print(f"   {data['total']} cikk ({no_url} URL nélkül)")

    print("🖥️  HTML frissítése...")
    OUTPUT_FILE.parent.mkdir(exist_ok=True)
    update_html(data)
    print(f"✅ Kész: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
