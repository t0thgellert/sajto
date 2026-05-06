#!/usr/bin/env python3
"""
Sajtófigyelő automatikus feldolgozó
Outlook-ból letölti a legújabb Word mellékletet, Claude API-val feldolgozza,
majd GitHub Pages-re generál HTML oldalt.
"""

import os, re, json, base64, datetime, requests
from pathlib import Path

SENDER_EMAIL      = os.environ["PR_SENDER_EMAIL"]
GRAPH_CLIENT_ID   = os.environ["MS_CLIENT_ID"]
GRAPH_CLIENT_SEC  = os.environ["MS_CLIENT_SECRET"]
GRAPH_TENANT_ID   = os.environ["MS_TENANT_ID"]
GRAPH_USER_EMAIL  = os.environ["MS_USER_EMAIL"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
OUTPUT_FILE       = Path("docs/index.html")

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

def extract_text_from_docx(docx_bytes):
    import io
    from docx import Document
    doc = Document(io.BytesIO(docx_bytes))
    return "\n\n".join(p.text.strip() for p in doc.paragraphs if p.text.strip())

def clean_text(s):
    """Eltávolítja a markdown jelölőket (####, [, stb.)"""
    if not s: return s
    s = re.sub(r'^#+\s*', '', s.strip())
    s = re.sub(r'^\[', '', s)
    return s.strip()

def clean_url(url):
    if not url: return ""
    url = url.strip()
    return url if url.startswith(("http://","https://")) else ""

SYSTEM_PROMPT = """Te egy sajtófigyelő elemző vagy. Dolgozd fel a PR ügynökség sajtófigyelő dokumentumát.

Azonosítsd az összes egyedi sajtómegjelenést. Minden cikkhez:
- cim: pontos cím, KIZÁRÓLAG sima szöveg (TILOS: #, ##, *, [, ] és más markdown)
- url: teljes URL https://... előtaggal, vagy null ha nincs
- forras: médium neve
- datum: "YYYY-MM-DD" vagy null
- fokategoria: "prohuman" | "piac" | "versenytars"
- alkategoria: pl. "Beruházások / leépítések", "Bérszínvonal / fizetések", "Vendégmunka", "Hosszú távú trendek", "Munkaerő-kölcsönzés", "Munkaerő-tartalék", "Munkaerőpiac / makrogazdaság", "Prohuman megjelenés", vagy versenytárs neve
- osszefoglalo: 2-3 mondatos összefoglaló, KIZÁRÓLAG sima szöveg (TILOS markdown)
- kulcsszavak: 2-5 kulcsszó lista
- het: a cikk hetének hétfője "YYYY-MM-DD" formátumban
- id: "ART-XXXX" 0001-től számozva

SZABÁLYOK:
1. cim és osszefoglalo mezőkben TILOS markdown (#, *, [, ] stb.)
2. Ha nincs URL, értéke null (nem üres string)
3. URL mindig http:// vagy https:// előtaggal

Válaszolj CSAK érvényes JSON-ban:
{"generated":"YYYY-MM-DD","total":42,"cikkek":[{"cim":"...","url":"https://...","forras":"...","datum":"2026-05-05","fokategoria":"piac","alkategoria":"...","osszefoglalo":"...","kulcsszavak":["..."],"het":"2026-05-05","id":"ART-0001"}]}""".strip()

def classify_with_claude(text):
    today = datetime.date.today().isoformat()
    r = requests.post("https://api.anthropic.com/v1/messages",
        headers={"x-api-key":ANTHROPIC_API_KEY,"anthropic-version":"2023-06-01","content-type":"application/json"},
        json={"model":"claude-opus-4-5","max_tokens":8192,"system":SYSTEM_PROMPT,
              "messages":[{"role":"user","content":f"Mai dátum: {today}\n\n{text[:40000]}"}]},
        timeout=180)
    r.raise_for_status()
    raw = r.json()["content"][0]["text"]
    m = re.search(r"\{[\s\S]+\}", raw)
    if not m: raise ValueError(f"Nincs JSON a válaszban:\n{raw[:300]}")
    return json.loads(m.group())

def sanitize_data(data):
    """Tisztítja a Claude outputot: markdown eltávolítás, URL validálás."""
    for i, c in enumerate(data.get("cikkek",[]), 1):
        c["cim"]          = clean_text(c.get("cim",""))
        c["osszefoglalo"] = clean_text(c.get("osszefoglalo",""))
        c["url"]          = clean_url(c.get("url",""))
        if not c.get("id"): c["id"] = f"ART-{i:04d}"
    data["total"] = len(data.get("cikkek",[]))
    if not data.get("generated"):
        data["generated"] = datetime.date.today().isoformat()
    return data

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
    db_json = json.dumps(data, ensure_ascii=False, separators=(',',':'))
    html = re.sub(r'const DB = \{.*?\};', f'const DB = {db_json};', html, flags=re.DOTALL)
    wl = {w: week_label(w) for w in sorted(set(c["het"] for c in data["cikkek"] if c.get("het")), reverse=True)}
    opts = '\n      '.join(f'<option value="{w}">{wl[w]}</option>' for w in wl)
    html = re.sub(r'(<option value="">Összes hét</option>).*?(</select>)',
                  f'\\1\n      {opts}\n    \\2', html, flags=re.DOTALL)
    OUTPUT_FILE.write_text(html, encoding="utf-8")

def main():
    print("🔐 Graph token...")
    token = get_graph_token()
    print("📧 Email letöltése...")
    subject, docx = fetch_latest_word_attachment(token)
    print("📄 Word kinyerése...")
    text = extract_text_from_docx(docx)
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
