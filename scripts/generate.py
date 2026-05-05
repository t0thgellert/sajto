#!/usr/bin/env python3
"""
Sajtófigyelő automatikus feldolgozó
Outlook-ból letölti a legújabb Word mellékletet, Claude API-val szortírozza,
majd GitHub Pages-re generál HTML oldalt.
"""

import os
import re
import json
import base64
import datetime
import requests
from pathlib import Path

# ── Config ───────────────────────────────────────────────────────────────────
SENDER_EMAIL      = os.environ["PR_SENDER_EMAIL"]        # pl. sajto@prugnokság.hu
GRAPH_CLIENT_ID   = os.environ["MS_CLIENT_ID"]
GRAPH_CLIENT_SEC  = os.environ["MS_CLIENT_SECRET"]
GRAPH_TENANT_ID   = os.environ["MS_TENANT_ID"]
GRAPH_USER_EMAIL  = os.environ["MS_USER_EMAIL"]          # a kollégád email-je
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
OUTPUT_FILE       = Path("docs/index.html")              # GitHub Pages gyökér

# ── Microsoft Graph auth ──────────────────────────────────────────────────────
def get_graph_token() -> str:
    resp = requests.post(
        f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}/oauth2/v2.0/token",
        data={
            "grant_type":    "client_credentials",
            "client_id":     GRAPH_CLIENT_ID,
            "client_secret": GRAPH_CLIENT_SEC,
            "scope":         "https://graph.microsoft.com/.default",
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def graph_get(token: str, path: str, **kwargs) -> dict:
    resp = requests.get(
        f"https://graph.microsoft.com/v1.0{path}",
        headers={"Authorization": f"Bearer {token}"},
        timeout=30,
        **kwargs,
    )
    resp.raise_for_status()
    return resp.json()


# ── Legfrissebb sajtófigyelő email + Word melléklet ──────────────────────────
def fetch_latest_word_attachment(token: str) -> tuple[str, bytes]:
    """
    Visszaadja: (email tárgya, Word fájl bináris tartalma)
    """
    # Keresés: feladó alapján, legfrissebb 5 levél
    params = {
        "$filter": f"from/emailAddress/address eq '{SENDER_EMAIL}'",
        "$orderby": "receivedDateTime desc",
        "$top": "5",
        "$select": "id,subject,receivedDateTime,hasAttachments",
    }
    messages = graph_get(
        token,
        f"/users/{GRAPH_USER_EMAIL}/messages",
        params=params,
    ).get("value", [])

    for msg in messages:
        if not msg.get("hasAttachments"):
            continue
        attachments = graph_get(
            token,
            f"/users/{GRAPH_USER_EMAIL}/messages/{msg['id']}/attachments",
        ).get("value", [])
        for att in attachments:
            name = att.get("name", "")
            if name.lower().endswith((".docx", ".doc")):
                content_bytes = base64.b64decode(att["contentBytes"])
                print(f"✓ Melléklet megtalálva: {name} ({msg['subject']})")
                return msg["subject"], content_bytes

    raise RuntimeError(
        f"Nem találtam .docx mellékletet {SENDER_EMAIL} legutóbbi 5 levelében."
    )


# ── Word → szöveg ─────────────────────────────────────────────────────────────
def extract_text_from_docx(docx_bytes: bytes) -> str:
    import io
    from docx import Document  # python-docx

    doc = Document(io.BytesIO(docx_bytes))
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return "\n\n".join(paragraphs)


# ── Claude API: szortírozás témák szerint ────────────────────────────────────
SYSTEM_PROMPT = """
Te egy sajtófigyelő elemző vagy. A feladatod:
1. Olvasd el a PR ügynökség által küldött sajtófigyelő szöveget.
2. Azonosítsd a benne szereplő egyedi sajtómegjelenéseket (cikkeket, híreket).
3. Csoportosítsd tematikus kategóriákba (pl. Technológia, Üzlet, HR, Oktatás,
   Egészségügy, Politika, Kultúra, Sport, Egyéb — de csak azokat hozd létre,
   amelyek tényleg előfordulnak a szövegben).
4. Minden megjelenésnél tüntesd fel: forrás neve, cikk/hír rövid összefoglalója
   (1-2 mondat), megjelenés dátuma ha szerepel.

Válaszolj KIZÁRÓLAG érvényes JSON-ban, más szöveg nélkül:
{
  "hét": "2026-W18",
  "összefoglaló": "Rövid 1-2 mondatos összefoglaló az egész hétről.",
  "kategóriák": [
    {
      "név": "Kategória neve",
      "megjelenések_száma": 3,
      "megjelenések": [
        {
          "forrás": "Médium neve",
          "összefoglaló": "...",
          "dátum": "2026-05-05"
        }
      ]
    }
  ]
}
""".strip()


def classify_with_claude(text: str) -> dict:
    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key":         ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
            "content-type":      "application/json",
        },
        json={
            "model":      "claude-opus-4-5",
            "max_tokens": 4096,
            "system":     SYSTEM_PROMPT,
            "messages": [
                {"role": "user", "content": text[:40000]},   # API limit
            ],
        },
        timeout=120,
    )
    resp.raise_for_status()
    raw = resp.json()["content"][0]["text"]

    # JSON kinyerése, ha véletlenül mégis lenne markdown fence
    match = re.search(r"\{[\s\S]+\}", raw)
    if not match:
        raise ValueError(f"Claude nem adott vissza JSON-t:\n{raw[:500]}")
    return json.loads(match.group())


# ── HTML generálás ────────────────────────────────────────────────────────────
CATEGORY_ICONS = {
    "Technológia": "💻",
    "Üzlet":       "📈",
    "HR":          "👥",
    "Oktatás":     "🎓",
    "Egészségügy": "🏥",
    "Politika":    "🏛️",
    "Kultúra":     "🎭",
    "Sport":       "⚽",
    "Egyéb":       "📰",
}
DEFAULT_ICON = "📰"


def _category_cards_html(categories: list) -> str:
    html_parts = []
    for cat in categories:
        icon = CATEGORY_ICONS.get(cat["név"], DEFAULT_ICON)
        count = cat.get("megjelenések_száma", len(cat.get("megjelenések", [])))
        items_html = ""
        for m in cat.get("megjelenések", []):
            date_str = m.get("dátum", "")
            date_badge = (
                f'<span class="item-date">{date_str}</span>' if date_str else ""
            )
            items_html += f"""
              <div class="media-item">
                <div class="media-item-header">
                  <span class="media-source">{m.get('forrás','–')}</span>
                  {date_badge}
                </div>
                <p class="media-summary">{m.get('összefoglaló','')}</p>
              </div>"""

        html_parts.append(f"""
    <section class="category-section">
      <div class="category-header">
        <span class="category-icon">{icon}</span>
        <h2 class="category-title">{cat['név']}</h2>
        <span class="category-badge">{count} megjelenés</span>
      </div>
      <div class="media-list">{items_html}
      </div>
    </section>""")

    return "\n".join(html_parts)


def generate_html(data: dict, email_subject: str) -> str:
    now         = datetime.datetime.now().strftime("%Y. %m. %d. %H:%M")
    week_label  = data.get("hét", "–")
    summary     = data.get("összefoglaló", "")
    categories  = data.get("kategóriák", [])
    total       = sum(c.get("megjelenések_száma", 0) for c in categories)
    cards_html  = _category_cards_html(categories)
    cat_count   = len(categories)

    return f"""<!DOCTYPE html>
<html lang="hu">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Sajtófigyelő — {week_label}</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Open+Sans:wght@300;400;600;700&display=swap" rel="stylesheet">
  <style>
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
    html {{ scroll-behavior: smooth; }}
    body {{
      font-family: 'Open Sans', sans-serif;
      font-size: 16px;
      color: #001D35;
      background: #F4F8FF;
      line-height: 1.6;
    }}

    /* ── Brand tokens ── */
    :root {{
      --blue-dark:  #001D35;
      --blue-ph:    #0245EB;
      --blue-ls:    #66CCFF;
      --grey-light: #DCDDDE;
      --orange:     #FF8715;
      --gradient:   linear-gradient(135deg, #66CCFF 0%, #0245EB 100%);
      --radius-box: 20px;
      --radius-cta: 200px;
    }}

    .container {{ max-width: 1100px; margin: 0 auto; padding: 0 32px; }}

    /* ── Nav ── */
    nav {{
      background: rgba(255,255,255,.95);
      backdrop-filter: blur(8px);
      border-bottom: 1px solid var(--grey-light);
      position: sticky; top: 0; z-index: 100;
    }}
    .nav-inner {{
      display: flex; align-items: center;
      justify-content: space-between; height: 64px;
    }}
    .nav-title {{
      font-size: 15px; font-weight: 700;
      color: var(--blue-dark); text-decoration: none;
    }}
    .nav-week {{
      font-size: 13px; color: #666;
      background: rgba(2,69,235,.07);
      padding: 5px 14px; border-radius: var(--radius-cta);
    }}

    /* ── Hero ── */
    .hero {{
      background: var(--blue-dark);
      padding: 72px 0 64px;
      position: relative; overflow: hidden;
    }}
    .hero::before {{
      content: '';
      position: absolute; right: -80px; top: -80px;
      width: 600px; height: 600px;
      background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='80' height='80' viewBox='0 0 80 80'%3E%3Ccircle cx='40' cy='16' r='10' fill='none' stroke='%2366CCFF' stroke-width='1.5' opacity='.12'/%3E%3Cpath d='M10 70 A30 30 0 0 1 70 70' fill='none' stroke='%2366CCFF' stroke-width='1.5' opacity='.12'/%3E%3C/svg%3E");
      background-repeat: repeat; opacity: .4; pointer-events: none;
    }}
    .hero-tag {{
      display: inline-block;
      font-size: 11px; font-weight: 700; letter-spacing: .12em;
      text-transform: uppercase;
      color: var(--blue-ls);
      background: rgba(102,204,255,.12);
      padding: 5px 14px; border-radius: var(--radius-cta);
      margin-bottom: 20px;
    }}
    .hero h1 {{
      font-size: clamp(32px, 4.5vw, 52px);
      font-weight: 700; line-height: 1.1; color: #fff;
      margin-bottom: 16px;
    }}
    .hero h1 em {{
      font-style: normal;
      background: var(--gradient);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      background-clip: text;
    }}
    .hero-subtitle {{
      font-size: 17px; color: rgba(255,255,255,.7);
      max-width: 620px; line-height: 1.7; margin-bottom: 40px;
    }}
    .hero-stats {{
      display: flex; gap: 40px; flex-wrap: wrap; margin-top: 40px;
    }}
    .hero-stat {{ display: flex; flex-direction: column; }}
    .hero-stat-value {{
      font-size: 32px; font-weight: 700;
      color: var(--blue-ls); line-height: 1;
    }}
    .hero-stat-label {{
      font-size: 12px; color: rgba(255,255,255,.5);
      margin-top: 4px;
    }}

    /* ── Main content ── */
    .main {{ padding: 56px 0 80px; }}

    /* ── Category section ── */
    .category-section {{
      background: #fff;
      border-radius: var(--radius-box);
      padding: 36px 40px;
      margin-bottom: 24px;
      box-shadow: 0 2px 12px rgba(0,29,53,.06);
    }}
    .category-header {{
      display: flex; align-items: center; gap: 14px;
      margin-bottom: 28px;
      padding-bottom: 20px;
      border-bottom: 2px solid #F4F8FF;
    }}
    .category-icon {{ font-size: 26px; }}
    .category-title {{
      font-size: 20px; font-weight: 700;
      color: var(--blue-dark); flex: 1;
    }}
    .category-badge {{
      font-size: 12px; font-weight: 700;
      color: var(--blue-ph);
      background: rgba(2,69,235,.08);
      padding: 4px 12px; border-radius: var(--radius-cta);
    }}

    /* ── Media items ── */
    .media-list {{ display: flex; flex-direction: column; gap: 16px; }}
    .media-item {{
      background: #F8FAFF;
      border-left: 3px solid var(--blue-ph);
      border-radius: 0 10px 10px 0;
      padding: 16px 20px;
      transition: background .15s;
    }}
    .media-item:hover {{ background: #EEF3FF; }}
    .media-item-header {{
      display: flex; align-items: center;
      justify-content: space-between; gap: 12px;
      margin-bottom: 6px;
    }}
    .media-source {{
      font-size: 13px; font-weight: 700;
      color: var(--blue-ph);
    }}
    .item-date {{
      font-size: 12px; color: #888;
      background: #EBEBEB;
      padding: 2px 10px; border-radius: 20px;
      white-space: nowrap;
    }}
    .media-summary {{
      font-size: 14px; color: #444; line-height: 1.65;
    }}

    /* ── Footer ── */
    footer {{
      background: var(--blue-dark);
      padding: 40px 0;
      text-align: center;
    }}
    .footer-note {{
      font-size: 13px; color: rgba(255,255,255,.4);
      margin-top: 8px;
    }}

    /* ── Responsive ── */
    @media (max-width: 640px) {{
      .container {{ padding: 0 16px; }}
      .category-section {{ padding: 24px 20px; }}
      .hero-stats {{ gap: 24px; }}
      .hero {{ padding: 48px 0 40px; }}
    }}
  </style>
</head>
<body>

<nav>
  <div class="container">
    <div class="nav-inner">
      <span class="nav-title">📰 Sajtófigyelő</span>
      <span class="nav-week">{week_label}</span>
    </div>
  </div>
</nav>

<header class="hero">
  <div class="container">
    <div class="hero-tag">Heti sajtófigyelő</div>
    <h1>Sajtómegjelenések<br><em>{week_label}</em></h1>
    <p class="hero-subtitle">{summary}</p>
    <div class="hero-stats">
      <div class="hero-stat">
        <span class="hero-stat-value">{total}</span>
        <span class="hero-stat-label">összes megjelenés</span>
      </div>
      <div class="hero-stat">
        <span class="hero-stat-value">{cat_count}</span>
        <span class="hero-stat-label">témakör</span>
      </div>
    </div>
  </div>
</header>

<main class="main">
  <div class="container">
    {cards_html}
  </div>
</main>

<footer>
  <div class="container">
    <div style="font-size:13px;color:rgba(255,255,255,.6);">
      Forrás: {email_subject}
    </div>
    <div class="footer-note">Automatikusan generálva — {now}</div>
  </div>
</footer>

</body>
</html>"""


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    print("🔐 Microsoft Graph token...")
    token = get_graph_token()

    print("📧 Legújabb sajtófigyelő letöltése...")
    subject, docx_bytes = fetch_latest_word_attachment(token)

    print("📄 Word szöveg kinyerése...")
    text = extract_text_from_docx(docx_bytes)
    print(f"   {len(text)} karakter kinyerve.")

    print("🤖 Claude feldolgozás...")
    data = classify_with_claude(text)
    cat_count = len(data.get("kategóriák", []))
    total     = sum(c.get("megjelenések_száma", 0) for c in data.get("kategóriák", []))
    print(f"   {cat_count} kategória, {total} megjelenés azonosítva.")

    print("🖥️  HTML generálás...")
    OUTPUT_FILE.parent.mkdir(exist_ok=True)
    html = generate_html(data, subject)
    OUTPUT_FILE.write_text(html, encoding="utf-8")
    print(f"✅ Kész: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
