#!/usr/bin/env python3
"""
Sajtófigyelő – automatikus feldolgozó
======================================
Működés:
  1. Microsoft Graph API → legújabb Word melléklet letöltése Outlookból
  2. Word XML → szöveg + hyperlinkek kinyerése
  3. Claude API → strukturált JSON
  4. Adattisztítás + forrás-azonosítás URL alapján
  5. docs/index.html frissítése (DB + hét-legördülő)

Környezeti változók (GitHub Secrets):
  PR_SENDER_EMAIL    – a PR ügynökség emailcíme
  MS_CLIENT_ID       – Azure App Registration Client ID
  MS_CLIENT_SECRET   – Azure App Registration Client Secret
  MS_TENANT_ID       – Azure / Entra Tenant ID
  MS_USER_EMAIL      – a postafiók email-je (akinek az Outlookját olvassuk)
  ANTHROPIC_API_KEY  – Anthropic API kulcs
"""

import os
import re
import io
import json
import base64
import datetime
import time
import sys
from pathlib import Path
from urllib.parse import urlparse

import requests
from docx import Document
from lxml import etree  # noqa: F401 – indirekt használat a docx-on belül

# ── Konfiguráció ──────────────────────────────────────────────────────────────

# Tárgy-alapú szűrés: a feladó BÁRKI lehet (közvetlen vagy továbbított levél)
SUBJECT_KEYWORD  = os.environ.get("SUBJECT_KEYWORD", "piaci körkép")
SENDER_EMAIL     = os.environ.get("PR_SENDER_EMAIL", "")  # már nem kötelező
GRAPH_CLIENT_ID  = os.environ["MS_CLIENT_ID"]
GRAPH_CLIENT_SEC = os.environ["MS_CLIENT_SECRET"]
GRAPH_TENANT_ID  = os.environ["MS_TENANT_ID"]
GRAPH_USER_EMAIL = os.environ["MS_USER_EMAIL"]
ANTHROPIC_KEY    = os.environ["ANTHROPIC_API_KEY"]
OUTPUT_FILE      = Path("docs/index.html")

# Hány legfrissebb emailt vizsgálunk (ha az első nem tartalmaz .docx mellékletet)
EMAIL_SEARCH_TOP = 25

# Domain → olvasható forrás név
DOMAIN_MAP: dict[str, str] = {
    "hrportal.hu":           "HR Portál",
    "portfolio.hu":          "Portfolio",
    "penzcentrum.hu":        "Pénzcentrum",
    "vg.hu":                 "Világgazdaság",
    "economx.hu":            "Economx",
    "origo.hu":              "Origo",
    "hvg.hu":                "HVG",
    "index.hu":              "Index",
    "autopro.hu":            "Autopro",
    "hrpwr.hu":              "HRPWR",
    "blikk.hu":              "Blikk",
    "forbes.hu":             "Forbes",
    "magyarnemzet.hu":       "Magyar Nemzet",
    "telex.hu":              "Telex",
    "mmonline.hu":           "MM Online",
    "dehir.hu":              "Dehir",
    "profitline.hu":         "Profitline",
    "turizmus.com":          "Turizmus.com",
    "uzletem.hu":            "Üzletem",
    "baon.hu":               "Baon",
    "vehir.hu":              "Vehír",
    "trademagazin.hu":       "Trade Magazin",
    "storeinsider.hu":       "Store Insider",
    "behaviour.hu":          "Behaviour",
    "piacesprofit.hu":       "Piac és Profit",
    "magyarhirlap.hu":       "Magyar Hírlap",
    "mandiner.hu":           "Mandiner",
    "g7.hu":                 "G7",
    "oeconomus.hu":          "Oeconomus",
    "nepszava.hu":           "Népszava",
    "privatbankar.hu":       "Privátbankár",
    "demokrata.hu":          "Magyar Demokrata",
    "magyarmezogazdasag.hu": "Magyar Mezőgazdaság",
    "magro.hu":              "Magro.hu",
    "napi.hu":               "Napi",
}

HU_MONTHS = {
    1: "január",   2: "február",  3: "március",  4: "április",
    5: "május",    6: "június",   7: "július",   8: "augusztus",
    9: "szeptember", 10: "október", 11: "november", 12: "december",
}

SYSTEM_PROMPT = """Te egy sajtófigyelő feldolgozó vagy. A PR ügynökség által küldött
Word dokumentum már strukturált: szekciókra és alkategóriákra van osztva.

A dokumentum felépítése:
- Főszekciók: "Prohumanra vonatkozó megjelenések", "Munkaerőpiaci hírek és elemzések",
  "Versenytársakra vonatkozó megjelenések"
- Alkategóriák: pl. "Trenkwalder", "Randstad", "WHC", "Munkaerő-kölcsönzés" stb.
- Minden cikknél: Cím (néha [URL: ...] jelöléssel), forrás és dátum zárójelben:
  "(Médium, 2026.03.20.)", majd összefoglaló szöveg

A főkategória meghatározása a szekció alapján:
- "prohuman"   → "Prohumanra vonatkozó megjelenések"
- "versenytars" → "Versenytársakra vonatkozó megjelenések"
- "piac"       → "Munkaerőpiaci hírek és elemzések"

Nyerd ki az összes cikket és add vissza KIZÁRÓLAG érvényes JSON-ban,
más szöveg, magyarázat vagy markdown nélkül:

{"generated":"YYYY-MM-DD","total":0,"cikkek":[
  {"cim":"...","url":"https://...","forras":"...","datum":"2026-03-20",
   "fokategoria":"piac","alkategoria":"...","osszefoglalo":"...",
   "kulcsszavak":["..."],"het":"2026-03-17","id":"ART-0001"}
]}

SZABÁLYOK:
1. cim és osszefoglalo: CSAK sima szöveg – semmi markdown (#, *, [, ], **)
2. url: ha nincs link, értéke null
3. het: az adott dátum hetének hétfője YYYY-MM-DD formátumban
4. Ha "Nem volt releváns megjelenés" szerepel egy alkategóriánál → hagyd ki
5. kulcsszavak: 2–5 releváns magyar kulcsszó
6. id: ART-0001 formátum, sorszám szerint növekvő""".strip()


# ── Microsoft Graph ───────────────────────────────────────────────────────────

def _graph_token() -> str:
    """OAuth2 client-credentials token kérése."""
    r = requests.post(
        f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}/oauth2/v2.0/token",
        data={
            "grant_type":    "client_credentials",
            "client_id":     GRAPH_CLIENT_ID,
            "client_secret": GRAPH_CLIENT_SEC,
            "scope":         "https://graph.microsoft.com/.default",
        },
        timeout=30,
    )
    r.raise_for_status()
    return r.json()["access_token"]


def _graph_get(token: str, path: str, **kwargs) -> dict:
    r = requests.get(
        f"https://graph.microsoft.com/v1.0{path}",
        headers={"Authorization": f"Bearer {token}"},
        timeout=30,
        **kwargs,
    )
    r.raise_for_status()
    return r.json()


def fetch_latest_docx(token: str) -> tuple[str, bytes]:
    """
    Visszaadja a legújabb .docx melléklet (subject, bytes) párját.

    TÁRGY-ALAPÚ keresés: a legutóbbi EMAIL_SEARCH_TOP emailt vizsgálja,
    és azt fogadja el, amelynek tárgya tartalmazza a SUBJECT_KEYWORD-öt.
    Így működik közvetlen küldésnél ÉS továbbított (FW:) levélnél is,
    függetlenül attól, ki a feladó.
    """
    msgs = _graph_get(token, f"/users/{GRAPH_USER_EMAIL}/messages", params={
        "$orderby": "receivedDateTime desc",
        "$top":     str(EMAIL_SEARCH_TOP),
        "$select":  "id,subject,receivedDateTime,hasAttachments",
    }).get("value", [])

    keyword = SUBJECT_KEYWORD.lower()
    msgs = [m for m in msgs if keyword in (m.get("subject") or "").lower()]

    for msg in msgs:
        if not msg.get("hasAttachments"):
            continue
        atts = _graph_get(
            token,
            f"/users/{GRAPH_USER_EMAIL}/messages/{msg['id']}/attachments",
        ).get("value", [])
        for att in atts:
            if att.get("name", "").lower().endswith((".docx", ".doc")):
                print(f"  ✓ Melléklet: {att['name']}  ({msg['subject']})")
                return msg["subject"], base64.b64decode(att["contentBytes"])

    raise RuntimeError(
        f"Nem találtam '{SUBJECT_KEYWORD}' tárgyú emailt .docx melléklettel "
        f"a legutóbbi {EMAIL_SEARCH_TOP} levél között."
    )


# ── Word kinyerés ─────────────────────────────────────────────────────────────

_W_NS  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R_NS  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_HL_TAG = f"{{{_W_NS}}}hyperlink"
_RID    = f"{{{_R_NS}}}id"


def extract_docx(docx_bytes: bytes) -> str:
    """
    Kinyeri a szöveget és a hyperlinkeket a Word dokumentumból.
    Minden olyan bekezdésnél, ahol link is van: 'Szöveg [URL: https://...]'
    """
    doc = Document(io.BytesIO(docx_bytes))

    # Relationships: rid → url
    rels = {
        rid: rel._target
        for rid, rel in doc.part.rels.items()
        if "hyperlink" in rel.reltype
    }

    lines: list[str] = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        # Első hyperlink az adott bekezdésben
        url = next(
            (rels[hl.get(_RID)] for hl in para._element.iter(_HL_TAG)
             if hl.get(_RID) in rels),
            None,
        )
        lines.append(f"{text} [URL: {url}]" if url else text)

    return "\n\n".join(lines)


# ── Claude API ────────────────────────────────────────────────────────────────

def call_claude(text: str, retries: int = 3) -> dict:
    """
    Elküldi a szöveget a Claude API-nak, visszaadja a parsed JSON-t.
    Automatikusan újrapróbál rate-limit vagy szerver hiba esetén.
    """
    payload = {
        "model":      "claude-sonnet-4-6",   # gyorsabb + olcsóbb mint Opus
        "max_tokens": 8192,
        "system":     SYSTEM_PROMPT,
        "messages":   [{"role": "user", "content":
                         f"Mai dátum: {datetime.date.today()}\n\n{text[:40000]}"}],
    }
    headers = {
        "x-api-key":         ANTHROPIC_KEY,
        "anthropic-version": "2023-06-01",
        "content-type":      "application/json",
    }

    for attempt in range(1, retries + 1):
        try:
            r = requests.post(
                "https://api.anthropic.com/v1/messages",
                headers=headers,
                json=payload,
                timeout=180,
            )
            if r.status_code == 529 or r.status_code >= 500:
                raise requests.HTTPError(response=r)
            r.raise_for_status()
            break
        except requests.HTTPError as e:
            if attempt == retries:
                raise
            wait = 15 * attempt
            print(f"  ⚠ API hiba ({e.response.status_code}), várok {wait}s…")
            time.sleep(wait)

    raw = r.json()["content"][0]["text"].strip()
    # JSON fence eltávolítása ha van
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    m = re.search(r"\{[\s\S]+\}", raw)
    if not m:
        raise ValueError(f"Claude nem adott vissza JSON-t:\n{raw[:400]}")
    return json.loads(m.group())


# ── Adattisztítás ─────────────────────────────────────────────────────────────

def _clean_text(s: str | None) -> str:
    if not s:
        return ""
    s = re.sub(r"^#+\s*", "", s.strip())   # markdown heading
    s = re.sub(r"^\[+",   "", s)           # nyitó szögletes zárójelek
    s = re.sub(r"\*+",    "", s)           # bold/italic csillagok
    return s.strip()


def _clean_url(url: str | None) -> str | None:
    if not url:
        return None
    url = url.strip()
    return url if url.startswith(("http://", "https://")) else None


def _source_from_url(url: str | None) -> str:
    if not url:
        return ""
    try:
        domain = urlparse(url).netloc.replace("www.", "")
        return DOMAIN_MAP.get(domain, domain)
    except Exception:
        return ""


def _week_monday(date_str: str) -> str:
    """Visszaadja a megadott dátum hetének hétfőjét YYYY-MM-DD formátumban."""
    try:
        d = datetime.date.fromisoformat(date_str)
        return (d - datetime.timedelta(days=d.weekday())).isoformat()
    except Exception:
        return date_str


def sanitize(data: dict) -> dict:
    """Megtisztítja és kiegészíti a Claude által visszaadott adatokat."""
    for i, c in enumerate(data.get("cikkek", []), 1):
        c["cim"]          = _clean_text(c.get("cim"))
        c["osszefoglalo"] = _clean_text(c.get("osszefoglalo"))
        c["url"]          = _clean_url(c.get("url"))
        # Forrás: ha hiányzik vagy Ismeretlen, próbálja URL-ből
        if not c.get("forras") or c["forras"] in ("Ismeretlen", "Médium"):
            c["forras"] = _source_from_url(c["url"]) or "Ismeretlen"
        # het mező: biztosan hétfő legyen
        if c.get("datum"):
            c["het"] = _week_monday(c["datum"])
        # id fallback
        if not c.get("id"):
            c["id"] = f"ART-{i:04d}"
    data["total"]     = len(data.get("cikkek", []))
    data["generated"] = data.get("generated") or datetime.date.today().isoformat()
    return data


# ── HTML frissítés ────────────────────────────────────────────────────────────

def _week_label(monday: str) -> str:
    """'2026-03-16' → '11. hét (2026. március 16. – 2026. március 22.)'"""
    try:
        mon = datetime.date.fromisoformat(monday)
        sun = mon + datetime.timedelta(days=6)
        wn  = mon.isocalendar()[1]
        def fmt(d: datetime.date) -> str:
            return f"{d.year}. {HU_MONTHS[d.month]} {d.day}."
        return f"{wn}. hét ({fmt(mon)} – {fmt(sun)})"
    except Exception:
        return monday


def _replace_db_block(html: str, db_json: str) -> str:
    """Lecseréli a const DB = {...}; blokkot brace-mélység alapján."""
    marker = "const DB = {"
    start  = html.find(marker)
    if start == -1:
        raise ValueError("Nem találom a 'const DB = {' blokkot az index.html-ben.")
    pos   = start + len("const DB = ")
    depth = 0
    end   = pos
    for i, ch in enumerate(html[pos:]):
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                end = pos + i + 2  # + '}'  + ';'
                break
    return html[:start] + f"const DB = {db_json};" + html[end:]


def _rebuild_week_dropdown(html: str, data: dict) -> str:
    """Újragenerálja a hét-legördülő <option> elemeit."""
    weeks = sorted(
        {c["het"] for c in data["cikkek"] if c.get("het")},
        reverse=True,
    )
    opts  = "\n      ".join(
        f'<option value="{w}">{_week_label(w)}</option>' for w in weeks
    )
    return re.sub(
        r'(<option value="">Összes hét</option>).*?(</select>)',
        f'\\1\n      {opts}\n    \\2',
        html,
        flags=re.DOTALL,
    )


def update_html(data: dict) -> None:
    if not OUTPUT_FILE.exists():
        raise FileNotFoundError(f"Nem találom: {OUTPUT_FILE}")
    html     = OUTPUT_FILE.read_text(encoding="utf-8")
    db_json  = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    html     = _replace_db_block(html, db_json)
    html     = _rebuild_week_dropdown(html, data)
    OUTPUT_FILE.write_text(html, encoding="utf-8")


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> None:
    print("🔐 Graph token megszerzése…")
    token = _graph_token()

    print("📧 Legújabb email + Word melléklet letöltése…")
    subject, docx_bytes = fetch_latest_docx(token)

    print("📄 Word szöveg + linkek kinyerése…")
    text = extract_docx(docx_bytes)
    print(f"   {len(text):,} karakter")

    print("🤖 Claude feldolgozás…")
    data = call_claude(text)

    print("🧹 Adattisztítás…")
    data    = sanitize(data)
    no_url  = sum(1 for c in data["cikkek"] if not c.get("url"))
    print(f"   {data['total']} cikk  |  {no_url} URL nélkül")

    print("🖥️  HTML frissítése…")
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    update_html(data)
    print(f"✅ Kész → {OUTPUT_FILE}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\n❌ HIBA: {exc}", file=sys.stderr)
        sys.exit(1)
