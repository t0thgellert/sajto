#!/usr/bin/env python3
"""
Sajtófigyelő – lokális feldolgozó (CoWork verzió)
==================================================
Ez a script CoWork által futtatva dolgozza fel a Word dokumentumot.
NEM kell Anthropic API kulcs — a CoWork saját Claude-ja végzi a feldolgozást.
NEM kell Microsoft Graph — a CoWork tölti le az emailből a Word fájlt.

Használat:
    python generate_local.py <word_fajl.docx>

A script:
    1. Kinyeri a szöveget + linkeket a Word fájlból
    2. Kiírja a feldolgozandó szöveget egy közbülső fájlba (sajto_input.txt)
    3. A CoWork instrukció alapján Claude feldolgozza → sajto_output.json
    4. Frissíti a docs/index.html-t az új adatokkal

CoWork workflow (COWORK_INSTRUKCIÓ.md) fájlból indul.
"""

import sys
import io
import re
import json
import datetime
from pathlib import Path
from urllib.parse import urlparse

# ── Konfiguráció ──────────────────────────────────────────────────────────────

OUTPUT_HTML   = Path("docs/index.html")
INPUT_TXT     = Path("sajto_input.txt")    # CoWork ezt adja Claude-nak
OUTPUT_JSON   = Path("sajto_output.json")  # CoWork Claude-ja ide írja

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

# ── Word kinyerés ─────────────────────────────────────────────────────────────

_W_NS  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R_NS  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_HL    = f"{{{_W_NS}}}hyperlink"
_RID   = f"{{{_R_NS}}}id"


def extract_docx(path: str) -> str:
    """Kinyeri a szöveget és hyperlinkeket a Word fájlból."""
    from docx import Document
    doc  = Document(path)
    rels = {
        rid: rel._target
        for rid, rel in doc.part.rels.items()
        if "hyperlink" in rel.reltype
    }
    lines = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        url = next(
            (rels[hl.get(_RID)] for hl in para._element.iter(_HL)
             if hl.get(_RID) in rels),
            None,
        )
        lines.append(f"{text} [URL: {url}]" if url else text)
    return "\n\n".join(lines)


# ── Adattisztítás ─────────────────────────────────────────────────────────────

def _clean(s: str | None) -> str:
    if not s:
        return ""
    s = re.sub(r"^#+\s*", "", s.strip())
    s = re.sub(r"^\[+",   "", s)
    s = re.sub(r"\*+",    "", s)
    return s.strip()


def _clean_url(u: str | None) -> str | None:
    if not u:
        return None
    u = u.strip()
    return u if u.startswith(("http://", "https://")) else None


def _source(url: str | None) -> str:
    if not url:
        return ""
    try:
        d = urlparse(url).netloc.replace("www.", "")
        return DOMAIN_MAP.get(d, d)
    except Exception:
        return ""


def _monday(date_str: str) -> str:
    try:
        d = datetime.date.fromisoformat(date_str)
        return (d - datetime.timedelta(days=d.weekday())).isoformat()
    except Exception:
        return date_str


def sanitize(data: dict) -> dict:
    for i, c in enumerate(data.get("cikkek", []), 1):
        c["cim"]          = _clean(c.get("cim"))
        c["osszefoglalo"] = _clean(c.get("osszefoglalo"))
        c["url"]          = _clean_url(c.get("url"))
        if not c.get("forras") or c["forras"] in ("Ismeretlen", "Médium"):
            c["forras"] = _source(c["url"]) or "Ismeretlen"
        if c.get("datum"):
            c["het"] = _monday(c["datum"])
        if not c.get("id"):
            c["id"] = f"ART-{i:04d}"
    data["total"]     = len(data.get("cikkek", []))
    data["generated"] = data.get("generated") or datetime.date.today().isoformat()
    return data


# ── HTML frissítés ────────────────────────────────────────────────────────────

def _week_label(monday: str) -> str:
    try:
        mon = datetime.date.fromisoformat(monday)
        sun = mon + datetime.timedelta(days=6)
        wn  = mon.isocalendar()[1]
        def f(d): return f"{d.year}. {HU_MONTHS[d.month]} {d.day}."
        return f"{wn}. hét ({f(mon)} – {f(sun)})"
    except Exception:
        return monday


def update_html(data: dict) -> None:
    if not OUTPUT_HTML.exists():
        raise FileNotFoundError(f"Nem találom: {OUTPUT_HTML}")
    html = OUTPUT_HTML.read_text(encoding="utf-8")

    # DB blokk csere
    db_json = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    start   = html.find("const DB = {")
    if start == -1:
        raise ValueError("Nem találom a 'const DB = {' blokkot.")
    pos = start + len("const DB = ")
    depth, end = 0, pos
    for i, ch in enumerate(html[pos:]):
        if ch == "{":   depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                end = pos + i + 2
                break
    html = html[:start] + f"const DB = {db_json};" + html[end:]

    # Hét dropdown
    weeks = sorted({c["het"] for c in data["cikkek"] if c.get("het")}, reverse=True)
    opts  = "\n      ".join(
        f'<option value="{w}">{_week_label(w)}</option>' for w in weeks
    )
    html = re.sub(
        r'(<option value="">Összes hét</option>).*?(</select>)',
        f'\\1\n      {opts}\n    \\2',
        html, flags=re.DOTALL,
    )
    OUTPUT_HTML.write_text(html, encoding="utf-8")


# ── Main ──────────────────────────────────────────────────────────────────────

def step1_extract(docx_path: str) -> None:
    """1. lépés: Word → sajto_input.txt (CoWork ezt futtatja először)"""
    print(f"📄 Word kinyerése: {docx_path}")
    text = extract_docx(docx_path)
    INPUT_TXT.write_text(
        f"Mai dátum: {datetime.date.today()}\n\n{text}",
        encoding="utf-8"
    )
    print(f"✅ Mentve → {INPUT_TXT}  ({len(text):,} karakter)")
    print()
    print("⏳ Következő lépés: a CoWork Claude-ja feldolgozza a sajto_input.txt-t")
    print("   és az eredményt sajto_output.json-ba menti.")
    print("   Lásd: COWORK_INSTRUKCIÓ.md")


def step2_update() -> None:
    """2. lépés: sajto_output.json → docs/index.html frissítése"""
    if not OUTPUT_JSON.exists():
        raise FileNotFoundError(
            f"Nem találom: {OUTPUT_JSON}\n"
            "Futtasd előbb a CoWork Claude feldolgozást!"
        )
    print(f"📥 JSON betöltése: {OUTPUT_JSON}")
    data = json.loads(OUTPUT_JSON.read_text(encoding="utf-8"))
    data = sanitize(data)
    no_url = sum(1 for c in data["cikkek"] if not c.get("url"))
    print(f"   {data['total']} cikk  |  {no_url} URL nélkül")

    print("🖥️  HTML frissítése…")
    OUTPUT_HTML.parent.mkdir(parents=True, exist_ok=True)
    update_html(data)
    print(f"✅ Kész → {OUTPUT_HTML}")


if __name__ == "__main__":
    args = sys.argv[1:]

    if not args:
        # Ha nincs argumentum: step2 (JSON → HTML)
        step2_update()
    elif args[0] == "--step1" and len(args) > 1:
        step1_extract(args[1])
    elif args[0].endswith((".docx", ".doc")):
        step1_extract(args[0])
    else:
        print("Használat:")
        print("  python generate_local.py <fajl.docx>   → Word kinyerés")
        print("  python generate_local.py               → JSON → HTML frissítés")
        sys.exit(1)
