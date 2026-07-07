#!/usr/bin/env python3
"""
Sajtófigyelő – Cowork pipeline runner
======================================
A Cowork scheduled task ezt a scriptet használja. Három alparancs:

  scan                      – feldolgozatlan .docx fájlok listája az incoming/ mappából
  extract <docx>            – Word → szöveg + hyperlinkek → /tmp/sajto_input.txt
  apply <json> <docx_név>   – új cikkek merge-elése a repo DB-jébe, HTML frissítés,
                              git commit + push, fájl megjelölése feldolgozottként

A repót minden futásnál frissen klónozza /tmp alá (a mountolt mappában a git
nem működik). A PAT a projektmappa github_pat.txt fájljából jön.

Elvárt JSON az apply-hoz (CSAK az új cikkek):
{"cikkek":[{"cim":"...","url":"https://... vagy null","forras":"...",
  "datum":"YYYY-MM-DD","fokategoria":"prohuman|piac|versenytars",
  "alkategoria":"...","osszefoglalo":"...","kulcsszavak":["..."]}]}
(het és id mezőket a script számolja, nem kell megadni)
"""

import sys
import re
import json
import datetime
import subprocess
import shutil
from pathlib import Path
from urllib.parse import urlparse

# ── Útvonalak (a script helyéből számolva — session-független) ────────────────

BASE      = Path(__file__).resolve().parent.parent   # a Sajtófigyelés mappa
INCOMING  = BASE / "incoming"
STATE     = BASE / "state" / "processed.json"
PAT_FILE  = BASE / "github_pat.txt"

REPO_URL  = "github.com/t0thgellert/sajto.git"
REPO_DIR  = Path("/tmp/sajto_work")
INPUT_TXT = Path("/tmp/sajto_input.txt")

DOMAIN_MAP: dict[str, str] = {
    "hrportal.hu": "HR Portál", "portfolio.hu": "Portfolio",
    "penzcentrum.hu": "Pénzcentrum", "vg.hu": "Világgazdaság",
    "economx.hu": "Economx", "origo.hu": "Origo", "hvg.hu": "HVG",
    "index.hu": "Index", "autopro.hu": "Autopro", "hrpwr.hu": "HRPWR",
    "blikk.hu": "Blikk", "forbes.hu": "Forbes", "magyarnemzet.hu": "Magyar Nemzet",
    "telex.hu": "Telex", "mmonline.hu": "MM Online", "dehir.hu": "Dehir",
    "profitline.hu": "Profitline", "turizmus.com": "Turizmus.com",
    "uzletem.hu": "Üzletem", "baon.hu": "Baon", "vehir.hu": "Vehír",
    "trademagazin.hu": "Trade Magazin", "storeinsider.hu": "Store Insider",
    "behaviour.hu": "Behaviour", "piacesprofit.hu": "Piac és Profit",
    "magyarhirlap.hu": "Magyar Hírlap", "mandiner.hu": "Mandiner",
    "g7.hu": "G7", "oeconomus.hu": "Oeconomus", "nepszava.hu": "Népszava",
    "privatbankar.hu": "Privátbankár", "demokrata.hu": "Magyar Demokrata",
    "magyarmezogazdasag.hu": "Magyar Mezőgazdaság", "magro.hu": "Magro.hu",
    "napi.hu": "Napi", "mfor.hu": "MFOR", "adozona.hu": "Adózóna",
    "bama.hu": "BAMA", "novekedes.hu": "Növekedés.hu",
    "vasmegye.hu": "Vasmegye.hu", "magyarkurir.hu": "Magyar Kurír",
}

HU_MONTHS = {
    1: "január", 2: "február", 3: "március", 4: "április",
    5: "május", 6: "június", 7: "július", 8: "augusztus",
    9: "szeptember", 10: "október", 11: "november", 12: "december",
}

# ── Word kinyerés ─────────────────────────────────────────────────────────────

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_HL   = f"{{{_W_NS}}}hyperlink"
_RID  = f"{{{_R_NS}}}id"


def extract_docx(path: str) -> str:
    from docx import Document
    doc = Document(path)
    rels = {rid: rel._target for rid, rel in doc.part.rels.items()
            if "hyperlink" in rel.reltype}
    lines = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        url = next((rels[hl.get(_RID)] for hl in para._element.iter(_HL)
                    if hl.get(_RID) in rels), None)
        lines.append(f"{text} [URL: {url}]" if url else text)
    return "\n\n".join(lines)

# ── Tisztítás ─────────────────────────────────────────────────────────────────

def _clean(s):
    if not s:
        return ""
    s = re.sub(r"^#+\s*", "", s.strip())
    s = re.sub(r"^\[+", "", s)
    s = re.sub(r"\*+", "", s)
    return s.strip()


def _clean_url(u):
    if not u:
        return None
    u = u.strip()
    return u if u.startswith(("http://", "https://")) else None


def _source(url):
    if not url:
        return ""
    try:
        d = urlparse(url).netloc.replace("www.", "")
        return DOMAIN_MAP.get(d, d)
    except Exception:
        return ""


def _monday(date_str):
    try:
        d = datetime.date.fromisoformat(date_str)
        return (d - datetime.timedelta(days=d.weekday())).isoformat()
    except Exception:
        return date_str


def _week_label(monday):
    try:
        mon = datetime.date.fromisoformat(monday)
        sun = mon + datetime.timedelta(days=6)
        wn = mon.isocalendar()[1]
        def f(d): return f"{d.year}. {HU_MONTHS[d.month]} {d.day}."
        return f"{wn}. hét ({f(mon)} – {f(sun)})"
    except Exception:
        return monday

# ── Git ───────────────────────────────────────────────────────────────────────

def _pat() -> str:
    if not PAT_FILE.exists():
        raise RuntimeError(f"Nincs PAT: {PAT_FILE} — hozd létre a tokennel.")
    return PAT_FILE.read_text(encoding="utf-8").strip()


def _run(cmd, cwd=None):
    r = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError(f"$ {' '.join(cmd)}\n{r.stdout}{r.stderr}")
    return r.stdout


def clone_repo() -> Path:
    if REPO_DIR.exists():
        shutil.rmtree(REPO_DIR)
    _run(["git", "clone", "--depth", "1",
          f"https://t0thgellert:{_pat()}@{REPO_URL}", str(REPO_DIR)])
    _run(["git", "config", "user.name", "sajtofigyelo-bot"], cwd=REPO_DIR)
    _run(["git", "config", "user.email", "toth.gellert@prohuman.hu"], cwd=REPO_DIR)
    return REPO_DIR

# ── DB merge + HTML ───────────────────────────────────────────────────────────

def read_db(html: str) -> tuple[dict, int, int]:
    start = html.find("const DB = {")
    if start == -1:
        raise ValueError("Nem találom a 'const DB = {' blokkot.")
    pos = start + len("const DB = ")
    depth = 0
    for i, ch in enumerate(html[pos:]):
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                end = pos + i + 1
                break
    return json.loads(html[pos:end]), start, end


def _keys(c) -> set[str]:
    """Egy cikk összes azonosító kulcsa: URL és cím+dátum is."""
    ks = {f"{_clean(c.get('cim','')).lower()}|{c.get('datum','')}"}
    if c.get("url"):
        ks.add(c["url"].strip().rstrip("/"))
    return ks


def merge(db: dict, new_articles: list[dict]) -> tuple[dict, int]:
    existing: set[str] = set()
    for c in db["cikkek"]:
        existing |= _keys(c)
    max_id = 0
    for c in db["cikkek"]:
        m = re.match(r"ART-(\d+)", c.get("id", ""))
        if m:
            max_id = max(max_id, int(m.group(1)))
    added = 0
    for c in new_articles:
        c["cim"] = _clean(c.get("cim"))
        c["osszefoglalo"] = _clean(c.get("osszefoglalo"))
        c["url"] = _clean_url(c.get("url"))
        if not c.get("forras") or c["forras"] in ("Ismeretlen", "Médium"):
            c["forras"] = _source(c["url"]) or "Ismeretlen"
        if c.get("datum"):
            c["het"] = _monday(c["datum"])
        if _keys(c) & existing:
            continue
        max_id += 1
        c["id"] = f"ART-{max_id:04d}"
        db["cikkek"].append(c)
        existing |= _keys(c)
        added += 1
    db["total"] = len(db["cikkek"])
    db["generated"] = datetime.date.today().isoformat()
    return db, added


def write_html(repo: Path, db: dict) -> None:
    out = repo / "docs" / "index.html"
    html = out.read_text(encoding="utf-8")
    _, start, end = read_db(html)
    db_json = json.dumps(db, ensure_ascii=False, separators=(",", ":"))
    html = html[:start] + f"const DB = {db_json}" + html[end:]
    # A hét-legördülőt a kliensoldali JS építi a DB-ből — itt nem nyúlunk hozzá.
    out.write_text(html, encoding="utf-8")

# ── State ─────────────────────────────────────────────────────────────────────

def _load_state() -> dict:
    if STATE.exists():
        return json.loads(STATE.read_text(encoding="utf-8"))
    return {"processed": []}


def _save_state(state: dict) -> None:
    STATE.parent.mkdir(parents=True, exist_ok=True)
    STATE.write_text(json.dumps(state, ensure_ascii=False, indent=2),
                     encoding="utf-8")

# ── Alparancsok ───────────────────────────────────────────────────────────────

def cmd_scan():
    INCOMING.mkdir(parents=True, exist_ok=True)
    done = set(_load_state()["processed"])
    todo = sorted(p.name for p in INCOMING.glob("*.doc*")
                  if p.name not in done and not p.name.startswith("~"))
    if not todo:
        print("NINCS_UJ")
    for name in todo:
        print(name)


def cmd_extract(docx_name: str):
    path = INCOMING / docx_name
    if not path.exists():
        path = Path(docx_name)  # abszolút út is elfogadott
    text = extract_docx(str(path))
    INPUT_TXT.write_text(
        f"Mai dátum: {datetime.date.today()}\n\n{text}", encoding="utf-8")
    print(f"OK {INPUT_TXT} ({len(text)} karakter)")


def cmd_apply(json_path: str, docx_name: str):
    new = json.loads(Path(json_path).read_text(encoding="utf-8"))
    articles = new["cikkek"] if isinstance(new, dict) else new

    repo = clone_repo()
    html = (repo / "docs" / "index.html").read_text(encoding="utf-8")
    db, _, _ = read_db(html)
    before = len(db["cikkek"])
    db, added = merge(db, articles)
    print(f"DB: {before} cikk + {added} új = {len(db['cikkek'])}")

    if added == 0:
        print("Nincs új cikk — nincs commit.")
    else:
        write_html(repo, db)
        _run(["git", "add", "docs/index.html"], cwd=repo)
        _run(["git", "commit", "-m",
              f"Sajtófigyelő frissítés: {docx_name} (+{added} cikk)"], cwd=repo)
        _run(["git", "push"], cwd=repo)
        print("PUSH OK")

    state = _load_state()
    if docx_name not in state["processed"]:
        state["processed"].append(docx_name)
    state["last_run"] = datetime.datetime.now().isoformat(timespec="seconds")
    _save_state(state)
    print("KESZ")


if __name__ == "__main__":
    a = sys.argv[1:]
    try:
        if a and a[0] == "scan":
            cmd_scan()
        elif a and a[0] == "extract" and len(a) == 2:
            cmd_extract(a[1])
        elif a and a[0] == "apply" and len(a) == 3:
            cmd_apply(a[1], a[2])
        else:
            print(__doc__)
            sys.exit(1)
    except Exception as e:
        print(f"HIBA: {e}", file=sys.stderr)
        sys.exit(1)
