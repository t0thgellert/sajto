#!/usr/bin/env python3
"""
Sajtófigyelő – Cowork pipeline runner
======================================
A Cowork scheduled task ezt a scriptet használja. Három alparancs:

  scan                      – feldolgozatlan .docx fájlok listája az incoming/ mappából
  extract <docx>            – Word → szöveg + hyperlinkek → work/sajto_input_<név>.txt
  apply <json> <docx_név>   – új cikkek merge-elése a repo DB-jébe, HTML frissítés,
                              git commit + push, fájl megjelölése feldolgozottként

A repót minden futásnál friss, egyedi ideiglenes mappába klónozza (tempfile) —
a mountolt mappában a git nem működik, a fix /tmp-útvonalak pedig ütköznek a
korábbi sessionök 'nobody' tulajdonú fájljaival. A PAT a projektmappa
github_pat.txt fájljából jön, és a hibaüzenetekből ki van maszkolva.

A köztes fájlok a projektmappa work/ almappájában vannak (docx-enként egyedi
névvel), így a bash és a Windows-oldali fájleszközök is elérik.

Elvárt JSON az apply-hoz (CSAK az új cikkek):
{"cikkek":[{"cim":"...","url":"https://... vagy null","forras":"...",
  "datum":"YYYY-MM-DD","fokategoria":"prohuman|piac|versenytars",
  "alkategoria":"...","osszefoglalo":"...","kulcsszavak":["..."]}]}
(het és id mezőket a script számolja, nem kell megadni)

Validáció az apply-ban: érvénytelen datum vagy fokategoria esetén a cikk
KIMARAD és FIGYELEM sor jelzi — ezek nélkül a dashboard JS-e összeomlana.
"""

import sys
import re
import json
import datetime
import subprocess
import shutil
import tempfile
from pathlib import Path
from urllib.parse import urlparse

# ── Útvonalak (a script helyéből számolva — session-független) ────────────────

BASE      = Path(__file__).resolve().parent.parent   # a Sajtófigyelés mappa
INCOMING  = BASE / "incoming"
STATE     = BASE / "state" / "processed.json"
PAT_FILE  = BASE / "github_pat.txt"
WORK      = BASE / "work"   # köztes fájlok — bash és Windows felől is látszik

REPO_URL  = "github.com/t0thgellert/sajto.git"

VALID_FOKAT = {"prohuman", "piac", "versenytars"}

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

# ── Word kinyerés ─────────────────────────────────────────────────────────────

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_HL   = f"{{{_W_NS}}}hyperlink"
_RID  = f"{{{_R_NS}}}id"
_TAG_P   = f"{{{_W_NS}}}p"
_TAG_TBL = f"{{{_W_NS}}}tbl"


def extract_docx(path: str) -> str:
    """Bekezdések ÉS táblázatok szövege dokumentum-sorrendben, hyperlinkekkel."""
    from docx import Document
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    doc = Document(path)
    rels = {rid: rel._target for rid, rel in doc.part.rels.items()
            if "hyperlink" in rel.reltype}

    def para_line(p_elem):
        para = Paragraph(p_elem, doc)
        text = para.text.strip()
        if not text:
            return None
        url = next((rels[hl.get(_RID)] for hl in p_elem.iter(_HL)
                    if hl.get(_RID) in rels), None)
        return f"{text} [URL: {url}]" if url else text

    lines = []
    for child in doc.element.body.iterchildren():
        if child.tag == _TAG_P:
            line = para_line(child)
            if line:
                lines.append(line)
        elif child.tag == _TAG_TBL:
            table = Table(child, doc)
            for row in table.rows:
                cells = [c.text.strip() for c in row.cells]
                line = " | ".join(c for c in cells if c)
                if line:
                    lines.append(line)
    return "\n\n".join(lines)

# ── Tisztítás / validálás ─────────────────────────────────────────────────────

def _clean(s):
    if not s:
        return ""
    s = re.sub(r"^#+\s*", "", str(s).strip())
    s = re.sub(r"^\[+", "", s)
    s = re.sub(r"\*+", "", s)
    return s.strip()


def _clean_url(u):
    if not u:
        return None
    u = str(u).strip()
    return u if u.startswith(("http://", "https://")) else None


def _source(url):
    if not url:
        return ""
    try:
        d = urlparse(url).netloc.removeprefix("www.")
        return DOMAIN_MAP.get(d, d)
    except Exception:
        return ""


def _valid_date(s):
    """YYYY-MM-DD ISO dátum, vagy None ha érvénytelen."""
    try:
        return datetime.date.fromisoformat(str(s).strip()).isoformat()
    except Exception:
        return None


def _monday(date_str):
    d = datetime.date.fromisoformat(date_str)
    return (d - datetime.timedelta(days=d.weekday())).isoformat()


def _tags(v):
    """kulcsszavak mindig string-lista legyen — a dashboard JS enélkül elszáll."""
    if isinstance(v, str):
        v = [t.strip() for t in v.split(",")]
    if not isinstance(v, list):
        return []
    return [str(t).strip() for t in v if str(t).strip()]

# ── Git ───────────────────────────────────────────────────────────────────────

def _pat():
    if not PAT_FILE.exists():
        raise RuntimeError(f"Nincs PAT: {PAT_FILE} — hozd létre a tokennel.")
    pat = PAT_FILE.read_text(encoding="utf-8").strip()
    if not pat:
        raise RuntimeError(f"Üres PAT-fájl: {PAT_FILE}")
    return pat


def _redact(s):
    """PAT kimaszkolása minden kimenetből/hibaüzenetből."""
    try:
        pat = PAT_FILE.read_text(encoding="utf-8").strip()
        if pat:
            s = s.replace(pat, "***PAT***")
    except Exception:
        pass
    return re.sub(r"github_pat_[A-Za-z0-9_]+", "***PAT***", s)


def _run(cmd, cwd=None):
    r = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError(_redact(f"$ {' '.join(cmd)}\n{r.stdout}{r.stderr}"))
    return r.stdout


def clone_repo():
    """Friss klón egyedi temp mappába (fix /tmp-útvonal TILOS — más session
    'nobody' tulajdonú fájljaival ütközne)."""
    repo = Path(tempfile.mkdtemp(prefix="sajto_repo_"))
    _run(["git", "clone", "--depth", "1",
          f"https://t0thgellert:{_pat()}@{REPO_URL}", str(repo)])
    _run(["git", "config", "user.name", "sajtofigyelo-bot"], cwd=repo)
    _run(["git", "config", "user.email", "toth.gellert@prohuman.hu"], cwd=repo)
    if not (repo / "docs" / "index.html").exists():
        raise RuntimeError("A klónban nincs docs/index.html — repo-struktúra változott?")
    return repo


def _push_with_retry(repo):
    """Push; ütközésnél (közben más commitolt) egy rebase-újrapróba."""
    try:
        _run(["git", "push"], cwd=repo)
    except RuntimeError:
        _run(["git", "pull", "--rebase"], cwd=repo)
        _run(["git", "push"], cwd=repo)

# ── DB merge + HTML ───────────────────────────────────────────────────────────

def read_db(html):
    """A 'const DB = {...}' blokk kiolvasása valódi JSON-parserrel.
    (Karakterszámolás helyett raw_decode — a szövegmezőkben lévő { } nem zavarja.)"""
    marker = "const DB = "
    start = html.find(marker + "{")
    if start == -1:
        raise ValueError("Nem találom a 'const DB = {' blokkot az index.html-ben.")
    pos = start + len(marker)
    try:
        db, consumed = json.JSONDecoder().raw_decode(html[pos:])
    except json.JSONDecodeError as e:
        raise ValueError(f"A DB blokk nem érvényes JSON: {e}") from e
    return db, start, pos + consumed


def _keys(c):
    """Egy cikk összes azonosító kulcsa: URL és cím+dátum is."""
    ks = {f"{_clean(c.get('cim','')).lower()}|{c.get('datum','')}"}
    if c.get("url"):
        ks.add(str(c["url"]).strip().rstrip("/"))
    return ks


def _normalize(c, warnings):
    """Egy bejövő cikk tisztítása. None = kihagyandó (az ok a warnings-ba kerül)."""
    c = dict(c)
    c["cim"] = _clean(c.get("cim"))
    if not c["cim"]:
        warnings.append("FIGYELEM: cím nélküli cikk kihagyva.")
        return None
    c["osszefoglalo"] = _clean(c.get("osszefoglalo"))
    c["url"] = _clean_url(c.get("url"))
    if not c.get("forras") or c["forras"] in ("Ismeretlen", "Médium"):
        c["forras"] = _source(c["url"]) or "Ismeretlen"

    fokat = str(c.get("fokategoria", "")).strip().lower()
    if fokat not in VALID_FOKAT:
        warnings.append(f"FIGYELEM: érvénytelen fokategoria ({c.get('fokategoria')!r}) "
                        f"— kihagyva: {c['cim'][:60]}")
        return None
    c["fokategoria"] = fokat

    datum = _valid_date(c.get("datum"))
    if not datum:
        warnings.append(f"FIGYELEM: hiányzó/érvénytelen datum ({c.get('datum')!r}) "
                        f"— kihagyva: {c['cim'][:60]}")
        return None
    c["datum"] = datum
    c["het"] = _monday(datum)          # kötelező — enélkül a hét-szűrő JS elszáll

    c["alkategoria"] = _clean(c.get("alkategoria")) or "Egyéb"
    c["kulcsszavak"] = _tags(c.get("kulcsszavak"))
    return c


def merge(db, new_articles):
    db.setdefault("cikkek", [])
    warnings = []
    existing = set()
    max_id = 0
    for c in db["cikkek"]:
        existing |= _keys(c)
        m = re.match(r"ART-(\d+)", str(c.get("id", "")))
        if m:
            max_id = max(max_id, int(m.group(1)))
    added = 0
    for raw in new_articles:
        c = _normalize(raw, warnings)
        if c is None:
            continue
        if _keys(c) & existing:
            continue
        max_id += 1
        c["id"] = f"ART-{max_id:04d}"
        db["cikkek"].append(c)
        existing |= _keys(c)
        added += 1
    db["total"] = len(db["cikkek"])
    db["generated"] = datetime.date.today().isoformat()
    return db, added, warnings


def write_html(repo, db):
    out = repo / "docs" / "index.html"
    html = out.read_text(encoding="utf-8")
    _, start, end = read_db(html)
    db_json = json.dumps(db, ensure_ascii=False, separators=(",", ":"))
    new_html = html[:start] + f"const DB = {db_json}" + html[end:]
    # Önellenőrzés: a beírt blokk visszaolvasható-e (különben üres oldal lenne)
    check, _, _ = read_db(new_html)
    if len(check.get("cikkek", [])) != len(db["cikkek"]):
        raise RuntimeError("Önellenőrzés hiba: a visszaolvasott DB cikkszáma eltér.")
    out.write_text(new_html, encoding="utf-8")
    # A hét-legördülőt a kliensoldali JS építi a DB-ből — itt nem nyúlunk hozzá.

# ── State ─────────────────────────────────────────────────────────────────────

def _load_state():
    if STATE.exists():
        try:
            state = json.loads(STATE.read_text(encoding="utf-8"))
            if isinstance(state.get("processed"), list):
                return state
        except json.JSONDecodeError:
            pass
        backup = STATE.with_suffix(".corrupt.json")
        shutil.copy(STATE, backup)
        print(f"FIGYELEM: sérült state fájl — mentve ide: {backup.name}, "
              "üres state-tel indulok (a duplikátum-szűrő véd az újraimport ellen).")
    return {"processed": []}


def _save_state(state):
    STATE.parent.mkdir(parents=True, exist_ok=True)
    STATE.write_text(json.dumps(state, ensure_ascii=False, indent=2),
                     encoding="utf-8")

# ── Alparancsok ───────────────────────────────────────────────────────────────

def cmd_scan():
    INCOMING.mkdir(parents=True, exist_ok=True)
    done = set(_load_state()["processed"])
    files = [p for p in INCOMING.iterdir()
             if p.is_file() and not p.name.startswith("~")]
    todo = sorted(p.name for p in files
                  if p.suffix.lower() == ".docx" and p.name not in done)
    old_doc = sorted(p.name for p in files if p.suffix.lower() == ".doc")
    if not todo:
        print("NINCS_UJ")
    for name in todo:
        print(name)
    for name in old_doc:
        print(f"FIGYELEM: {name} régi .doc formátum — a runner nem tudja "
              "olvasni, mentsd el .docx-ként.")


def cmd_extract(docx_name):
    path = INCOMING / docx_name
    if not path.exists():
        path = Path(docx_name)  # abszolút út is elfogadott
    if not path.exists():
        raise RuntimeError(f"Nincs ilyen fájl: {docx_name} (incoming/-ban sem).")
    text = extract_docx(str(path))
    WORK.mkdir(parents=True, exist_ok=True)
    # docx-enként egyedi név: friss fájl = nincs elavult mount-cache
    out = WORK / f"sajto_input_{Path(docx_name).stem}.txt"
    out.write_text(
        f"Mai dátum: {datetime.date.today()}\n\n{text}", encoding="utf-8")
    if len(text) < 40:
        print(f"FIGYELEM: nagyon rövid kinyert szöveg ({len(text)} karakter) — "
              "lehet, hogy a docx nem a várt formátumú.")
    print(f"OK {out} ({len(text)} karakter)")


def _read_json_retry(json_path, tries=4, wait=3):
    """A mount-szinkron késhet (csonka fájl) — pár újrapróba parse-hibánál."""
    import time
    last = None
    for i in range(tries):
        try:
            return json.loads(Path(json_path).read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError) as e:
            last = e
            if i < tries - 1:
                time.sleep(wait)
    raise RuntimeError(f"A(z) {json_path} nem olvasható érvényes JSON-ként "
                       f"(mount-szinkron?): {last}")


def cmd_apply(json_path, docx_name):
    new = _read_json_retry(json_path)
    articles = new["cikkek"] if isinstance(new, dict) else new
    if not isinstance(articles, list):
        raise RuntimeError("A JSON-ban nincs cikk-lista ('cikkek' kulcs vagy tömb).")

    repo = clone_repo()
    try:
        html = (repo / "docs" / "index.html").read_text(encoding="utf-8")
        db, _, _ = read_db(html)
        before = len(db.get("cikkek", []))
        db, added, warnings = merge(db, articles)
        for w in warnings:
            print(w)
        print(f"DB: {before} cikk + {added} új = {len(db['cikkek'])}")

        if added == 0:
            print("Nincs új cikk — nincs commit.")
        else:
            write_html(repo, db)
            _run(["git", "add", "docs/index.html"], cwd=repo)
            _run(["git", "commit", "-m",
                  f"Sajtófigyelő frissítés: {docx_name} (+{added} cikk)"], cwd=repo)
            _push_with_retry(repo)
            print("PUSH OK")
    finally:
        shutil.rmtree(repo, ignore_errors=True)

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
        print(f"HIBA: {_redact(str(e))}", file=sys.stderr)
        sys.exit(1)
