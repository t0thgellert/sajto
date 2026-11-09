#!/usr/bin/env python3
"""
Sajtófigyelő – Cowork pipeline runner
======================================
A Cowork scheduled task ezt a scriptet használja. Alparancsok:

  scan [mappa]              – feldolgozatlan .docx fájlok listája (alapból incoming/,
                              de bármelyik mappa megadható — pl. csatolt OneDrive-mappa)
  extract <docx>            – Word → szöveg + hyperlinkek → work/sajto_input_<név>.txt
  apply <json> <docx_név>   – új cikkek merge-elése a repo DB-jébe, HTML frissítés,
                              git commit + push, fájl megjelölése feldolgozottként
  relink-list [--limit=N]   – az URL nélküli cikkek listája JSON-ban (link-pótlás)
  relink-apply <json>       – talált linkek visszaírása id alapján + commit/push
                              (--dry-run: validálás commit nélkül;
                               --fix-sources: forrásnév-egységesítés is)
  fix-sources [--dry-run]   – a forrásnevek egységesítése az egész DB-ben
  audit-links [--dry-run]   – a MÁR bent lévő linkek auditja: forrásnév kontra
                              URL-domain (slug-bizonyítékkal), URL-dátum kontra
                              `datum`. Csak a `forras`-t írja át, URL-t soha.
  sync-cowork [--dry-run]   – a lokális cowork/*.py, *.md és a watchdog.yml
                              feltöltése a repóba (az `apply` csak a
                              docs/index.html-t commitolja, ezért a repóbeli
                              instrukció külön szinkron nélkül elavul)
  kezi-javitas <json> [--dry-run]
                            – kézi link-javítás id alapján: URL beírása,
                              cseréje vagy törlése (meglévő URL-t is felülír)
  ws-backfill [--dry-run]   – a WS-kivonatok (kivonat mező) visszatöltése
                              a work/sajto_input_*.txt fájlokból
  docx-relink <mappa> [--dry-run]
                            – a mappa összes körképéből pótolja a link nélküli
                              cikkek linkjét, ha a WS megadta (idempotens)

LINK-PÓTLÁS (relink-*)
----------------------
Az `apply` csak HOZZÁFŰZ, meglévő cikket nem módosít — ezért kellett külön út.
Az `url_forras` mező mondja meg, honnan jött a link: "docx" (Word-hyperlink),
"talalt" (kereséssel pótolt), "nincs" (ellenőrizve nincs online változat).
A `relink-apply` minden beírást validál; a részleteket lásd _validate_link().
MÉRT ADAT (2026-08-05): 449 cikkből 62 volt URL nélküli.

2026-10-14 változások:
- Az extract a Word-MEZŐKÓDOS linkeket (HYPERLINK "…") is kiolvassa. A WS
  dokumentumai a linkek egy részét így tárolják; a 07-31-i körképben 14-ből 2
  link veszett el emiatt, és csak 9 nap múlva, kereséssel került elő.
- relink-list: „cikkek" = keresendők (a legkevesebbszer keresettek elöl),
  „kezi_lista" = amit a kereső nem ér el (NEM_KERESHETO domain) vagy
  MAX_KERESES sikertelen kör után sem lett meg — kézi pótlásra.
- relink-apply: {"id", "nincs_talalat": true} = sikertelen keresés naplózása
  (url_keresve). Ütközés (az URL más cikknél van) és elavult javaslat (a
  cikknek már van URL-je) NEM növeli az url_probak-ot. „nincs" lezárás csak
  'indok' mezővel. "tipus": "kezi" → url_forras "kezi" (ember találta).
- Új url_forras érték: "kezi". Új mezők: url_keresve, url_keresve_utoljara,
  kivonat (a körkép tételenkénti összefoglalója, szó szerint).
- A publikus oldalra (DB, felület, commit-üzenet) NEM kerül utalás a körkép
  forrására, és belső jegyzet sem: az 'indok' csak a work/*.json-ban él
  (2026-10-14, Gellért kérése). A mezőnév ezért 'kivonat', nem 'ws_kivonat'.

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
FONTOS (2026-08-05f): ha BÁRMELY cikk kiesik a validáción, az apply a jó
cikkeket kimenti és pusholja, de a docx-et NEM jelöli feldolgozottnak, és
hibakóddal áll le. Korábban ilyenkor is bekerült a processed.json-ba, tehát egy
hibás JSON véglegesen elnyelte az egész körképet. Az apply újrafuttatása
biztonságos: a merge cím+dátum / URL alapján dedupál.
A kulcsszavakat a cowork/vocab.py ZÁRT listájára képezi; ami nincs a listán
(egyszeri cégnév, város, szóvivő neve), az eldobásra kerül FIGYELEM sorral —
a cikk nem marad ki miatta. A vocab.py fejében írja, hogyan bővítsd.
"""

import sys
import re
import json
import difflib
import datetime
import subprocess
import shutil
import tempfile
import unicodedata
from pathlib import Path
from urllib.parse import urlparse

# A zárt címkelista a runner mellett lakik (cowork/vocab.py). Nélküle nincs
# feldolgozás: címke-validáció nélkül visszatérne az 588-címkés zaj.
sys.path.insert(0, str(Path(__file__).resolve().parent))
from vocab import VOCAB, normalize_tags  # noqa: E402

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
    # 2026-08-05: a link-pótlás során felmerült domainek
    "gyartastrend.hu": "Gyártástrend", "ado.hu": "Adó Online",
    "hipa.hu": "Hipa", "ictglobal.hu": "ICT Global",
    "pakspress.hu": "Paks Press", "beol.hu": "Békés Vármegyei Hírportál",
    "uniside.hu": "uniside.hu", "rtl.hu": "RTL", "atv.hu": "ATV",
    "hirado.hu": "Híradó.hu", "pharmaonline.hu": "PharmaOnline",
    "fuhu.hu": "Független Hírügynökség", "kontroll.hu": "Kontroll",
    "profession.hu": "Profession.hu", "hrpwr.hu": "HRPWR",
    # 2026-10-14: a link nélküli tételek felméréséből
    "nlc.hu": "NLC", "24.hu": "24.hu",
    # 2026-11-09: Gellért kérésére (ART-0658, Klubrádió domain_tipp hiányzott)
    "klubradio.hu": "Klubrádió",
}

# Több domainen élő médiumok. A G7 2026-ban a Telex aloldala
# (telex.hu/g7/<kat>/2026/07/14/<slug>), de régi cikkei g7.hu-n vannak —
# a domain-guard mindkettőt fogadja el. Kulcs: a forrásnév _fold()-olt alakja.
MEDIA_ALIASES: dict[str, set[str]] = {
    "g7": {"g7.hu", "telex.hu"},
    "g7telex": {"g7.hu", "telex.hu"},
    "telex": {"telex.hu", "g7.hu"},
    # A WS „HRPOWER"-ként írja, a domain hrpwr.hu (2026-10-14).
    "hrpower": {"hrpwr.hu"},
    "hrpwr": {"hrpwr.hu"},
}

# ── Word kinyerés ─────────────────────────────────────────────────────────────

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_HL   = f"{{{_W_NS}}}hyperlink"
_RID  = f"{{{_R_NS}}}id"
_TAG_P   = f"{{{_W_NS}}}p"
_TAG_TBL = f"{{{_W_NS}}}tbl"


def _unwrap_if_base64(path: str) -> str:
    """A Power Automate flow base64 SZÖVEGET írhat a docx helyett.

    Ha a 'Create file' akció File Content mezője a contentBytes-ot dekódolás
    nélkül adja át, a lementett fájl a dokumentum base64 szövege lesz: pontosan
    4/3-szor nagyobb, és NEM érvényes ZIP (a docx ZIP-alapú). Mérve 2026-08-04:
    a 155 669 bájtos valódi docx-ből 207 560 bájtos base64-fájl lett.

    Felismerés: a docx 'PK\\x03\\x04'-gyel kezdődik, ennek base64-e 'UEsDB'.
    Ha ezt látjuk, dekódoljuk egy temp fájlba, és azt adjuk vissza. Így a
    pipeline akkor is működik, ha a flow-oldali javítás elmarad.
    """
    import base64
    with open(path, "rb") as f:
        head = f.read(5)
    if head[:2] == b"PK":
        return path                      # rendes docx, nincs teendő
    if head != b"UEsDB":
        return path                      # nem base64-be csomagolt docx — a
                                         # python-docx adja majd a hibát
    with open(path, "rb") as f:
        data = f.read()
    try:
        raw = base64.b64decode(b"".join(data.split()), validate=True)
    except Exception as e:
        raise RuntimeError(
            f"A fájl base64-nek tűnik ('UEsDB' kezdet), de nem dekódolható: {e}")
    if raw[:2] != b"PK":
        raise RuntimeError("A base64-dekódolás nem docx-et adott.")
    tmp = Path(tempfile.mkdtemp(prefix="sajto_b64_")) / Path(path).name
    tmp.write_bytes(raw)
    print(f"FIGYELEM: {Path(path).name} base64 szövegként volt mentve "
          f"({len(data)} bájt) — dekódolva {len(raw)} bájt. A Power Automate "
          f"'Create file' File Content mezője base64ToBinary() nélkül fut.")
    return str(tmp)


_INSTR     = f"{{{_W_NS}}}instrText"
_FLDSIMPLE = f"{{{_W_NS}}}fldSimple"
_FLD_URL   = re.compile(r'HYPERLINK\s+"([^"]+)"')
_WS_MINTASOR = re.compile(r"^Cím\s*\(Médium\s*,")


def _field_hyperlink(p_elem):
    """Mezőkódos hyperlink (HYPERLINK "url") egy bekezdésben, vagy None.

    A WS-dokumentumok a linkek egy részét nem w:hyperlink elemként, hanem
    Word-mezőkódként tárolják (w:instrText, ritkábban w:fldSimple). Mérve
    2026-10-14: a 07-31-i körképben 14 cikkből 2 linkje volt így, és a régi
    kinyerés mindkettőt elvesztette — a link-pótlás 9 nap múlva kereséssel
    találta meg ugyanazt a két URL-t. A tartalomjegyzék belső hivatkozása
    (HYPERLINK \\l "_Toc…") nem URL, azt kihagyjuk.
    """
    instr = "".join(t.text or "" for t in p_elem.iter(_INSTR))
    instr += " " + " ".join(f.get(f"{{{_W_NS}}}instr") or ""
                            for f in p_elem.iter(_FLDSIMPLE))
    for m in _FLD_URL.finditer(instr):
        u = m.group(1).strip()
        if u.lower().startswith(("http://", "https://")):
            return u
    return None


def extract_docx(path: str) -> str:
    """Bekezdések ÉS táblázatok szövege dokumentum-sorrendben, hyperlinkekkel."""
    from docx import Document
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    doc = Document(_unwrap_if_base64(path))
    rels = {rid: rel._target for rid, rel in doc.part.rels.items()
            if "hyperlink" in rel.reltype}

    def para_line(p_elem):
        para = Paragraph(p_elem, doc)
        text = para.text.strip()
        if not text:
            return None
        if _WS_MINTASOR.match(text):
            # A WS-sablon mintasora, nem cikk. A linkje (blikk.hu) a sablonból
            # öröklődik — 2026-08-05-én hat cikkhez is „hozzáragadt".
            return None
        url = next((rels[hl.get(_RID)] for hl in p_elem.iter(_HL)
                    if hl.get(_RID) in rels), None)
        if not url:
            url = _field_hyperlink(p_elem)
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


def _tags(v, warnings=None, cim=""):
    """kulcsszavak -> a zárt vokabuláriumra (vocab.py) képezett string-lista.

    Miért kötelező a szűkítés: amíg a címkék szabad szövegként jöttek, 435
    cikkre 588 különböző címke keletkezett, ebből 389 pontosan EGY cikkhez —
    ez nem szűrő, csak zaj. A listán kívüli címke ELDOBÁSRA kerül (a hosszú
    farkat a szabadszavas kereső fedi), és FIGYELEM sorban megjelenik, hogy
    lásd, ha egy visszatérő új szereplőt fel kell venni a vocab.py-ba.
    """
    if isinstance(v, str):
        v = [t.strip() for t in v.split(",")]
    if not isinstance(v, list):
        v = []
    raw = [str(t).strip() for t in v if str(t).strip()]
    kept = normalize_tags(raw)
    if warnings is not None:
        dropped = [t for t in raw if not normalize_tags([t])]
        if dropped:
            warnings.append(
                f"FIGYELEM: vokabuláriumon kívüli címke eldobva "
                f"({', '.join(dropped)}) — {cim[:50]}")
        if not kept:
            warnings.append(
                f"FIGYELEM: a cikk egyetlen érvényes címke nélkül maradt "
                f"— {cim[:50]}")
    return kept


# ── Link-pótlás / forrásnév-egységesítés segédfüggvények ─────────────────────

def _fold(s):
    """Ékezet- és írásjel-független összehasonlító kulcs."""
    s = unicodedata.normalize("NFKD", str(s).lower())
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]", "", s)


def _src_key(s):
    """Forrásnév-kulcs: a 'economx.hu' és az 'Economx' ugyanide esik."""
    return _fold(s).removesuffix("hu").removesuffix("com")


_SRC2DOM = {_fold(v): d for d, v in DOMAIN_MAP.items()}


def _domain(url):
    try:
        return urlparse(url).netloc.lower().removeprefix("www.")
    except Exception:
        return ""


def _domain_ok(url, forras):
    """(status, elvart) — status: 'ok' | 'eltero' | 'ismeretlen'.

    'ismeretlen' NEM elutasítás: a DOMAIN_MAP hiányos, és egy jó linket nem
    dobunk el amiatt, hogy a térképünk nem ismeri a médiumot. Ilyenkor viszont
    slug- vagy oldal_cim-bizonyíték kell, és a riport kéri a felvételt.
    """
    dom = _domain(url)
    if not dom:
        return "eltero", None
    # ELŐBB a teljes forrásnév ("Piac és Profit" NEM vágható szét!), és csak
    # utána a tagok ("G7 / Telex", "G7–Telex" — ott bármelyik tag elég).
    parts = [str(forras)] + [x for x in re.split(r"[/|;–—+&]", str(forras))
                             if x.strip() and x.strip() != str(forras).strip()]
    known = []
    for x in parts:
        for cand in MEDIA_ALIASES.get(_fold(x), set()):
            if dom == cand or dom.endswith("." + cand):
                return "ok", cand
            known.append(cand)
        exp = _SRC2DOM.get(_fold(x))
        if exp:
            known.append(exp)
            if dom == exp or dom.endswith("." + exp):
                return "ok", exp
        # a forrásnév maga is domain ("vg.hu"), vagy a domain törzse ("GyártásTrend")
        if _fold(x) and _fold(x) in (_fold(dom), _fold(dom.rsplit(".", 1)[0])):
            return "ok", dom
    return ("eltero", known[0]) if known else ("ismeretlen", None)


_URL_DAY = (
    r"/(20\d{2})/(\d{2})/(\d{2})/",              # telex.hu/gazdasag/2026/05/15/...
    r"/(20\d{2})(\d{2})(\d{2})/",                # portfolio.hu/gazdasag/20260722/...
    r"[-_/](20\d{2})(\d{2})(\d{2})(?:\D|$)",     # hrportal.hu/...-20260727.html
)
_URL_MONTH = r"/(20\d{2})/(\d{2})/(?!\d{2}/)"    # vg.hu/cegvilag/2026/04/<slug>


def _url_date(url):
    """A URL-be sütött napra pontos dátum, vagy None."""
    for pat in _URL_DAY:
        m = re.search(pat, str(url))
        if m:
            try:
                return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                continue
    return None


def _url_month(url):
    """(év, hónap) a URL-ből, ha csak év/hónap van benne. Gyenge jel."""
    m = re.search(_URL_MONTH, str(url))
    return (int(m.group(1)), int(m.group(2))) if m else None


def _slug(url):
    """A URL utolsó beszédes szegmense (a záró cikk-ID és a .html nélkül)."""
    try:
        segs = [s for s in urlparse(url).path.split("/") if s]
    except Exception:
        return ""
    while segs and re.fullmatch(r"\d+", segs[-1]):
        segs.pop()                                   # autopro.hu/.../<slug>/1609009
    if not segs:
        return ""
    s = re.sub(r"\.(html?|php|aspx)$", "", segs[-1])
    s = re.sub(r"[-_]\d{4,}$", "", s)                # profitline.hu/<slug>-487869
    s = re.sub(r"^\d{4,}[-_]", "", s)                # nepszava.hu/3322570_<slug>
    return s


def _ratio(a, b):
    return difflib.SequenceMatcher(None, a, b).ratio()


def _slug_ok(url, cim):
    """A URL slugja a cím slugosított alakja-e.

    Ez OFFLINE ellenőrizhető bizonyíték, és több médiumnál az egyetlen elérhető:
    az autopro.hu 520-at ad, az ado.hu robots.txt-je tiltja a lekérést, tehát
    oldal_cim-et nem lehet róluk beszerezni. A slug viszont a médium saját CMS-e
    által a címből generált string — nehezebb véletlenül egyezésbe hozni, mint
    egy keresőtalálat sorrendjét.
    """
    a, b = _fold(_slug(url)), _fold(cim)
    if len(a) < 15 or len(b) < 15:
        return False
    return a in b or b in a or _ratio(a, b) >= 0.80


# Ahol a link nem cikk, hanem videó vagy PDF-szkennelés: ott a slug SOHA nem
# egyezhet a címmel, és a forrásnév (Kontroll, ATV, Népszava-print) helyesen tér
# el a domaintől. Ezeket ki kell hagyni, különben minden futásban hamis riasztás.
NEM_CIKK_DOMAIN = {
    "youtube.com", "youtu.be", "e.pcloud.link", "facebook.com", "fb.watch",
    "drive.google.com", "1drv.ms", "open.spotify.com", "vimeo.com",
}


def _words(s):
    """Ékezetmentes, 5+ karakteres szavak — a slug és a cím összevetéséhez."""
    s = unicodedata.normalize("NFKD", str(s).lower())
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return [w for w in re.split(r"[^a-z0-9]+", s) if len(w) >= 5]


def _slug_tokens_ok(url, cim):
    """Szó-szintű egyezés a slug és a cím között, MINDKÉT irányban.

    Miért kell a `_slug_ok` MELLÉ: az teljes stringet hasonlít, és elhasal, ha
    a CMS extra kulcsszavakat fűz a slugba
    (`schiffer-andras-munkaugyek-...-hatalomvaltas-kizsakmanyolas-politika`),
    vagy ha a cím ragozott alakot használ (`szabalyait` vs a slug `szabalyai`).
    Szavak részhalmazaként mindkettő átmegy, a magyar ragozást pedig a
    substring-keresés nyeli le. A két irány maximumát vesszük, mert hol a cím,
    hol a slug a bővebb.

    A `|` elválasztó azért kell, hogy egy szó ne tudjon két szomszédos szó
    határán átnyúlva „egyezni".
    """
    sw, cw = _words(_slug(url)), _words(cim)
    if len(sw) < 3 or len(cw) < 3:
        return False

    def share(src, dst):
        blob = "|".join(dst)
        return sum(1 for w in src if w in blob) / len(src)

    return max(share(cw, sw), share(sw, cw)) >= 0.7


def _link_bizonyitott(url, cim):
    """A link bizonyítottan EHHEZ a cikkhez tartozik-e (offline, slug alapján)."""
    return _slug_ok(url, cim) or _slug_tokens_ok(url, cim)


def _title_ok(oldal_cim, cim):
    """A lekért oldal <title>-je egyezik-e a cikk címével."""
    a, b = _fold(oldal_cim), _fold(cim)
    if not a or not b:
        return False
    return a in b or b in a or _ratio(a, b) >= 0.82


def _existing_urls(db):
    return {str(c["url"]).strip().rstrip("/")
            for c in db.get("cikkek", []) if c.get("url")}


def _stamp_docx_source(db):
    """Visszamenőleg: minden URL-lel bíró, jelöletlen cikk linkje docx-ből jött."""
    n = 0
    for c in db.get("cikkek", []):
        if c.get("url") and not c.get("url_forras"):
            c["url_forras"] = "docx"
            n += 1
    return n


def _dup_ok(art, other):
    """Törölhető-e `art` mint `other` duplikátuma? Mind a négy feltétel kell."""
    if other is None or not other.get("url"):
        return False, "a megadott párnak nincs URL-je (vagy nem létezik)"
    if art.get("datum") != other.get("datum"):
        return False, f"eltérő dátum ({art.get('datum')} vs {other.get('datum')})"
    if _src_key(art.get("forras", "")) != _src_key(other.get("forras", "")):
        return False, f"eltérő forrás ({art.get('forras')} vs {other.get('forras')})"
    a, b = _fold(art.get("cim", "")), _fold(other.get("cim", ""))
    if not (a in b or b in a or _ratio(a, b) >= 0.85):
        return False, "a két cím nem elég hasonló"
    return True, "azonos dátum, forrás és cím"


def _validate_link(entry, art, db_urls):
    """(ok, url_forras, indok, figyelmeztetes) egyetlen javasolt linkre.

    Elfogadás: legalább EGY erős bizonyíték (napra pontos URL-dátum a cikk
    dátumától ±3 napon belül / slug-egyezés / lekért oldal_cim egyezése).
    Dátum-ELLENTMONDÁS (a URL dátuma vagy éve-hónapja más) mindig elutasítás,
    akkor is, ha egyébként van erős bizonyíték.
    """
    if str(entry.get("url_forras", "")).strip() == "nincs" and not entry.get("url"):
        return True, "nincs", "lezárva: nincs online változat", None

    url = _clean_url(entry.get("url"))
    if not url:
        return False, None, "nem http(s) URL", None
    if re.search(r"\s", str(entry.get("url", ""))):
        return False, None, "a URL szóközt tartalmaz (elírás?)", None
    if art.get("url"):
        return False, None, f"a cikknek MÁR van URL-je ({art['url'][:50]}) — nem írom át", None
    if url.strip().rstrip("/") in db_urls:
        return False, None, "ez a URL már szerepel a DB-ben (más cikknél)", None

    dstat, expected = _domain_ok(url, art.get("forras", ""))
    if dstat == "eltero":
        return False, None, (f"domain eltérés: {_domain(url)} ≠ {expected or '?'} "
                             f"(forrás: {art.get('forras')})"), None
    warn = None
    if dstat == "ismeretlen":
        warn = (f"FIGYELEM: {_domain(url)} nincs a DOMAIN_MAP-ban "
                f"(forrás: {art.get('forras')}) — vedd fel a runner.py DOMAIN_MAP-jába.")

    try:
        adate = datetime.date.fromisoformat(art["datum"])
    except Exception:
        return False, None, "a cikk datum mezője érvénytelen", warn

    # dátum-ellentmondás vizsgálata
    ud = _url_date(url)
    if ud:
        delta = abs((ud - adate).days)
        if delta > 3:
            return False, None, (f"URL-dátum {ud} vs cikk dátum {adate} "
                                 f"({delta} nap, max. 3)"), warn
    else:
        um = _url_month(url)
        if um and um != (adate.year, adate.month):
            return False, None, (f"URL éve-hónapja {um[0]}-{um[1]:02d} vs cikk "
                                 f"dátuma {adate}"), warn

    bizonyitek = []
    if ud:
        bizonyitek.append(f"URL-dátum {ud}")
    if _slug_ok(url, art.get("cim", "")):
        bizonyitek.append("slug-egyezés")
    oc = entry.get("oldal_cim")
    if oc and _title_ok(oc, art.get("cim", "")):
        bizonyitek.append("lekért oldal címe egyezik")
    elif oc:
        return False, None, (f"oldal_cim nem egyezik a cikk címével "
                             f"(oldal: {str(oc)[:60]!r})"), warn

    if not bizonyitek:
        return False, None, ("nincs erős bizonyíték: a URL-ben nincs napra pontos "
                             "dátum, a slug nem egyezik a címmel, és nincs "
                             "'oldal_cim' sem"), warn
    if dstat == "ismeretlen" and bizonyitek == [f"URL-dátum {ud}"]:
        return False, None, ("ismeretlen domain: a puszta URL-dátum nem elég, "
                             "slug- vagy oldal_cim-egyezés kell"), warn
    return True, "talalt", " + ".join(bizonyitek), warn


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


def clone_repo(read_only=False):
    """Friss klón egyedi temp mappába (fix /tmp-útvonal TILOS — más session
    'nobody' tulajdonú fájljaival ütközne)."""
    repo = Path(tempfile.mkdtemp(prefix="sajto_repo_"))
    url = (f"https://{REPO_URL}" if read_only
           else f"https://t0thgellert:{_pat()}@{REPO_URL}")
    _run(["git", "clone", "--depth", "1", url, str(repo)])
    if not read_only:
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
    c["kulcsszavak"] = _tags(c.get("kulcsszavak"), warnings, c["cim"])
    # A forrásnév és az URL domainje nem mondhat ellent egymásnak. A docx
    # hyperlink-rId újrahasznosítása pontosan ilyen hibát termel: 2026-08-05-én
    # a 449 cikkből 17-nél tért el a kettő, és ötnél a slug bizonyította, hogy
    # az URL a jó és a forrásnév a rossz (pl. „Pénzcentrum" egy portfolio.hu-s
    # linken). URL-t itt SOHA nem dobunk el — az emberi döntés.
    if c["url"] and _domain(c["url"]) not in NEM_CIKK_DOMAIN:
        _st, _ = _domain_ok(c["url"], c["forras"])
        if _st == "eltero":
            _dnev = _source(c["url"])
            if (_dnev and _src_key(_dnev) != _src_key(c["forras"])
                    and _link_bizonyitott(c["url"], c["cim"])):
                # A slug a médium CMS-e által a CÍMBŐL generált string: ha
                # egyezik, a link bizonyítottan ehhez a cikkhez tartozik,
                # tehát a domain az igazság, nem a docx-ben olvasott név.
                warnings.append(
                    f"FIGYELEM: forras átírva {c['forras']!r} -> {_dnev!r} "
                    f"(URL-domain {_domain(c['url'])}, a slug egyezik a címmel) "
                    f"— {c['cim'][:50]}")
                c["forras"] = _dnev
            else:
                warnings.append(
                    f"FIGYELEM: gyanús link — a forras ({c['forras']!r}) nem "
                    f"illik az URL domainjéhez ({_domain(c['url'])}), és a slug "
                    f"sem egyezik a címmel. ELLENŐRIZD kézzel: {c['cim'][:50]}")
    # A docx hyperlinkjéből jött link a megbízható eset — jelöljük, hogy a
    # relink-list ne akarja „pótolni", és a keresett linkek auditálhatók legyenek.
    if c["url"]:
        c["url_forras"] = "docx"
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
    added = elutasitva = duplikatum = 0
    for raw in new_articles:
        c = _normalize(raw, warnings)
        if c is None:
            elutasitva += 1          # validáción kiesett — NEM feldolgozott
            continue
        if _keys(c) & existing:
            duplikatum += 1          # már bent van — ez a jóindulatú eset
            continue
        max_id += 1
        c["id"] = f"ART-{max_id:04d}"
        db["cikkek"].append(c)
        existing |= _keys(c)
        added += 1
    db["total"] = len(db["cikkek"])
    db["generated"] = datetime.date.today().isoformat()
    # A hármas bontás azért kell, mert az `added == 0` KÉT teljesen különböző
    # dolgot jelenthet: (a) duplikált küldés — rendben van, (b) az egész fájl
    # kiesett a validáción — adatvesztés. A cmd_apply ezen dönt.
    stats = {"bemenet": len(new_articles), "elutasitva": elutasitva,
             "duplikatum": duplikatum, "uj": added}
    return db, added, warnings, stats


def write_html(repo, db):
    out = repo / "docs" / "index.html"
    html = out.read_text(encoding="utf-8")
    _, start, end = read_db(html)
    # A '<' escape-elése: egy „</script>" bármely szövegmezőben üres oldalt
    # csinálna (a WS-kivonatok óta több a szabad szöveg). JSON-ként azonos érték.
    db_json = json.dumps(db, ensure_ascii=False,
                         separators=(",", ":")).replace("<", "\\u003c")
    new_html = html[:start] + f"const DB = {db_json}" + html[end:]
    # Önellenőrzés: a beírt blokk visszaolvasható-e (különben üres oldal lenne)
    check, _, _ = read_db(new_html)
    if len(check.get("cikkek", [])) != len(db["cikkek"]):
        raise RuntimeError("Önellenőrzés hiba: a visszaolvasott DB cikkszáma eltér.")
    out.write_text(new_html, encoding="utf-8")
    # A hét-legördülőt a kliensoldali JS építi a DB-ből — itt nem nyúlunk hozzá.

# ── WS-kivonatok (a körkép saját összefoglalói) ───────────────────────────────

_WS_TITLE = re.compile(
    r"^(?P<cim>.+?)\s*\((?P<forras>[^()]*?),\s*"
    r"(?P<datum>\d{4}\.\s*\d{1,2}\.\s*\d{1,2}\.?)\s*\)\s*"
    r"(?:\[URL:\s*(?P<url>[^\]]+)\])?\s*$")


def _ws_kivonatok(text):
    """A work/sajto_input_*.txt tételei: cím, forrás, dátum, URL, WS-kivonat.

    A körképben minden tétel egy címsor („Cím (Médium, 2026.09.30.)",
    opcionális [URL: …]-lel), utána a WS saját, 3–5 mondatos kivonata egy
    (ritkán két) bekezdésben. Mérve 2026-10-14: 29 körkép, 508 tétel, 497 egy
    bekezdéses, 8 két bekezdéses. A 80 karakternél rövidebb sor alcím
    („Vendégmunka", „Nem volt releváns megjelenés.") — az lezárja a kivonatot.
    """
    lines = [l.strip() for l in str(text).split("\n") if l.strip()]
    out, i = [], 0
    while i < len(lines):
        m = _WS_TITLE.match(lines[i])
        if not m:
            i += 1
            continue
        body, j = [], i + 1
        while (j < len(lines) and not _WS_TITLE.match(lines[j])
               and len(lines[j]) >= 80):
            body.append(lines[j])
            j += 1
        d = re.findall(r"\d+", m.group("datum"))
        out.append({
            "cim": m.group("cim").strip(),
            "forras": m.group("forras").strip(),
            "datum": f"{int(d[0]):04d}-{int(d[1]):02d}-{int(d[2]):02d}",
            "url": (m.group("url") or "").strip() or None,
            "kivonat": "\n\n".join(body),
        })
        i = j
    return out


def _ws_match(art, tetelek):
    """A cikkhez tartozó WS-tétel: azonos dátum és (közel) azonos cím."""
    fc = _fold(art.get("cim", ""))
    jelolt = [t for t in tetelek
              if t["datum"] == art.get("datum") and t["kivonat"]]
    for t in jelolt:
        if _fold(t["cim"]) == fc:
            return t
    best, best_r = None, 0.0
    for t in jelolt:
        ft = _fold(t["cim"])
        r = 1.0 if (fc and ft and (fc in ft or ft in fc)) else _ratio(ft, fc)
        if r > best_r:
            best, best_r = t, r
    return best if best_r >= 0.85 else None


def _set_ws_kivonat(art, tetel):
    """A körkép saját kivonata szó szerint a kivonat mezőbe. True = változott."""
    kiv = (tetel or {}).get("kivonat") or ""
    if not kiv or art.get("kivonat") == kiv:
        return False
    art.pop("ws_kivonat", None)
    art["kivonat"] = kiv
    return True


def _ws_kivonat_ujakhoz(uj, docx_name):
    """Az apply-ban: a most bekerült cikkek WS-kivonata a docx kinyeréséből."""
    if not uj:
        return
    txt = WORK / f"sajto_input_{Path(docx_name).stem}.txt"
    if not txt.exists():
        print(f"FIGYELEM: nincs {txt.name} — a WS-kivonatok kimaradnak "
              f"(utólag pótolható: ws-backfill).")
        return
    tetelek = _ws_kivonatok(txt.read_text(encoding="utf-8"))
    n = sum(1 for c in uj if _set_ws_kivonat(c, _ws_match(c, tetelek)))
    print(f"WS-kivonat: {n} / {len(uj)} új cikkhez.")
    if n < len(uj):
        print("FIGYELEM: nem minden új cikknek lett WS-kivonata — a cím vagy a "
              "dátum eltér a docx-ben olvashatótól.")

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

def cmd_scan(folder=None):
    """Feldolgozatlan .docx fájlok listája. A mappa opcionálisan megadható —
    így a bemenet bárhol lehet (pl. egy másik csatolt mappa), nem csak az
    incoming/-ban. Megadott mappát NEM hoz létre: ha nincs, az hiba."""
    if folder:
        src = Path(folder)
        if not src.is_dir():
            raise RuntimeError(f"Nincs ilyen mappa: {folder}")
    else:
        src = INCOMING
        src.mkdir(parents=True, exist_ok=True)
    done = set(_load_state()["processed"])
    # rglob: alkönyvtárakat is látja. A dedupe és az extract fájlNÉV szerint
    # megy, nem útvonal szerint — két azonos nevű fájl külön alkönyvtárban
    # ÜTKÖZIK (lásd cmd_extract).
    files = [p for p in src.rglob("*")
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
        # alkönyvtárban is (junctionölt OneDrive-mappa)
        hits = [p for p in INCOMING.rglob(Path(docx_name).name) if p.is_file()]
        if len(hits) > 1:
            names = ", ".join(str(p.relative_to(INCOMING)) for p in hits)
            raise RuntimeError(
                f"Több '{docx_name}' nevű fájl van az incoming/ alatt ({names}) — "
                "a dedupe fájlnév szerint megy, töröld a duplikátumot.")
        if hits:
            path = hits[0]
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
    if not articles:
        raise RuntimeError(
            "A JSON 0 cikket tartalmaz. Ez feldolgozási hiba, nem üres körkép — "
            "a fájl NEM lett feldolgozottként megjelölve. Nézd meg a "
            "work/sajto_input_*.txt-t, és írd újra a JSON-t.")

    repo = clone_repo()
    try:
        html = (repo / "docs" / "index.html").read_text(encoding="utf-8")
        db, _, _ = read_db(html)
        before = len(db.get("cikkek", []))
        db, added, warnings, stats = merge(db, articles)
        _ws_kivonat_ujakhoz(db["cikkek"][before:], docx_name)
        for w in warnings:
            print(w)
        print(f"DB: {before} cikk + {added} új = {len(db['cikkek'])}  "
              f"[bemenet {stats['bemenet']}, már megvolt {stats['duplikatum']}, "
              f"validáción kiesett {stats['elutasitva']}]")

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

    # A fájl CSAK akkor kerül a processed.json-ba, ha egyetlen cikk sem esett ki
    # a validáción. Korábban az `added == 0` ág is ide futott, tehát egy hibás
    # JSON (rossz dátumformátum, elírt fokategoria) VÉGLEGESEN elnyelte a
    # körképet: commit nem lett, a fájl mégis feldolgozottnak számított, és a
    # `scan` soha többé nem hozta elő. Az újrafuttatás biztonságos: a merge
    # cím+dátum / URL alapján dedupál, tehát a már bekerült cikkek nem
    # duplikálódnak.
    if stats["elutasitva"]:
        raise RuntimeError(
            f"{stats['elutasitva']} cikk kiesett a validáción (lásd a FIGYELEM "
            f"sorokat feljebb). {added} cikk bekerült és ki van pusholva, de a(z) "
            f"'{docx_name}' NEM lett feldolgozottként megjelölve. Javítsd a "
            f"JSON-ban a jelzett cikkeket, és futtasd újra ugyanezt az apply-t.")

    state = _load_state()
    if docx_name not in state["processed"]:
        state["processed"].append(docx_name)
    state["last_run"] = datetime.datetime.now().isoformat(timespec="seconds")
    _save_state(state)
    print("KESZ")



MAX_PROBAK = 3          # ennyi ELUTASÍTOTT találat után a cikket nem listázzuk többé
MAX_KERESES = 3         # ennyi „kerestem, nincs találat" kör után kézi listára kerül
# Domainek, amelyeket a Cowork keresője nem ér el (mérve 2026-10-14: a
# domain-szűrt WebSearch proxy HTTP 400-at ad, az nlc.hu lekérése tiltott).
# Ezek tételeit a relink-list nem adja keresésre, hanem kézi listára teszi:
# ott a link csak a docx-ből vagy emberi kereséssel jöhet.
NEM_KERESHETO = {"nlc.hu", "24.hu", "rtl.hu"}
_LINK_ALLAPOT = ("url_probak", "url_keresve", "url_keresve_utoljara")
BROADCAST = {"rtl", "atv", "hirado", "m1", "kossuthradio", "inforadio",
             "trendfm", "echotv", "kontroll"}

# Forrásnév-kanonizálás. A _src_key() a mechanikus változatokat összehozza
# (HR Portál / Hrportál / HR Portal; Economx / economx.hu), ezért itt csak azt
# adjuk meg, MELYIK írásmód a hivatalos. Amit nem sorolunk fel, ott a DB-ben
# leggyakoribb írásmód nyer, és a riport kiírja.
CANON_SOURCE = {
    "hrportal": "HR Portál", "economx": "Economx", "mfor": "MFOR",
    "autopro": "Autopro", "gyartastrend": "Gyártástrend", "atv": "ATV",
    "kontroll": "Kontroll", "novekedes": "Növekedés.hu", "hirado": "Híradó.hu",
    "baon": "Baon", "mezohir": "Mezőhír", "trendfm": "TrendFM",
    "g7": "G7", "g7telex": "G7", "vg": "Világgazdaság", "telex": "Telex",
    "portfolio": "Portfolio", "penzcentrum": "Pénzcentrum",
    "profitline": "Profitline", "nepszava": "Népszava",
}


def cmd_relink_list(limit=None, het=None):
    """Az URL nélküli cikkek, amelyeknél még érdemes keresni.

    Kimenet: „cikkek" = a keresendők (legfeljebb `limit`), „kezi_lista" = amit a
    gép nem tud megkeresni (NEM_KERESHETO domain, vagy MAX_KERESES sikertelen
    kör után sincs meg) — ezeket a záró összefoglaló sorolja fel kézi pótlásra.
    Sorrend: előbb a legkevesebbszer keresettek, azon belül a legfrissebbek.
    Így a heti backlog-kör végigforog a listán, és nem ugyanazt a 10 tételt
    keresi minden hétfőn (2026-10-14 előtt a sikertelen keresés nyomtalan volt).
    """
    repo = clone_repo(read_only=True)
    try:
        db, _, _ = read_db((repo / "docs" / "index.html").read_text(encoding="utf-8"))
    finally:
        shutil.rmtree(repo, ignore_errors=True)

    todo, kezi = [], []
    for c in db.get("cikkek", []):
        if c.get("url") or c.get("url_forras") == "nincs":
            continue
        if int(c.get("url_probak", 0) or 0) >= MAX_PROBAK:
            continue
        if het and c.get("het") != het:
            continue
        dom = _SRC2DOM.get(_fold(c.get("forras", "")))
        alias = sorted(MEDIA_ALIASES.get(_fold(c.get("forras", "")), set()))
        domain_tipp = " VAGY ".join(alias) if alias else (dom or "")
        keresve = int(c.get("url_keresve", 0) or 0)
        tetel = {
            "id": c.get("id"), "cim": c.get("cim"), "forras": c.get("forras"),
            "datum": c.get("datum"), "het": c.get("het"),
            "domain_tipp": domain_tipp,
            "broadcast": _src_key(c.get("forras", "")) in BROADCAST,
            "probak": int(c.get("url_probak", 0) or 0),
            "keresve": keresve,
        }
        nem_keresheto = any(d in NEM_KERESHETO
                            for d in domain_tipp.split(" VAGY ") if d)
        if nem_keresheto or keresve >= MAX_KERESES:
            tetel["ok"] = ("a kereső nem éri el a domaint" if nem_keresheto
                           else f"{keresve} keresés, nincs találat")
            kezi.append(tetel)
        else:
            todo.append(tetel)
    # a legfrissebbek előre (frissen még megtalálhatók), de a ritkábban
    # keresettek mindig előbb — a két rendezés stabil, így mindkettő érvényes
    todo.sort(key=lambda x: x.get("datum") or "", reverse=True)
    todo.sort(key=lambda x: x["keresve"])
    kezi.sort(key=lambda x: x.get("datum") or "", reverse=True)
    osszes = len(todo)
    if limit:
        todo = todo[:int(limit)]
    print(json.dumps({"osszes_hianyzo": osszes + len(kezi),
                      "osszes_keresendo": osszes, "listazva": len(todo),
                      "cikkek": todo, "kezi_lista": kezi},
                     ensure_ascii=False, indent=2))
    if limit and osszes > len(todo):
        print(f"\nFIGYELEM: {osszes - len(todo)} további cikk vár link-pótlásra "
              f"(limit={limit}) — a következő körben jönnek.", file=sys.stderr)
    if kezi:
        print(f"KÉZI PÓTLÁSRA VÁR: {len(kezi)} tétel (kezi_lista) — sorold fel "
              f"a záró összefoglalóban címmel, forrással, dátummal.",
              file=sys.stderr)


def _canon_sources(db):
    """Forrásnév-egységesítés. Visszaad: (változások listája, FIGYELEM-sorok)."""
    import collections
    groups = collections.defaultdict(collections.Counter)
    for c in db.get("cikkek", []):
        groups[_src_key(c.get("forras", ""))][c.get("forras", "")] += 1
    changes, warns = [], []
    target = {}
    for key, spellings in groups.items():
        if not key:
            continue
        if key in CANON_SOURCE:
            target[key] = CANON_SOURCE[key]
        elif len(spellings) > 1:
            best = spellings.most_common(1)[0][0]
            target[key] = best
            warns.append(f"FIGYELEM: '{key}' több írásmóddal szerepel "
                         f"({dict(spellings)}) — a leggyakoribbat ('{best}') "
                         f"választottam. Ha más kell, vedd fel a CANON_SOURCE-ba.")
    for c in db.get("cikkek", []):
        key = _src_key(c.get("forras", ""))
        want = target.get(key)
        if want and c.get("forras") != want:
            changes.append((c.get("id"), c.get("forras"), want))
            c["forras"] = want
    return changes, warns


def _commit(repo, db, msg, dry_run):
    if dry_run:
        write_html(repo, db)   # lefut az önellenőrzés is
        print("DRY-RUN: a HTML újraírása és önellenőrzése sikeres, nincs commit/push.")
        return
    db["generated"] = datetime.date.today().isoformat()
    write_html(repo, db)
    _run(["git", "add", "docs/index.html"], cwd=repo)
    _run(["git", "commit", "-m", msg], cwd=repo)
    _push_with_retry(repo)
    print("PUSH OK")


def cmd_fix_sources(dry_run=False):
    """A forrásnevek egységesítése az egész DB-ben (külön futtatható)."""
    repo = clone_repo(read_only=dry_run)
    try:
        db, _, _ = read_db((repo / "docs" / "index.html").read_text(encoding="utf-8"))
        changes, warns = _canon_sources(db)
        for w in warns:
            print(w)
        for aid, old, new in changes:
            print(f"{aid}: {old!r} -> {new!r}")
        print(f"\nÖsszesen {len(changes)} cikk forrásneve egységesítve.")
        if not changes:
            print("Nincs teendő.")
            return
        _commit(repo, db, f"Forrásnevek egységesítése ({len(changes)} cikk)", dry_run)
    finally:
        shutil.rmtree(repo, ignore_errors=True)


def cmd_audit_links(dry_run=False):
    """A MÁR BENT LÉVŐ linkek auditja — a link-pótlás hiányzó fele.

    Miért kell: a `relink-*` kör csak a HIÁNYZÓ URL-eket kereste meg, a meglévő
    387-et senki nem validálta. Pedig a docx hyperlink-rId újrahasznosítása
    bizonyítottan rossz linket ragaszt cikkekhez (a WS-dokumentum minta-sorának
    blikk.hu-s linkje hat különböző cikk címéhez is hozzáragadt). Mérés
    2026-08-05: 449 cikk, 387 URL, ebből 17-nél a forrásnév ellentmond a
    domainnek.

    Amit MÓDOSÍT: kizárólag a `forras` mezőt, és csak ha a slug bizonyítja, hogy
    a link ehhez a cikkhez tartozik. URL-t sem törölni, sem átírni nem fog — az
    emberi döntés, mert egy rossz link eldobása adatvesztés.

    Amit csak JELEZ: a gyanús linkeket (a domain sem illik, a slug sem egyezik)
    és a 3 napnál nagyobb URL-dátum / `datum` eltéréseket.
    """
    repo = clone_repo(read_only=dry_run)
    try:
        db, _, _ = read_db((repo / "docs" / "index.html").read_text(encoding="utf-8"))
        changes, gyanus, datumelt = [], [], []
        for c in db.get("cikkek", []):
            url = c.get("url")
            if not url:
                continue
            if _domain(url) in NEM_CIKK_DOMAIN:
                continue          # videó / PDF-szkennelés — lásd NEM_CIKK_DOMAIN
            st, _ = _domain_ok(url, c.get("forras", ""))
            if st == "eltero":
                dnev = _source(url)
                if (dnev and _src_key(dnev) != _src_key(c.get("forras", ""))
                        and _link_bizonyitott(url, c.get("cim", ""))):
                    changes.append((c.get("id"), c.get("forras"), dnev, _domain(url)))
                    c["forras"] = dnev
                else:
                    gyanus.append((c.get("id"), c.get("forras"), _domain(url),
                                   c.get("cim", "")[:70]))
            ud = _url_date(url)
            if ud and c.get("datum"):
                try:
                    elt = abs((ud - datetime.date.fromisoformat(c["datum"])).days)
                except ValueError:
                    elt = 0
                if elt > 3:
                    datumelt.append((c.get("id"), c.get("datum"),
                                     ud.isoformat(), elt))

        for aid, old, new, dom in changes:
            print(f"{aid}: forras {old!r} -> {new!r}   (domain {dom}, slug egyezik)")
        for aid, forras, dom, cim in gyanus:
            print(f"FIGYELEM: {aid} gyanús link — forras {forras!r}, "
                  f"URL-domain {dom}, a slug NEM egyezik a címmel: {cim}")
        for aid, datum, ud, elt in datumelt:
            print(f"FIGYELEM: {aid} dátumeltérés — DB {datum}, URL {ud} "
                  f"({elt} nap)")
        print(f"\nÖsszesen: {len(changes)} forrásnév javítva, {len(gyanus)} gyanús "
              f"link kézi ellenőrzésre, {len(datumelt)} dátumeltérés.")
        if not changes:
            print("Nincs mit commitolni.")
            return
        _commit(repo, db, f"Link-audit: {len(changes)} forrásnév javítva az "
                          f"URL-domain alapján", dry_run)
    finally:
        shutil.rmtree(repo, ignore_errors=True)


# A repóba szinkronizált lokális fájlok. Az `apply` csak a docs/index.html-t
# commitolja, ezért a repóbeli cowork/ hónapokig elavult maradhat: a 2026-08-05-i
# audit szerint a repóbeli TASK_PROMPT.md bizonyíthatóan HAMIS állítást
# tartalmazott a kódról (》a cmd_scan iterdir()-t használ《, miközben már rglob
# volt) és hiányos alparancs-listát. Egy elavult instrukció nem elhasal, hanem
# félreinformálja a következő futást — ezért kell egy egylépéses szinkron.
COWORK_SYNC = ("runner.py", "vocab.py", "TASK_PROMPT.md", "SCHEDULED_TASK_PROMPT.md")
WORKFLOW_SRC = "watchdog.yml"
WORKFLOW_DST = ".github/workflows/sajtofigyelő.yml"


def cmd_sync_cowork(dry_run=False):
    """A lokális cowork/ fájlok és a watchdog feltöltése a repóba."""
    here = Path(__file__).resolve().parent
    repo = clone_repo(read_only=dry_run)
    try:
        changed = []
        for name, dst_rel in ([(n, f"cowork/{n}") for n in COWORK_SYNC]
                              + [(WORKFLOW_SRC, WORKFLOW_DST)]):
            src = here / name
            if not src.exists():
                print(f"FIGYELEM: {name} nincs meg lokálisan — kihagyva.")
                continue
            dst = repo / dst_rel
            txt = src.read_text(encoding="utf-8")
            if dst.exists() and dst.read_text(encoding="utf-8") == txt:
                print(f"változatlan: {dst_rel}")
                continue
            dst.parent.mkdir(parents=True, exist_ok=True)
            dst.write_text(txt, encoding="utf-8")
            changed.append(dst_rel)
            print(f"FRISSÍTVE:   {dst_rel}")
        if not changed:
            print("\nA repo már szinkronban van — nincs teendő.")
            return
        if dry_run:
            print(f"\nDRY-RUN: {len(changed)} fájl frissülne, nincs commit/push.")
            return
        _run(["git", "add", "--"] + changed, cwd=repo)
        _run(["git", "commit", "-m",
              f"cowork/ és a watchdog szinkronba tétele ({len(changed)} fájl)"],
             cwd=repo)
        _push_with_retry(repo)
        print("PUSH OK")
    finally:
        shutil.rmtree(repo, ignore_errors=True)


def cmd_relink_apply(json_path, dry_run=False, fix_sources=False):
    """Talált linkek (és sikertelen keresések) visszaírása id alapján.

    Bejegyzés-típusok a JSON 'linkek' listájában:
      {"id", "url", ["oldal_cim"], ["tipus": "kezi"|"docx"], ["indok"]}
                                         – link; validálás után beíródik
      {"id", "nincs_talalat": true}       – sikertelen keresés (url_keresve+1)
      {"id", "url_forras": "nincs", "indok": "…"}
                                         – lezárás; indok nélkül elutasítva
      {"id", "duplikatuma": "ART-…"}      – duplikátum törlése (_dup_ok)
    """
    payload = _read_json_retry(json_path)
    entries = payload["linkek"] if isinstance(payload, dict) else payload
    if not isinstance(entries, list):
        raise RuntimeError("A JSON-ban nincs lista ('linkek' kulcs vagy tömb).")

    repo = clone_repo(read_only=dry_run)
    try:
        db, _, _ = read_db((repo / "docs" / "index.html").read_text(encoding="utf-8"))
        by_id = {c.get("id"): c for c in db.get("cikkek", [])}
        db_urls = _existing_urls(db)
        url_owner = {str(c["url"]).strip().rstrip("/"): c.get("id")
                     for c in db.get("cikkek", []) if c.get("url")}

        stamped = _stamp_docx_source(db)
        if stamped:
            print(f'Visszamenőleges jelölés: {stamped} cikk url_forras="docx".')
        # A 2026-09-28-i késleltetett push egy elavult javaslatlistát alkalmazott:
        # nyolc, addigra már linkelt cikk „elutasítást" és url_probak-ot kapott.
        # Linkkel bíró cikknél a számlálónak nincs értelme.
        takaritva = 0
        for c in db.get("cikkek", []):
            if c.get("url") and "url_probak" in c:
                c.pop("url_probak")
                takaritva += 1
        if takaritva:
            print(f"Takarítás: {takaritva} linkkel bíró cikkről lekerült az url_probak.")

        beirt = lezart = elutasitott = torolt = 0
        kihagyva = utkozes = nincs_talalat = probak_no = 0
        ma = datetime.date.today().isoformat()
        for e in entries:
            aid = str(e.get("id", "")).strip()
            art = by_id.get(aid)
            if art is None:
                print(f"FIGYELEM: nincs ilyen id a DB-ben: {aid!r} — kihagyva.")
                elutasitott += 1
                continue
            if e.get("duplikatuma"):
                other = by_id.get(str(e["duplikatuma"]).strip())
                dok, dindok = _dup_ok(art, other)
                if dok:
                    db["cikkek"] = [c for c in db["cikkek"] if c.get("id") != aid]
                    torolt += 1
                    print(f"TÖRÖLVE   {aid}: {e['duplikatuma']} duplikátuma ({dindok})")
                else:
                    elutasitott += 1
                    print(f"ELUTASÍTVA {aid}: duplikátum-törlés megtagadva — {dindok}")
                continue
            if e.get("nincs_talalat"):
                # „Kerestem, nincs." 2026-10-14 előtt nyomtalan volt, ezért nem
                # látszott, hogy a backlog-kör egyáltalán keresett-e egy tételt.
                if art.get("url"):
                    kihagyva += 1
                    continue
                art["url_keresve"] = int(art.get("url_keresve", 0) or 0) + 1
                art["url_keresve_utoljara"] = ma
                nincs_talalat += 1
                print(f"NINCS     {aid}: {art['url_keresve']}. sikertelen keresés"
                      + (" — kézi listára kerül." if art["url_keresve"] >= MAX_KERESES
                         else ""))
                continue
            if (str(e.get("url_forras", "")).strip() == "nincs" and not e.get("url")
                    and not str(e.get("indok", "")).strip()):
                elutasitott += 1
                print(f"ELUTASÍTVA {aid}: „nincs” lezáráshoz 'indok' kell. "
                      f"(2026-08-05-én két TV-tétel egyetlen egzakt címkeresés után "
                      f"tévesen zárult le; az egyiknek van atv.hu-s változata.)")
                continue
            url_c = _clean_url(e.get("url")) if e.get("url") else None
            if url_c and art.get("url"):
                kihagyva += 1
                print(f"KIHAGYVA  {aid}: már van URL-je ({str(art['url'])[:50]}) — "
                      f"elavult javaslat, nem számít próbálkozásnak.")
                continue
            owner = url_owner.get(url_c.strip().rstrip("/")) if url_c else None
            if owner and owner != aid:
                utkozes += 1
                print(f"ÜTKÖZÉS   {aid}: {url_c} már {owner}-nél van. Ha a WS rossz "
                      f"tételhez tette, javítsd kézzel (kezi-javitas). Nem számít "
                      f"próbálkozásnak.")
                continue
            if str(e.get("tipus", "")).strip() == "docx":
                ok, forras, indok, warn = _validate_docx_link(url_c, art)
                if not ok:
                    # A WS linkje bizonyíthatóan rossz (pl. sablonból öröklött).
                    # NEM növeli az url_probak-ot: a docx-relink minden futásban
                    # újra javasolná, és a tétel így kiesne a keresésből.
                    elutasitott += 1
                    print(f"DOCX-LINK ELUTASÍTVA {aid}: {indok}")
                    continue
            else:
                ok, forras, indok, warn = _validate_link(e, art, db_urls)
            if warn:
                print(warn)
            if not ok:
                art["url_probak"] = int(art.get("url_probak", 0) or 0) + 1
                elutasitott += 1
                probak_no += 1
                print(f"ELUTASÍTVA {aid}: {indok}")
                if art["url_probak"] >= MAX_PROBAK:
                    print(f"  ({aid} elérte a {MAX_PROBAK} próbálkozást — "
                          f"a relink-list nem adja vissza többé.)")
                continue
            if forras == "nincs":
                art["url_forras"] = "nincs"
                for k in _LINK_ALLAPOT:
                    art.pop(k, None)
                lezart += 1
                print(f"LEZÁRVA   {aid}: nincs online változat ({art.get('forras')}) "
                      f"— {e.get('indok')}")
                continue
            url = _clean_url(e["url"])
            tipus = str(e.get("tipus", "")).strip()
            art["url"] = url
            art["url_forras"] = tipus if tipus in ("kezi", "docx") else "talalt"
            for k in _LINK_ALLAPOT:
                art.pop(k, None)
            db_urls.add(url.strip().rstrip("/"))
            url_owner[url.strip().rstrip("/")] = aid
            kanon = _source(url)
            forras_most = str(art.get("forras", ""))
            puszta_domain = bool(re.fullmatch(r"[\w.-]+\.\w{2,4}", forras_most))
            if kanon and puszta_domain and kanon != forras_most:
                print(f"  forrásnév a domain alapján: {forras_most!r} -> {kanon!r}")
                art["forras"] = kanon
            beirt += 1
            print(f"OK        {aid}: {url}\n            ({indok})")

        if fix_sources:
            changes, warns = _canon_sources(db)
            for w in warns:
                print(w)
            if changes:
                print(f"Forrásnév egységesítve {len(changes)} cikknél.")

        print(f"\nÖsszesen: {beirt} link beírva, {lezart} lezárva (nincs online), "
              f"{torolt} duplikátum törölve, {elutasitott} elutasítva, "
              f"{utkozes} ütközés, {kihagyva} kihagyva, "
              f"{nincs_talalat} sikertelen keresés naplózva.")
        if torolt:
            db["total"] = len(db["cikkek"])
        maradek = sum(1 for c in db.get("cikkek", [])
                      if not c.get("url") and c.get("url_forras") != "nincs")
        print(f"URL nélküli cikk a DB-ben: {maradek} / {len(db.get('cikkek', []))}")
        # Commit csak valódi DB-változásnál (az indok nélküli lezárás vagy az
        # ismeretlen id elutasítása nem változtat semmin).
        if not (beirt or lezart or stamped or probak_no or torolt
                or nincs_talalat or takaritva):
            print("Nincs változás — nincs commit.")
            return
        msg = (f"Link-pótlás: +{beirt} link, {lezart} lezárva, "
               f"{torolt} duplikátum törölve")
        if nincs_talalat:
            msg += f", {nincs_talalat} sikertelen keresés"
        _commit(repo, db, msg, dry_run)
    finally:
        shutil.rmtree(repo, ignore_errors=True)


def _validate_docx_link(url, art):
    """A docx-ből jött link ellenőrzése — ugyanaz a bizalmi szint, mint az
    apply-ban (_normalize): a WS linkje hiteles, kivéve ha bizonyíthatóan rossz
    (a domain nem illik a forráshoz és a slug sem a címhez, vagy a URL dátuma
    több mint 3 nappal eltér). Videó / PDF-szkennelés (NEM_CIKK_DOMAIN) rendben."""
    if not url:
        return False, None, "nem http(s) URL", None
    dom = _domain(url)
    if dom in NEM_CIKK_DOMAIN:
        return True, "docx", f"docx-link ({dom}: videó / PDF-szkennelés)", None
    st, exp = _domain_ok(url, art.get("forras", ""))
    if st == "eltero" and not _link_bizonyitott(url, art.get("cim", "")):
        return False, None, (f"a docx-link domainje ({dom}) nem illik a forráshoz "
                             f"({art.get('forras')}), és a slug sem egyezik a címmel"), None
    try:
        adate = datetime.date.fromisoformat(art["datum"])
    except Exception:
        return False, None, "a cikk datum mezője érvénytelen", None
    ud = _url_date(url)
    if ud and abs((ud - adate).days) > 3:
        return False, None, f"URL-dátum {ud} vs cikk dátum {adate} (több mint 3 nap)", None
    return True, "docx", "docx-link (a WS körképéből)", None


_KORKEP_NEV = re.compile(r"(?i)^HR[ -]?piaci\s+k")


def cmd_docx_relink(folder, dry_run=False):
    """A link nélküli cikkek linkje a körképekből (docx), keresés nélkül.

    2026-10-14 előtt a kinyerés nem olvasta a Word-mezőkódos linkeket, ezért a
    WS által megadott linkek egy része nem került be. Ez a parancs a mappa
    (rekurzívan) összes „HR piaci körkép" docx-ét újra kinyeri, és a link
    nélküli cikkekhez a docx linkjét adja be a relink-apply-nak (tipus=docx).
    Meglévő URL-t nem ír felül. Idempotens: ha nincs mit pótolni, nincs commit.
    """
    src = Path(folder)
    if not src.is_dir():
        raise RuntimeError(f"Nincs ilyen mappa: {folder}")
    files = sorted(p for p in src.rglob("*.docx")
                   if p.is_file() and not p.name.startswith("~")
                   and _KORKEP_NEV.match(p.name))
    tetelek = []
    for f in files:
        try:
            for t in _ws_kivonatok(extract_docx(str(f))):
                if t.get("url"):
                    t["fajl"] = f.name
                    tetelek.append(t)
        except Exception as e:
            print(f"FIGYELEM: {f.name} nem olvasható ({e}) — kihagyva.")
    print(f"{len(files)} körkép, {len(tetelek)} linkes tétel.")
    repo = clone_repo(read_only=True)
    try:
        db, _, _ = read_db((repo / "docs" / "index.html").read_text(encoding="utf-8"))
    finally:
        shutil.rmtree(repo, ignore_errors=True)
    entries = []
    for c in db.get("cikkek", []):
        if c.get("url") or c.get("url_forras") == "nincs":
            continue
        t = _ws_match(c, tetelek)
        if t:
            entries.append({"id": c["id"], "url": t["url"], "tipus": "docx",
                            "indok": f"docx-relink: {t['fajl']}"})
    if not entries:
        print("A körképekben nincs link a link nélküli cikkekhez — nincs teendő.")
        return
    nev = f"sajto_docx_relink_{datetime.date.today().isoformat()}.done.json"
    hova = Path(tempfile.mkdtemp(prefix="sajto_dr_")) if dry_run else WORK
    hova.mkdir(parents=True, exist_ok=True)
    out = hova / nev
    out.write_text(json.dumps({"linkek": entries}, ensure_ascii=False, indent=1),
                   encoding="utf-8")
    print(f"{len(entries)} javaslat a körképekből: {out}")
    cmd_relink_apply(str(out), dry_run=dry_run)


def cmd_kezi_javitas(json_path, dry_run=False):
    """Kézi link-javítás id alapján — meglévő URL-t is felülír vagy töröl.

    Mikor kell: ha a WS rossz tételhez tette a linket. 2026-08-14: ART-0474 a
    0479 cikkének URL-jét kapta, ezért a 0479 helyes linkjét a relink-apply
    „már szerepel" indokkal eldobta. A relink-apply szándékosan nem ír felül
    meglévő URL-t — ez a külön, kézi út, kötelező indoklással.

    JSON: {"javitasok": [{"id": "ART-0474", "url": "https://…" | null,
                          "indok": "…"}]}
    url: null = a link törlése; a tétel ezután a link-pótlásba kerül.
    A javítások után ellenőrzi, hogy egy URL ne legyen két cikknél; ha mégis,
    semmi nem kerül commitolásra.
    """
    payload = _read_json_retry(json_path)
    entries = payload.get("javitasok") if isinstance(payload, dict) else payload
    if not isinstance(entries, list) or not entries:
        raise RuntimeError("A JSON-ban nincs 'javitasok' lista.")
    repo = clone_repo(read_only=dry_run)
    try:
        db, _, _ = read_db((repo / "docs" / "index.html").read_text(encoding="utf-8"))
        by_id = {c.get("id"): c for c in db.get("cikkek", [])}
        ma = datetime.date.today().isoformat()
        n = 0
        for e in entries:
            aid = str(e.get("id", "")).strip()
            art = by_id.get(aid)
            if art is None:
                raise RuntimeError(f"Nincs ilyen id a DB-ben: {aid!r}")
            indok = str(e.get("indok", "")).strip()
            if not indok:
                raise RuntimeError(f"{aid}: az 'indok' kötelező (auditálhatóság).")
            regi = art.get("url")
            uj = _clean_url(e.get("url")) if e.get("url") else None
            if e.get("url") and not uj:
                raise RuntimeError(f"{aid}: nem http(s) URL: {e.get('url')!r}")
            if uj:
                st, exp = _domain_ok(uj, art.get("forras", ""))
                if st == "eltero":
                    print(f"FIGYELEM: {aid} domain {_domain(uj)} ≠ {exp} "
                          f"(forrás: {art.get('forras')}) — kézi döntés, beírom.")
                art["url"] = uj
                art["url_forras"] = "kezi"
            else:
                art["url"] = None
                art.pop("url_forras", None)
            for k in _LINK_ALLAPOT:
                art.pop(k, None)
            n += 1
            print(f"JAVÍTVA   {aid}: {regi or '—'}\n            -> {uj or '— (törölve)'}")
        seen = {}
        for c in db.get("cikkek", []):
            u = str(c.get("url") or "").strip().rstrip("/")
            if not u:
                continue
            if u in seen:
                raise RuntimeError(f"Ütközés a javítás után: {u} — {seen[u]} és "
                                   f"{c.get('id')}. Semmi nem lett commitolva.")
            seen[u] = c.get("id")
        print(f"\nÖsszesen: {n} tétel javítva; URL-ütközés nincs.")
        _commit(repo, db, f"Kézi linkjavítás: {n} tétel", dry_run)
    finally:
        shutil.rmtree(repo, ignore_errors=True)


def cmd_ws_backfill(dry_run=False):
    """A WS-kivonatok (kivonat mező) visszatöltése a work/*.txt-kből.

    A 2026-04-02 előtti tételekhez nincs kinyert txt, ezekhez nem lesz kivonat.
    """
    files = sorted(WORK.glob("sajto_input_*.txt"))
    tetelek = []
    for f in files:
        tetelek += _ws_kivonatok(f.read_text(encoding="utf-8"))
    print(f"{len(files)} körkép-kinyerés, {len(tetelek)} WS-tétel.")
    repo = clone_repo(read_only=dry_run)
    try:
        db, _, _ = read_db((repo / "docs" / "index.html").read_text(encoding="utf-8"))
        n = nincs = 0
        for c in db.get("cikkek", []):
            t = _ws_match(c, tetelek)
            if t is None:
                nincs += 1
                continue
            if _set_ws_kivonat(c, t):
                n += 1
        print(f"WS-kivonat beállítva/frissítve: {n} cikk; párja nincs: {nincs} "
              f"(főleg a 2026 áprilisa előtti, txt nélküli tételek).")
        if not n:
            print("Nincs változás — nincs commit.")
            return
        _commit(repo, db, f"Összefoglalók visszatöltése: {n} cikk", dry_run)
    finally:
        shutil.rmtree(repo, ignore_errors=True)


if __name__ == "__main__":
    a = sys.argv[1:]
    try:
        if a and a[0] == "scan" and len(a) <= 2:
            cmd_scan(a[1] if len(a) == 2 else None)
        elif a and a[0] == "extract" and len(a) == 2:
            cmd_extract(a[1])
        elif a and a[0] == "relink-list":
            o = dict(z.split("=", 1) for z in a[1:] if "=" in z)
            cmd_relink_list(limit=o.get("--limit"), het=o.get("--het"))
        elif a and a[0] == "relink-apply" and len(a) >= 2:
            cmd_relink_apply(a[1], dry_run="--dry-run" in a[2:],
                             fix_sources="--fix-sources" in a[2:])
        elif a and a[0] == "fix-sources":
            cmd_fix_sources(dry_run="--dry-run" in a[1:])
        elif a and a[0] == "audit-links":
            cmd_audit_links(dry_run="--dry-run" in a[1:])
        elif a and a[0] == "sync-cowork":
            cmd_sync_cowork(dry_run="--dry-run" in a[1:])
        elif a and a[0] == "kezi-javitas" and len(a) >= 2:
            cmd_kezi_javitas(a[1], dry_run="--dry-run" in a[2:])
        elif a and a[0] == "ws-backfill":
            cmd_ws_backfill(dry_run="--dry-run" in a[1:])
        elif a and a[0] == "docx-relink" and len(a) >= 2:
            cmd_docx_relink(a[1], dry_run="--dry-run" in a[2:])
        elif a and a[0] == "apply" and len(a) == 3:
            cmd_apply(a[1], a[2])
        else:
            print(__doc__)
            sys.exit(1)
    except Exception as e:
        print(f"HIBA: {_redact(str(e))}", file=sys.stderr)
        sys.exit(1)
