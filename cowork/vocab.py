# -*- coding: utf-8 -*-
"""
Sajtófigyelő – kontrollált címketár (vokabulárium) + leképezés.
=================================================================
Miért van erre szükség: a címkék korábban szabad szövegként keletkeztek
(TASK_PROMPT „2–5 releváns magyar kulcsszó"), aminek eredménye 435 cikkre
588 különböző címke lett, ebből 389 pontosan EGY cikkhez tartozott. Így a
szűrőfelület nem szűr, csak zajol.

Ez a modul a szűrésre használható, ZÁRT címkelistát tartalmazza, plusz a
leképezést a korábbi szabad címkékről. Amit nem lehet a listára képezni, az
ELDOBÁSRA kerül — a hosszú farkat (egyszeri cégnevek, városok, személyek) a
szabadszavas kereső fedi, ami a címre + forrásra + összefoglalóra keres.

Kiegészítés menete, ha egy új szereplő visszatérővé válik: vedd fel a
VOCAB-ba, és tedd be az ALIASES-be az írásmód-variánsait.
"""

# ── A ZÁRT CÍMKELISTA ────────────────────────────────────────────────────────
# csoport -> címkék.  A csoport csak dokumentációs célú, a JSON-ban nem jelenik meg.

VOCAB_GROUPS = {
    "téma": [
        "AI",
        "automatizáció",
        "adózás",
        "beruházás",
        "bérek",
        "bértranszparencia",
        "demográfia",
        "diákmunka",
        "esélyegyenlőség",
        "foglalkoztatottság",
        "gazdasági kilátások",
        "idősebb munkavállalók",
        "képzés",
        "kutatás és statisztika",
        "leépítés",
        "megtartás",
        "mentális egészség",
        "munkaerőhiány",
        "munkaerő-kölcsönzés",
        "munkaerő-tartalék",
        "munkajog",
        "munkakörülmények",
        "munkanélküliség",
        "munkáltatói márka",
        "munkáltatói szövetség",
        "nők a munkaerőpiacon",
        "pályakezdők",
        "politika",
        "régiós különbségek",
        "rugalmas foglalkoztatás",
        "skills-alapú HR",
        "szakszervezet",
        "szezonális munka",
        "toborzás",
        "vendégmunka",
        "vendégmunkás-stop",
    ],
    "szektor": [
        "agrárium",
        "autóipar",
        "egészségügy",
        "építőipar",
        "ipar",
        "IT",
        "kereskedelem",
        "logisztika",
        "turizmus",
    ],
    "szereplő": [
        "Trenkwalder",
        "WHC",
        "Randstad",
        "Adecco",
        "Manpower",
        "Gi Group",
        "Work Force",
        "HR-Rent",
        "Pensum Group",
        "Menton Jobs",
        "Job Force",
        "Man at Work",
        "Pannon-Work",
        "Jobtain",
        "Meló-Diák",
        "Mind-Diák",
        "Prodiák",
        "Profession.hu",
        "Giggle",
        "KSH",
    ],
}

VOCAB = {t: g for g, ts in VOCAB_GROUPS.items() for t in ts}

# ── LEKÉPEZÉS: cél címke -> a régi, szabad címkék, amik ide tartoznak ────────
# Ami egyik listán sincs és nem is vokabulárium-elem, azt a normalizálás eldobja.

_TO = {
    # ---- téma -------------------------------------------------------------
    "AI": ["mesterséges intelligencia", "AI készségek", "AI-képzés",
           "prediktív analitika", "adatvezérelt HR", "oktatási robotok"],
    "automatizáció": ["robotika", "robotizáció", "robotizálás", "digitalizáció",
                      "digitális készségek"],
    "adózás": ["adócsökkentés", "adókedvezmény", "szja", "szja-bevallás",
               "szja-mentesség", "NAV", "KATA", "kata", "cafeteria", "juttatás",
               "juttatások", "bértámogatás", "bérköltség", "többletköltség",
               "osztalék", "ekho"],
    "beruházás": ["befektetés", "FDI", "gyáravatás", "akkugyár",
                  "transzformátorgyár", "beruházásösztönzés", "tőkebeáramlás",
                  "K+F", "kínai beruházás", "HIPA", "munkahelyteremtés",
                  "létszámbővítés", "létszámbővülés", "EU támogatás",
                  "támogatás", "gazdaságfejlesztés", "MAPI", "EBRD", "Paks II.",
                  "relokáció", "beszállító", "beszállítók"],
    "bérek": ["bér", "béremelés", "béremelkedés", "órabér", "minimálbér",
              "átlagbér", "átlagkereset", "átlagkeresetek", "bérfelzárkózás",
              "reálkereset", "reálbér", "fizetés", "kereset", "mediánbér",
              "bérstatisztika", "béradatok", "bérnövekedés", "bérnyomás",
              "bérigény", "bérprémium", "bérfordulat", "fizikai munkabér",
              "IT-bérek", "agrárbérek", "tanárbérek", "bérelégedettség",
              "közepes jövedelem csapda", "fizetéscsökkentés"],
    "bértranszparencia": ["bértitok", "EU irányelv"],
    "demográfia": ["kivándorlás", "migráció", "bevándorlás", "hazatérés",
                   "mobilitás", "külföldi munkavállalás", "külföldi karrier",
                   "szegénység", "depriváció", "AROPE", "családpolitika",
                   "háztartások"],
    "diákmunka": ["iskolaszövetkezet"],
    "esélyegyenlőség": ["inklúzió", "diverzitás", "neurodiverzitás",
                        "megváltozott munkaképesség", "megváltozott munkaképességű",
                        "rehabilitációs hozzájárulás", "roma esélyegyenlőség",
                        "kordiszkrimináció", "foglalkoztatási program",
                        "Apa Akadémia", "apai szabadság"],
    "foglalkoztatottság": ["foglalkoztatás", "létszámadatok", "munkahelyek",
                           "munkahely", "létszámváltozás", "foglalkoztatáspolitika",
                           "munkavállalók", "munkavállalás", "létszámtervezés",
                           "közmunkaprogram", "munkaügy", "munkaerőpiac",
                           "bújtatott foglalkoztatás"],
    "gazdasági kilátások": ["GDP", "gazdasági növekedés", "konjunktúra",
                            "előrejelzés", "üzleti hangulat", "vállalati bizalom",
                            "fogyasztói bizalom", "bizonytalanság",
                            "gazdasági stagnálás", "növekedési modell",
                            "növekedési csapda", "versenyképesség", "termelékenység",
                            "humántőke", "export", "exportőrök", "euró",
                            "eurócsatlakozás", "erős forint", "forinterősödés",
                            "infláció", "inflació", "költségvetés",
                            "költségvetési hiány", "GDP-kiesés", "reform",
                            "reformjavaslat", "gazdaságpolitika", "gazdasági modell",
                            "globalizáció", "verseny", "eredmény", "iparszerkezet",
                            "tudásalapú társadalom", "gazdálkodás", "megtakarítás",
                            "pénzügyi tartalék", "váratlan kiadás", "piacvezető",
                            "nagyvállalatok", "kkv", "kisvállalkozás",
                            "állami cégek", "trendek"],
    "idősebb munkavállalók": ["senior munkavállalók", "40 felettiek", "Férfi40",
                              "Nők40", "nyugdíj", "The Seniors",
                              "szendvicsgeneráció", "idősellátás"],
    "képzés": ["oktatás", "átképzés", "felnőttképzés", "szakképzés",
               "duális képzés", "élethosszig tartó tanulás", "tanulás", "STEM",
               "skills-alapú képzés", "akkreditáció", "NAT", "mentorprogram",
               "katolikus iskola", "PROGmasters", "Tanfolyamguru",
               "Lumen Educationis", "SkillCompass", "oktatási minisztérium",
               "államtitkárság", "WIN-projekt", "BME roadshow",
               "Pécsi Tudományegyetem", "munkatapasztalat"],
    "kutatás és statisztika": ["felmérés", "statisztika", "rangsor", "elemzők",
                               "kutatás", "Provident Barométer", "VOSZ Barométer",
                               "IMD rangsor", "Diplomás Pályakövetési Rendszer",
                               "GKI", "Eurostat", "MNB", "OECD", "WEF",
                               "Oeconomus", "Political Capital", "PwC",
                               "RSM Hungary", "GKI Digital", "Dani Rodrik",
                               "Morgan Stanley", "China Labor Watch"],
    "leépítés": ["elbocsátás", "létszámleépítés", "bezárás", "átszervezés"],
    "megtartás": ["fluktuáció", "munkaerő-megtartás", "tehetségmegtartás",
                  "munkahelyváltás", "karrierváltás", "pályamódosítás",
                  "karrierépítés", "skill-alapú karrier", "elégedettség"],
    "mentális egészség": ["kiégés", "stressz", "pénzügyi stressz",
                          "munkapszichológia", "szabadság-szorongás",
                          "életminőség", "félelem", "kánikula"],
    "munkaerőhiány": ["szakemberhiány", "szakmunkáshiány", "mérnökhiány",
                      "sofőrhiány", "ápolóhiány", "hiányszakmák",
                      "kamionsofőr-hiány", "tanárhiány", "targoncavezető",
                      "full-stack fejlesztő", "légiutaskísérő", "mérnökök"],
    "munkaerő-kölcsönzés": ["kölcsönzés", "vendégmunkás-kölcsönzés",
                            "gig economy", "munkaerő-közvetítés",
                            "HR szolgáltatás", "MMOSZ", "MVI"],
    "munkaerő-tartalék": ["munkaerőtartalék", "strukturális munkanélküliség"],
    "munkajog": ["szabályozás", "munka törvénykönyve", "rendelet",
                 "EU szabályozás", "munkavállalói jogok", "munkavédelem",
                 "tilalom", "tiltás", "stop", "szigorítás", "moratórium",
                 "engedélyezés", "engedély", "engedélyek", "idegenrendészet",
                 "tartózkodási engedély", "bevándorlási szabályozás",
                 "kitiltás", "kitoloncolás", "kötelezettségszegés",
                 "határozott idejű szerződés", "kiskapu", "ügyintézés",
                 "kártalanítás", "szabadság", "munkaidő", "kvóta"],
    "munkakörülmények": ["kizsákmányolás", "munkahelyi kultúra",
                         "vállalati kultúra", "HR-kommunikáció"],
    "munkanélküliség": ["álláskeresés"],
    "munkáltatói márka": ["employer branding", "tehetségvonzás",
                          "HR-stratégia", "HR-trendek", "HR", "HRKOMM Award"],
    # Munkáltatói / ágazati érdekképviseletek. Szándékosan NEM a
    # 'szakszervezet' címke alá kerülnek: a VOSZ, MAGOSZ, MMOSZ a munkáltatói
    # oldal, összemosni őket a munkavállalói érdekképviselettel tárgyi hiba.
    "munkáltatói szövetség": ["VOSZ", "MAGOSZ", "MOSZ", "VIMOSZ", "MŰISZ",
                              "MMOSZ", "HBLF", "FruitVeB", "AmCham", "DUIHK",
                              "MVI", "szövetség"],
    "nők a munkaerőpiacon": ["nők", "nemek", "bérszakadék", "nemek közötti bérkülönbség",
                             "bérkülönbség", "nemek közötti szakadék",
                             "női vezetők", "női kvóta", "nők foglalkoztatása",
                             "kisgyermekes nők", "nők a munkaerőpiacon",
                             "nők munkaerőpiaca", "nők a technológiában",
                             "esélyegyenlőség nők"],
    "pályakezdők": ["Z generáció", "generációk", "diploma", "diplomások aránya",
                    "mérnökhallgatók", "fiatalok"],
    "politika": ["Tisza Párt", "Tisza", "Tisza-kormány", "Fidesz", "Mi Hazánk",
                 "választás", "kormányváltás", "kormányprogram", "korrupció",
                 "trollfarm", "politikai elemzés", "Magyar Péter",
                 "Schiffer András", "Botka László", "Sára Botond",
                 "befolyással üzérkedés", "kormányszóvivő", "szociális miniszter",
                 "államtitkár", "kormányzati struktúra", "elnökségváltás",
                 "vezetőváltás", "Naderi Zsuzsanna", "Farkas Bertalan Péter",
                 "Joó István", "MI Hazánk"],
    "régiós különbségek": ["régió", "régiós különbség", "régiós munkaerőpiac",
                           "területi különbség", "Adria-régió", "Balkán",
                           "Baltikum", "Erdély", "Székelyföld", "Románia",
                           "Szlovákia", "Lengyelország", "Németország",
                           "német cégek", "Európa", "EU", "LHH"],
    "rugalmas foglalkoztatás": ["home office", "távmunka", "részmunkaidő",
                                "rugalmasság", "rugalmas munka",
                                "munka-magánélet egyensúly", "workation",
                                "irodai munka", "sabbatical", "IWG",
                                "rugalmas foglalkoztatás"],
    "skills-alapú HR": ["skills-first", "kompetencia", "képzettség",
                        "emberi készségek", "skills-alapú HR", "FEOR"],
    "szakszervezet": ["érdekképviselet", "érdekegyeztetés", "egyeztetés",
                      "sztrájk", "MASZSZ", "KASZ"],
    "szezonális munka": ["idénymunka", "szezonmunka", "szezonális foglalkoztatás",
                         "szezonális munkaerő", "nyári szezon", "nyár",
                         "nyári toborzás"],
    "toborzás": ["álláshirdetés", "álláshirdetések", "állásbörze", "karriernap",
                 "kékgalléros", "fizikai dolgozók", "szellemi dolgozók"],
    "vendégmunka": ["harmadik ország", "harmadik országbeli munkavállalók",
                    "külföldi munkavállalók", "vendégmunkások", "nepáli",
                    "filippínó", "filippínók", "Fülöp-szigetek",
                    "ukrán munkaerő", "kínai munkavállalók", "integráció",
                    "JD.com"],
    "vendégmunkás-stop": ["vendégmunkásstop"],

    # ---- szektor ----------------------------------------------------------
    "agrárium": ["mezőgazdaság", "agrárexport", "kertészet", "hajtatókertészet",
                 "zöldség-gyümölcs", "élelmiszeripar", "tejfeldolgozó ipar",
                 "Master Good", "élelmiszer-ipar"],
    "autóipar": ["elektromos átállás", "gumiabroncs", "BMW", "BYD",
                 "Mercedes-Benz", "Mercedes", "Hankook", "thyssenkrupp",
                 "Samsung"],
    "egészségügy": ["gyógyszeripar", "Richter Gedeon", "Egis", "idősápolás"],
    "építőipar": [],
    "ipar": ["feldolgozóipar", "gyártóipar", "alumíniumipar", "műanyagipar",
             "Ongropack", "Hidrofilt", "Fastron Hungária",
             "Kókai Tömítéstechnikai", "Emerson", "Siemens Energy",
             "vízkezelés", "LEGO", "Ball Corporation"],
    "IT": ["IT szektor", "IT munkaerőpiac", "kiberbiztonság", "ESET",
           "Magyar Telekom", "AWS", "Europion", "Vulcan Shield", "BinX"],
    "kereskedelem": ["kiskereskedelem", "SPAR", "Tesco", "IKEA", "MediaMarkt"],
    "logisztika": ["fuvarozás", "ellátásilánc-menedzsment", "Waberer's",
                   "Trans-Sped"],
    "turizmus": ["vendéglátás", "HORECA", "Balaton", "Wizz Air"],

    # ---- szereplő ---------------------------------------------------------
    "WHC": ["WHC Payroll", "WHC People & Culture"],
    "Pensum Group": ["Pensum"],
    "Prodiák": ["Prohuman Diákmunka"],
}

# régi címke -> cél vokabulárium-címke(k)
ALIASES: dict[str, tuple[str, ...]] = {}
for target, olds in _TO.items():
    assert target in VOCAB, f"ismeretlen cél: {target}"
    for o in olds:
        ALIASES.setdefault(o, ())
        ALIASES[o] = tuple(dict.fromkeys(ALIASES[o] + (target,)))

# ── Kifejezetten ELDOBOTT címkék ────────────────────────────────────────────
# Nem hiba, hanem döntés. Két csoport:
#  1) SZÓVIVŐK ÉS SZEMÉLYEK — a cégcímke már lefedi őket (Nógrádi József =
#     Trenkwalder), a szabadszavas kereső pedig névre is talál.
#  2) EGYSZERI CÉGEK, VÁROSOK, RENDEZVÉNYEK — egy cikkhez tartozó szűrő nem szűrő.
# A 'Prohuman' címke is ide tartozik: a főkategória-pill („Prohuman
# megjelenések") pontosan ugyanezt a halmazt adja, csak megbízhatóbban.

DROP_EXPLICIT = {"Prohuman", "Prohuman Learning Solutions", "Techtogether"}


def normalize_tags(tags, *, extra=()):
    """Szabad címkelista -> vokabulárium-címkék, sorrendtartóan, duplikátum nélkül."""
    out: list[str] = []
    for t in list(tags) + list(extra):
        t = str(t).strip()
        if not t or t in DROP_EXPLICIT:
            continue
        if t in VOCAB:
            cands = (t,)
        elif t in ALIASES:
            cands = ALIASES[t]
        else:
            continue
        for c in cands:
            if c not in out:
                out.append(c)
    return out
