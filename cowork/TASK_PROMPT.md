# Sajtófigyelő – heti feldolgozás (scheduled task instrukció)

Ez a referencia-instrukció a hétfői és pénteki scheduled taskhoz.
A projektmappa Windows-útvonala: `C:\Users\toth.gellert\Claude\Projects\Sajtófigyelés`
A bash sandboxban: `/sessions/<session>/mnt/Sajtófigyelés/` (a pontos mount-útvonal a session-leírásban).
Lent `<mappa>` = a sandbox-útvonal. A runnert MINDIG teljes útvonallal hívd.

**Verzió: 2026-10-14g.** A `b`-ben jött a magyar stílusszabály-blokk, a `c`-ben a
zárt **Címketár** (a `kulcsszavak` mező nem szabad szöveg, a runner validálja).
A `d`-ben egyszerűsítés: a bemeneti mappa szabadon megválasztható (`scan
[mappa]`), és kikerült a OneDrive-connector-alapú észlelés — mérve nem működött.
Az `e`-ben a **link-pótlás** (4. lépés, `relink-list` / `relink-apply`), az
`url_forras` mező és a `fix-sources` alparancs.

Az **`f`-ben három javítás** (2026-08-05-i pipeline-audit):

1. Az `apply` **nem nyeli el többé a hibás körképet.** Ha bármelyik cikk kiesik
   a validáción, a jó cikkek bekerülnek és ki lesznek pusholva, de a docx NEM
   lesz feldolgozottként megjelölve, és a runner hibakóddal áll le. Korábban
   ilyenkor is bekerült a `processed.json`-ba, tehát a körkép véglegesen
   elveszett. **Ilyenkor a te dolgod: javítsd a jelzett cikkeket a JSON-ban, és
   futtasd újra ugyanazt az `apply`-t.** Ez biztonságos, a merge dedupál.
2. Új alparancs: **`audit-links`** — a MÁR bent lévő linkek auditja
   (forrásnév kontra URL-domain). **NEM a heti futás része**, ne hívd meg
   magadtól: eseti karbantartó kör, Gellért indítja.
3. Az `apply` figyelmeztet, ha a `forras` ellentmond az URL domainjének.

A **`g`-ben (2026-10-14)** a link nélküli cikkek felmérése után:

1. Az `extract` a **Word-mezőkódos linkeket** (`HYPERLINK "…"`) is kiolvassa,
   és a WS-sablon „Cím (Médium, …)" mintasorát kihagyja. Mérve: a 07-31-i
   körképben a WS mind a 14 tételt belinkelte, a régi kinyerés 2-t elvesztett.
   A link nélkül érkező cikkek fő oka ez volt, nem a WS.
2. Az `apply` a cikkekhez eltárolja a **WS saját kivonatát** (`kivonat`
   mező) a `work/sajto_input_*.txt`-ből. Neked ehhez nincs teendőd; a kiírt
   `WS-kivonat: N / M` sort add tovább, ha N < M.
3. **Link-pótlás:** a sikertelen keresést naplózni kell (`nincs_talalat`), a
   `relink-list` külön `kezi_lista`-t ad (ezeket NE keresd), a TV-tételeket
   egy egzakt címkeresés után NEM szabad lezárni, `nincs` lezárás csak
   `indok`-kal. Részletek: 4/B–4/G.
**Fontos (2026-10-14, Gellért kérése):** a publikus oldalon nem szerepelhet
utalás a körkép forrására. A felületen a kivonat neve egyszerűen „Összefoglaló",
a mezőnév `kivonat`. Commit-üzenetbe és a DB-be se írj ilyen utalást vagy belső
jegyzetet — az `indok` csak a `work/*.json`-ban él.
4. Új alparancsok: `kezi-javitas`, `ws-backfill`. **Egyik sem a heti futás
   része**, ne hívd meg magadtól — Gellért kéri.
5. Új lépés a feldolgozás végén: **`docx-relink "<b>"`** (Feldolgozás 5.). A
   bemeneti mappa összes körképéből — a régiekből is — pótolja a link nélküli
   cikkek linkjét, ha a WS megadta. Keresés nélkül, idempotensen.

---

## ⚠ A LEGFONTOSABB MEGKÖTÉS

A cikkek linkjei **Word-hyperlinkek**, nem kiírt URL-ek. A `[URL: ...]` jelölést
**a runner.py maga gyártja** a docx hyperlink-relationshipjeiből és (2026-10-14
óta) a Word-mezőkódos `HYPERLINK`-ekből (`extract_docx`). Ez kizárólag a lokális, python-docx alapú `extract` során jön
létre — tehát a fájl **bináris tartalmának** elérhetőnek kell lennie egy csatolt
mappában.

Mért tény (2026-08-04): a Microsoft 365 connector (`read_resource`) szövegéből a
14 URL-ből **0** marad meg, és a szöveg is szétszabdalva jön
(`V ersenytársakra`, `2026. 04 . 01 .`). Ezért **TILOS** `apply`-t hívni olyan
JSON-ra, ami connector-szövegből készült.

---

## Lépések

1. `pip install python-docx --break-system-packages` (ha még nincs telepítve).
2. **Bemeneti mappa** meghatározása.
3. **Feldolgozás** fájlonként.
4. **Link-pótlás** — a hiányzó cikk-linkek megkeresése (lásd lent).
5. **Zárás** és ellenőrzés.

---

## Bemeneti mappa

A `.docx` körképek **egy lemezen létező, csatolt mappában** vannak. Két eset:

- a projektmappa `incoming/` almappája, vagy
- egy külön csatolt mappa (pl. egy lokálisan szinkronizált OneDrive-mappa).

Listázd a `/sessions/<session>/mnt/` tartalmát. Ha a `Sajtófigyelés` mellett van
másik csatolt mappa, és abban `HR piaci körkép ....docx` nevű fájlok vannak, az
a bemenet. Legyen `<b>` = ez a mappa; ha nincs ilyen, `<b>` = `<mappa>/incoming`.

A `scan` opcionálisan mappát fogad, tehát a bemenet bárhol lehet:
`runner.py scan "<b>"`. A dedupe fájlNÉV szerint megy a
`state/processed.json` ellen, függetlenül attól, melyik mappából jött a fájl —
tehát mappát váltani biztonságos, semmi nem lesz kétszer feldolgozva.

**Amit NE tegyél:** ne próbáld a Microsoft 365 connectorral (`read_resource`)
kinyerni a dokumentum szövegét. Mérve: a 14 URL-ből 0 marad meg, és a szöveg is
szétszabdalva jön (`V ersenytársakra`, `2026. 04 . 01 .`). A `processed.json`
pedig véglegesen lezárná a fájlt. Csak a lokális `extract` használható.

---

## Feldolgozás

1. `python3 -B <mappa>/cowork/runner.py scan "<b>"`

2. Ha `NINCS_UJ`: nincs teendő, egysoros jelentéssel fejezd be. (A körkép
   jellemzően péntek délután–kora este érkezik; ha még nincs itt, valószínűleg
   csak késik, és a következő futás elkapja.)

3. A `FIGYELEM: ... .doc formátum` sorokat add tovább.

4. Minden új fájlra (a `scan` sorrendjében):

   a. `python3 -B <mappa>/cowork/runner.py extract "<fájlnév>"`
      → kiírja: `OK <útvonal> (N karakter)`. A kimenet a `work/` mappába kerül,
      docx-enként egyedi néven: `work/sajto_input_<fájlnév_stem>.txt`, ahol a
      `<fájlnév_stem>` a fájlnév `.docx` nélkül. A valódi nevek **szóközt és
      ékezetet tartalmaznak** (pl. `sajto_input_HR piaci körkép 2026-07-24.txt`)
      — bash-ben MINDIG idézőjelezd az útvonalat.

   b. Olvasd be ezt a fájlt (bash `cat`-tal a kiírt útvonalról, vagy a Windows
      útvonalon). Ellenőrizd, hogy a beolvasott hossz nagyságrendileg egyezik a
      kiírt karakterszámmal — ha rövidebb (mount-késés), olvasd újra bash-sel.
      Tájékoztatásul: a valódi körképek 12 000–32 000 karakter közöttiek.

   c. Dolgozd fel a „Feldolgozási szabályok" + „Stílus" szerint, és írd ki:
      `<mappa>/work/sajto_articles_<fájlnév_stem>.json`
      (docx-enként egyedi név — NE írj felül régi fájlt.)

   d. `python3 -B <mappa>/cowork/runner.py apply "<mappa>/work/sajto_articles_<fájlnév_stem>.json" "<fájlnév>"`

      A második argumentum a fájlnév **kiterjesztéssel, pontosan úgy, ahogy a
      `scan` kiírta** (pl. `HR piaci körkép 2026-07-24.docx`). NE toldj hozzá
      még egy `.docx`-et — a `processed.json` kulcsa ez a string lesz, és ha
      elrontod, a `scan` legközelebb újra újnak látja a fájlt.

      Az `apply` kiírja a kihagyott cikkeket (`FIGYELEM:` sorok) — add tovább.
      Ha JSON-hibát jelez többszöri próbálkozás után is, írd ki a JSON-t bash
      heredoc-kal (`cat > ... <<'EOF'`) és futtasd újra az `apply`-t.

5. Az összes új fájl után, **egyszer**:

   `python3 -B <mappa>/cowork/runner.py docx-relink "<b>"`

   (`<b>` = ugyanaz a bemeneti mappa, mint a `scan`-nél.) A mappa összes
   körképéből pótolja a link nélküli cikkek linkjét, ha a WS megadta — a régi
   körképekből is, mert 2026-10-14 előtt a mezőkódos linkek elvesztek. A docx
   linkje hiteles; a runner csak a bizonyíthatóan rosszat dobja el
   (`DOCX-LINK ELUTASÍTVA`), és az ütközést jelzi (`ÜTKÖZÉS`). Idempotens: ha
   nincs mit pótolni, nincs commit. Ezután jön a 4. lépés (link-pótlás
   kereséssel) — az már csak azt keresi, ami a docx-ben sem volt.

---

## Feldolgozási szabályok

A dokumentum szekciókra osztott. Főkategória-térkép:
- „Prohumanra vonatkozó megjelenések" → `"prohuman"`
- „Munkaerőpiaci hírek és elemzések" → `"piac"`
- „Versenytársakra vonatkozó megjelenések" → `"versenytars"`

Alkategória: a szekción belüli alcím (pl. „Trenkwalder", „Vendégmunka",
„Bérszínvonal / fizetések").

Minden cikknél: cím (`[URL: ...]` jelöléssel, ha volt hyperlink), forrás és
dátum zárójelben `(Médium, 2026.03.20.)`, majd összefoglaló szöveg.

Kimeneti JSON (CSAK az új cikkek, `het`/`id` mező NEM kell — a runner számolja):

```json
{"cikkek":[
  {"cim":"...","url":"https://... vagy null","forras":"...",
   "datum":"YYYY-MM-DD","fokategoria":"piac","alkategoria":"...",
   "osszefoglalo":"...","kulcsszavak":["..."]}
]}
```

Szabályok:
1. `cim` és `osszefoglalo`: CSAK sima szöveg, semmi markdown (`#`, `*`, `[`, `]`)
2. `url`: a `[URL: ...]` jelölésből; ha nincs, `null`. A cím szövegéből a
   `[URL: ...]` részt vedd ki — a `cim` mezőbe ne kerüljön bele.
   **Ha egy egész fájlban egyetlen `[URL: ...]` sincs, az hibára utal** (rossz
   feldolgozási út vagy sérült docx) — állj meg és jelezd, ne pusholj.
3. `datum`: KÖTELEZŐ, szigorúan `YYYY-MM-DD` (a magyar `2026.03.20.` formátumot
   konvertáld) — érvénytelen dátumú cikket a runner kihagy
4. `fokategoria`: KIZÁRÓLAG `prohuman` | `piac` | `versenytars` — mást a runner kihagy
5. Ha „Nem volt releváns megjelenés" áll egy alkategóriánál → hagyd ki
6. `kulcsszavak`: 2–4 címke **KIZÁRÓLAG a lenti zárt listáról** (lista, nem
   string). Ami nincs a listán, azt a runner eldobja — lásd „Címketár".
7. Ha van tartalomjegyzék-blokk („Tartalom" rész pontozott soraival), hagyd ki
   (az újabb dokumentumokban jellemzően nincs)
8. Ne találj ki adatot: ami nincs a szövegben, az `null` vagy üres
9. A WS-sablon mintasora („Cím (Médium, …)") nem cikk. Az `extract` 2026-10-14
   óta kihagyja; ha mégis látod, hagyd ki. Ha egy valódi cikknél maradt a
   „Médium" forrásnév, a runner az URL domainjéből pótolja.

---

## Címketár — a `kulcsszavak` mező ZÁRT listája

**Miért zárt:** amíg a címkék szabad szövegként keletkeztek, 435 cikkre 588
különböző címke jött létre, ebből 389 pontosan EGY cikkhez tartozott. Az oldal
címkefelhője így nem szűrő volt, hanem zaj — és a rutinból mindenre odaírt
`Prohuman` / `AI` címke ráadásul tárgyi tévedés volt a cikkek felében.

**A szabály:** csak az alábbi címkéket használhatod, szó szerinti írásmóddal.
Cikkenként **2–4** darab. A listán kívüli címkét a runner eldobja és
`FIGYELEM:` sorban jelzi. Ha egy cikkre egy sem illik, add meg a legközelebbi
témát — címke nélküli cikk hibás állapot.

**Amit szándékosan NE címkézz** (a szabadszavas kereső megtalálja őket, egy
cikkhez tartozó szűrő pedig használhatatlan):

- **szóvivők és személyek** — `Nógrádi József` helyett `Trenkwalder`,
- **városok, megyék** — `Szeged`, `Debrecen`, `Kecskemét`,
- **egyszeri cégnevek** — a beruházó, a gyár, a felmérést készítő tanácsadó,
- **`Prohuman`** — ezt a főkategória (`prohuman`) már pontosabban lefedi,
- **rendezvények** — `Business Summit`, `Ipari HR Konferencia`.

### Téma (36)

`AI` · `adózás` · `automatizáció` · `beruházás` · `bérek` ·
`bértranszparencia` · `demográfia` · `diákmunka` · `esélyegyenlőség` ·
`foglalkoztatottság` · `gazdasági kilátások` · `idősebb munkavállalók` ·
`képzés` · `kutatás és statisztika` · `leépítés` · `megtartás` ·
`mentális egészség` · `munkaerő-kölcsönzés` · `munkaerő-tartalék` ·
`munkaerőhiány` · `munkajog` · `munkakörülmények` · `munkanélküliség` ·
`munkáltatói márka` · `munkáltatói szövetség` · `nők a munkaerőpiacon` ·
`pályakezdők` · `politika` · `régiós különbségek` ·
`rugalmas foglalkoztatás` · `skills-alapú HR` · `szakszervezet` ·
`szezonális munka` · `toborzás` · `vendégmunka` · `vendégmunkás-stop`

Két gyakori félrenyúlás: a `szakszervezet` a **munkavállalói** oldal (MASZSZ,
KASZ, sztrájk); a VOSZ, MAGOSZ, MMOSZ a `munkáltatói szövetség`. A `képzés`
minden oktatást és átképzést visz, a `skills-alapú HR` viszont csak a
kompetencia-alapú kiválasztást/karriert.

### Szektor (9)

`IT` · `agrárium` · `autóipar` · `egészségügy` · `építőipar` · `ipar` ·
`kereskedelem` · `logisztika` · `turizmus`

### Szereplő (20)

`Adecco` · `Gi Group` · `Giggle` · `HR-Rent` · `Job Force` · `Jobtain` ·
`KSH` · `Man at Work` · `Manpower` · `Meló-Diák` · `Menton Jobs` ·
`Mind-Diák` · `Pannon-Work` · `Pensum Group` · `Prodiák` · `Profession.hu` ·
`Randstad` · `Trenkwalder` · `WHC` · `Work Force`

**Ha egy új szereplő visszatérővé válik** (két-három körképben is előjön), ne
kerülgesd: írd bele a záró összefoglalóba, hogy fel kell venni a
`cowork/vocab.py`-ba. A listát ott kell bővíteni, nem itt improvizálni.

---

## Stílus — az `osszefoglalo` mezőhöz (no-ai-slop, magyar adaptáció)

Ez a blokk a `no-ai-slop` skill magyar adaptációja. **Közvetlenül alkalmazd,
amikor az `osszefoglalo`-t írod.** Amit itt NE tegyél:

- NE hívd meg külön körben a `no-ai-slop` skillt (felügyelet nélküli futásban
  visszakérdezhet és elakadhat),
- NE írj „What changed" / „Mi változott" szekciót,
- NE futtass utólagos átíró kört a már megírt összefoglalókon.

### Alapelvek

1. **Kezdd a lényeggel.** Mi történt, ki mondta, mekkora a szám. Semmi felvezetés.
2. **A konkrétum szent.** Szám, dátum, név, arány, cégnév mindig maradjon meg. Soha
   ne cseréld le általánosításra: „a bérek 12 százalékkal emelkedtek", nem „a bérek
   jelentősen emelkedtek".
3. **Aktív ige.** „A cég 200 főt vett fel", nem „200 fő felvételére került sor".
4. **Ne értékelj.** Ez sajtófigyelés, nem vélemény. Ne vonj le következtetést, ne
   minősíts, ne tegyél hozzá jelentőséget, ami nincs a forrásban. A no-ai-slop
   „legyen véleményesebb" elve **itt nem érvényes** — a pontosság a termék.
5. **Hossz: 2–4 mondat.** Se bevezető, se lezáró keret.
6. **Ugyanaz a szó ismételhető.** Ne cserélgesd stílusból: ha „a cég", akkor
   végig „a cég", nem „a vállalat" / „a társaság" váltakozva.

### Kerülendő szavak

Ne használd: `kulcsfontosságú`, `kiemelt fontosságú`, `paradigmaváltás`,
`mérföldkő`, `áttörés` (kivéve ha a forrás szó szerint így nevezi),
`innovatív`, `dinamikusan fejlődő`, `kihívásokkal teli`, `átfogó`,
`transzformatív`, `élen jár`, `elősegíti`, `hozzájárul`, `támogatja a
folyamatokat`, `optimalizálja`, `felgyorsítja`, `robusztus`, `komplex kihívás`.

Gyakran üres fordulatok — töröld, ha csak késleltetik a lényeget:
`érdemes megjegyezni, hogy`, `fontos megjegyezni, hogy`, `a mai világban`,
`napjainkban`, `a jelenlegi helyzetben`, `ami azt illeti`, `végső soron`,
`összességében`, `mindezek fényében`, `nem véletlen, hogy`, `egyre inkább`.

### Kerülendő minták

- **Felszínes elemzés-farok.** Töröld a „…, rávilágítva arra, hogy",
  „…, jelezve, hogy", „…, ami jól mutatja, hogy", „…, alátámasztva azt, hogy"
  típusú lezárásokat. Vagy konkrét következményt írj (ha a forrásban benne van),
  vagy fejezd be a mondatot.
- **Fontosság-puffasztás.** „megerősíti pozícióját", „jelentős előrelépés",
  „kiemelt szerepet játszik", „mérföldkövet jelent". Írd le a tényt, és hagyd,
  hogy az olvasó eldöntse.
- **Bináris szembeállítás.** „Nem X, hanem Y", „A kérdés nem az, hogy X, hanem
  hogy Y". Mondd ki egyszerűen Y-t.
- **Kettőspontos leleplezés.** „A lényeg: a bérek nem emelkedtek." → sima mondat.
- **Homályos hivatkozás.** „a szakértők szerint", „az elemzők úgy vélik", „a
  kutatások azt mutatják". Ha a forrás megnevezi a személyt vagy szervezetet,
  **nevezd meg**. Ha a forrás maga is névtelenül fogalmaz, hagyd meg úgy, ahogy a
  forrásban van — de **soha ne találj ki forrást**.
- **Gondolatjel-halmozás.** Rövid összefoglalóban ne használj gondolatjelet;
  vessző, pont vagy zárójel jobb.
- **Drámai tagolás.** „Ennyi. Ez az egész." vagy „X. És Y. És Z." — teljes
  mondatokat írj.
- **Összegző lezárás.** Ne zárd „Összességében…" / „Végső soron…" mondattal. Az
  utolsó konkrét tényen fejezd be.

---

## 4. Link-pótlás (`relink-list` / `relink-apply`)

A docx-ekben a linkek Word-hyperlinkek, és néha hiányoznak. Kezdőállapot
2026-08-05: 449 cikkből 62 URL nélküli (13,8%). Az első nagy kör után **10**
maradt. 2026-10-14: 611-ből 22. A 07-24 óta link nélkül érkezett 40 cikkből
30-at a link-pótlás később megtalált — a linkek többsége valójában benne volt a
docx-ben mezőkódként (az `extract` ezt azóta olvassa). Az `apply` csak hozzáfűz, meglévő cikket nem módosít — ezért külön út.

**Miért itt fut, és nem külön időzített taskban:** a cikkek idővel *nehezebben*
találhatók meg (lekerülnek a főlapról, fizetőfal, archívum), tehát a legjobb
pillanat a keresésre közvetlenül a friss feldolgozás után.

### 4/A — Friss kör (minden futásban, az `apply` UTÁN)

```
python3 -B <mappa>/cowork/runner.py relink-list --het=<a most feldolgozott hét hétfője>
```

### 4/B — Backlog-kör (csak hétfőn, a friss kör után)

```
python3 -B <mappa>/cowork/runner.py relink-list --limit=10
```

A `--limit=10`-et ne emeld meg. A `relink-list` a **legkevesebbszer keresetteket**,
azon belül a legfrissebbeket adja előre (így a hétfői kör végigforog a listán), és
`stderr`-en kiírja, mennyi maradt — ezt a számot **mindig vedd bele a záró
összefoglalóba**, hogy látszódjon a fogyás.

A kimenetnek két listája van:

- `cikkek` — ezeket keresd meg (4/C).
- `kezi_lista` — ezeket **NE keresd**. Vagy a domainjüket nem éri el a kereső
  (`NEM_KERESHETO`: nlc.hu, 24.hu, rtl.hu — mérve 2026-10-14), vagy háromszor
  sem volt találat. Ezeket a zárásban sorold fel (id, cím, forrás, dátum, ok):
  Gellért kézzel pár perc alatt megtalálja őket (2026-10-14-én két ilyet
  talált meg perceken belül).

### 4/C — Keresés cikkenként

A `relink-list` minden cikkhez ad `domain_tipp`-et.

1. **Egzakt cím, a médium domainjére szűkítve.** `WebSearch`, `allowed_domains:
   ["<domain_tipp>"]`, a query a cikk címe **idézőjelben**.
2. **Ha nincs egzakt találat:** MÉG EGY próba a cím legjellegzetesebb 6-10 szavas
   részletével (a kettőspont előtti figyelemfelhívó részt hagyd el).
3. **Ha ezután sincs:** ne keress harmadszor, de **add be**
   `{"id":"ART-…","nincs_talalat":true}` formában. Ez naplózza a sikertelen
   keresést (`url_keresve`); 2026-10-14 előtt ez nyomtalan volt, ezért nem
   látszott, keresett-e egyáltalán a kör. Három sikertelen kör után a tétel a
   `kezi_lista`-ra kerül.
4. **Ha a domain-szűrt keresés hibát ad** (pl. HTTP 400), az NEM „nincs
   találat": ne naplózd, hanem írd a zárásba a domaint — a runner
   `NEM_KERESHETO` halmazába kell felvenni.

**A LEGFONTOSABB SZABÁLY:** a találati lista első helye NEM bizonyíték. A
keresők a magyar egzakt kifejezést gyakran figyelmen kívül hagyják, és
*tematikusan* hasonló, de teljesen más cikket adnak vissza az 1. helyen. Két
mért példa (2026-08-05):

- G7, „A vendégmunkásvita a magyar ipar egyik érzékeny pontját találta el"
  (2026-07-14) → az 1. találat egy **2025. januári** vendégmunkás-cikk volt.
- GyártásTrend, „70 új csomagológyári munkahely Makón" → az 1. találat a
  „Kapacitásbővítés 70,8 milliárd forintból **Dunaföldváron**" (a 70 ott a
  milliárd forint összegre utal, nem munkahelyszámra).

Csak akkor add be egy linket, ha a **találat címe** a cikk címe. Nem „hasonló",
nem „ugyanarról szól" — ugyanaz a cím.

**Médiumspecifikus tudás (mérve):**

- **G7** = ma a Telex aloldala: `telex.hu/g7/<kat>/2026/07/14/<slug>`. Régi
  cikkek g7.hu-n. A guard mindkettőt fogadja (`MEDIA_ALIASES`), és a `forras`
  marad „G7" — NE írd át Telexre.
- **rtl.hu, nlc.hu, 24.hu:** a domain-szűkített `WebSearch` HTTP 400-at ad
  (proxy), az rtl.hu lekérése 403, az nlc.hu lekérése tiltott. Ezeket a runner
  a `kezi_lista`-ra teszi — ne keresd és ne zárd le őket.
- **HRPOWER** = `hrpwr.hu` (a WS „HRPOWER"-ként írja). A `domain_tipp`
  2026-10-14 előtt üres volt erre a forrásra, tehát a keresés nem volt a
  domainre szűkítve.
- **Világgazdaság:** ugyanannak a cikknek háromféle címe lehet (headline, az
  oldal `<title>`-je, és a WS-ben szereplő cím). Itt a 6–10 szavas
  részlet-keresés a fontos, és az `oldal_cim` kötelező (a URL-ben csak év/hónap
  van).
- **autopro.hu** 520-at ad, **ado.hu** robots.txt-tiltást: ezeknél `oldal_cim`
  nem beszerezhető, a slug-egyezés viszont működik (lásd 4/E).
- **Nyomtatott megjelenés** (Budapest Business Journal, Dél-Dunántúli Gazdaság,
  Műszaki Magazin): a docx `[URL:]`-je itt `e.pcloud.link` PDF-szkennelés, nem
  cikk-link. Ezeknél nincs mit keresni.

### 4/D — TV / rádió (`broadcast: true`)

ATV, Híradó.hu, Trend FM, Kontroll: gyakran nincs online cikk, de **nem
automatikusan** — egy Kontroll-tételnél youtube.com a link.

**NE zárd le őket egy egzakt címkeresés után.** TV-tételnél a WS címe gyakran
csak címke („Interjú Nógrádi Józseffel…"), az online változat más címmel jelenik
meg, tehát az egzakt címkeresés eleve nem találhatja meg. 2026-08-05-én így
zárult le tévesen ART-0172: az atv.hu-n 04-25-én fent volt az ATV Start-interjú
„Ez lehet a Tisza-kormány első nagy dobása…" címmel.

Helyette: keress a **szereplő nevére + a téma 2–3 kulcsszavára** a médium
domainjén, a dátum ±2 napján belül. A találatot kérd le, és csak akkor add be,
ha a szereplő és a téma egyezik — `oldal_cim`-mel és `indok`-kal (pl. „Nógrádi
József az ATV Startban a minimálbér-szja-ról"). Ha nincs: `nincs_talalat`.

`"url_forras":"nincs"` lezárást **csak Gellért kérésére** adj be, `indok`
mezővel — a runner indok nélkül elutasítja.

### 4/E — A JSON és a visszaírás

```json
{"linkek":[
  {"id":"ART-0433","url":"https://www.portfolio.hu/gazdasag/20260722/..."},
  {"id":"ART-0339","url":"https://profitline.hu/van-elet-a-vendegmunka-utan-487869",
   "oldal_cim":"Van élet a vendégmunka után? - ProfitLine.hu"},
  {"id":"ART-0596","nincs_talalat":true},
  {"id":"ART-0269","url_forras":"nincs","indok":"Gellért: nyomtatott HVG, nincs online változat"},
  {"id":"ART-0124","duplikatuma":"ART-0001"}
]}
```

Írd a `<mappa>/work/sajto_relink_<dátum>.json` fájlba, majd:

```
python3 -B <mappa>/cowork/runner.py relink-apply "<mappa>/work/sajto_relink_<dátum>.json"
```

A runner **három bizonyítékfajtát** fogad el, legalább egy kell:

1. **napra pontos URL-dátum** a cikk dátumától ±3 napon belül
   (portfolio.hu `/20260722/`, telex.hu `/2026/05/15/`, hrportal.hu
   `-20260727.html`) — offline ellenőrizhető,
2. **slug-egyezés**: a URL utolsó beszédes szegmense a cím slugosított alakja
   (`profitline.hu/van-elet-a-vendegmunka-utan-487869`) — szintén offline, és
   olyan médiumoknál az egyetlen elérhető, amelyek nem engedik a lekérést,
3. **`oldal_cim`**: a WebFetch-csel lekért oldal valódi `<title>`-je.

Ha csak év-hónap van a URL-ben (vg.hu `/2026/05/`), az önmagában NEM elég —
kérd le az oldalt, és add meg az `oldal_cim`-et. Ha a URL dátuma vagy éve-hónapja
ELLENTMOND a cikk dátumának, a runner akkor is elutasít, ha a cím egyezik.

Az `oldal_cim`-be **soha ne másold vissza a cikk címét** — az a bizonyíték, nem
formalitás. Ha nem kérted le az oldalt, hagyd ki a mezőt.

### 4/F — Duplikátum

Ha két cikk ugyanaz (azonos dátum és forrás, az egyik címe csonka — a `_keys()`
dedupe ezt elszalasztja), add be `{"id":"<csonka>","duplikatuma":"<a jó>"}`
formában. A runner ellenőrzi a dátumot, a forrást és a cím-hasonlóságot, és csak
akkor törli, ha a párnak van URL-je.

### 4/G — Amit a runner elutasít (ez nem hiba, ez a védelem)

Az `ELUTASÍTVA` sorokat **szó szerint add tovább** a záró összefoglalóban. A
`FIGYELEM: <domain> nincs a DOMAIN_MAP-ban` sort is: azt a runner.py
`DOMAIN_MAP`-jába fel kell venni.

- `ÜTKÖZÉS`: a talált URL már egy másik cikknél van. Ez szinte mindig azt
  jelenti, hogy a WS rossz tételhez tette a linket (2026-08-14: ART-0474 kapta a
  0479 cikkének URL-jét). Ne próbáld újra, nem számít próbálkozásnak — add
  tovább szó szerint, Gellért dönt (`kezi-javitas`).
- `KIHAGYVA`: a cikknek közben lett linkje — nem hiba, nem számít
  próbálkozásnak.

### 4/H — Forrásnevek

A `relink-apply --fix-sources` (vagy külön `fix-sources`) egységesíti az
írásmódokat a `CANON_SOURCE` szerint (HR Portál / Hrportál / HR Portal → „HR
Portál"; Economx / economx.hu → „Economx"; G7 / Telex, G7–Telex → „G7"). Ha egy
csoportot nem ismer, a leggyakoribb írásmód nyer és `FIGYELEM` sorban jelzi —
akkor vedd fel a `CANON_SOURCE`-ba.

---

## Zárás

A végén foglald össze:
- hány fájlt dolgoztál fel, hány új cikk, push megtörtént-e (`PUSH OK`),
- **link-pótlás:** hány link került be, hány cikk lett lezárva (`nincs`), hány
  lett elutasítva és miért, hány sikertelen keresést naplóztál, és mennyi
  maradt a backlogban (a `relink-list` `stderr`-sora),
- **kézi pótlásra vár:** a `relink-list` `kezi_lista`-jának tételei
  soronként: id, cím, forrás, dátum, ok. Mindig sorold fel, akkor is, ha
  nem változott — ez Gellért teendőlistája,
- **WS-kivonat:** ha az `apply` `WS-kivonat: N / M` sorában N < M, add tovább,
- **docx-relink:** hány link jött a körképekből, és minden `DOCX-LINK
  ELUTASÍTVA` / `ÜTKÖZÉS` sor szó szerint,
- van-e olyan körkép, ami a OneDrive `Sajtofigyelo_Incoming` mappájában ott van,
  de a csatolt bemeneti mappában nincs meg (2026-08-05-i mérés: a flow működik,
  a mappában 18 fájl van 04-02-től 07-31-ig — tehát ha itt eltérés van, az a
  bejuttatás hibája, nem a flow-é),
- **link-audit:** minden `FIGYELEM: forras átírva ...` sor (ott a docx rossz
  médiumnevet adott, a runner az URL domainje alapján javította), és minden
  `FIGYELEM: gyanús link ...` sor **szó szerint** — ezek kézi ellenőrzést
  kérnek, a runner szándékosan nem nyúl hozzájuk,
- kihagyott cikkek, `FIGYELEM:` sorok, hibák. Ha `HIBA: N cikk kiesett a
  validáción` volt, azt a legelső sorban mondd ki, a fájlnévvel együtt.

Ellenőrzésképp kérd le a push után 2-3 perccel **mindkét** URL-t:
- `https://t0thgellert.github.io/sajto/`
- `https://sajto.prohuman.group/` (a CNAME 2026-07-20-án beállt; a véglegesítés
  IT-döntés alatt van — ha még nem válaszol, azt jelezd, de ne kezeld a futás
  hibájának)

Frissült-e a `generated` dátum? (A GitHub Pages build 1-2 percet késhet — ha még
régi, jelezd, de ne kezeld hibaként.)

## Fontos technikai megkötések

- A repót a runner egyedi temp mappába klónozza — a projektmappában SOHA ne
  futtass gitet, és fix /tmp-útvonalakat (pl. `/tmp/sajto_work`) SOHA ne
  használj: a /tmp más sessionök 'nobody' tulajdonú fájljait tartalmazhatja.
- Mindig `python3 -B`-vel és teljes útvonallal futtasd a runnert. Létező
  alparancsok: `scan`, `extract`, `apply`, `relink-list`, `relink-apply`,
  `fix-sources`, `audit-links`, `sync-cowork`, `docx-relink`, `kezi-javitas`,
  `ws-backfill` (az utolsó kettő nem a heti futás része). Más nincs — ha valami hiányzik, a runner
  kiírja a saját docstringjét és 1-es kóddal kilép, tehát a `python3 -B
  <mappa>/cowork/runner.py` (argumentum nélkül) mindig megmutatja az
  érvényes listát. NE ebben a fájlban keresd az igazságot: a runner
  docstringje az.
- A PAT a projektmappa `github_pat.txt` fájljában van; ha hiányzik, állj le és
  jelezd. (Publikus repónál a klón rossz PAT-tal is sikerül — a PAT-hiba csak
  push-nál derül ki.)
- Az `apply` háromféleképpen zárulhat, és **ezek NEM ugyanazok**:
  1. `KESZ` + `PUSH OK` → rendben, új cikkek bekerültek.
  2. `KESZ`, de „Nincs új cikk" és a `[... validáción kiesett 0]` → duplikált
     küldés, **nem hiba**, a fájl feldolgozottnak számít.
  3. `HIBA: N cikk kiesett a validáción` → **ez hiba.** A fájl szándékosan NEM
     lett feldolgozottként megjelölve. Olvasd el a `FIGYELEM:` sorokat, javítsd
     a JSON-ban a jelzett cikkeket (jellemzően `datum` formátum vagy
     `fokategoria`), és futtasd újra ugyanezt az `apply`-t. A már bekerült
     cikkek nem duplikálódnak. Ha kétszer sem megy át, a záró összefoglalóban
     nevezd meg a fájlt és a hibás cikkeket — így nem veszik el csendben.
- A `state/processed.json` mondja meg, mi volt már feldolgozva. Ez a fájlszintű
  dedupe. Cikkszinten a runner `merge` függvénye kisbetűsített `cim`+`datum`
  vagy URL egyezésre dedupál — ez másodlagos védelem, ne támaszkodj rá.
- Ha a `work/` vagy a `state/` mappa nem létezik, a runner létrehozza az
  `extract`/`apply` során. Ha te írsz oda kézzel, előbb `mkdir -p`.

## A bejuttatás — hogyan kerül a docx egy csatolt mappába

A feldolgozásnak egyetlen előfeltétele van: a `.docx` legyen ott egy csatolt
mappában. Erre két út működik, és a pipeline szempontjából felváltva
egyenértékűek — a `scan` mindkettőt ugyanúgy kezeli.

1. **Szinkronizált OneDrive-mappa csatolva.** A Power Automate flow
   (`87b404ad-b66e-4df6-8df0-1f5dcc208778`, On, 2026-08-04-én ellenőrizve: 5
   `Succeeded` futás) a `wshungary.hu`-s mellékleteket a OneDrive
   `Sajtofigyelo_Incoming` mappájába menti. Ha ez a mappa lokálisan
   szinkronizálva van (és „Megtartás mindig ezen az eszközön"), akkor
   csatolható, és a `scan "<b>"` közvetlenül onnan olvas.

2. **Lokális szkript az Outlookból.** A `cowork/save_attachments.ps1` az Outlook
   desktopból (COM) az `incoming/` mappába mentí a `.docx`-eket. Nem kell hozzá
   OneDrive-szinkron, se licenc. Regisztrálás egyszer, Windows Task
   Schedulerbe — a szkript fejében van a parancs.

Ha egyik sincs beállítva, a fájlt kézzel kell az `incoming/`-ba tenni. A
`scan` ekkor is működik, csak nem magától.
