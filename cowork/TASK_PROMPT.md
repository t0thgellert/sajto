# Sajtófigyelő – heti feldolgozás (scheduled task instrukció)

Ez a referencia-instrukció a hétfői és pénteki scheduled taskhoz.
A projektmappa Windows-útvonala: `C:\Users\toth.gellert\Claude\Projects\Sajtófigyelés`
A bash sandboxban: `/sessions/<session>/mnt/Sajtófigyelés/` (a pontos mount-útvonal a session-leírásban).
Lent `<mappa>` = a sandbox-útvonal. A runnert MINDIG teljes útvonallal hívd.

**Verzió: 2026-07-28c.** A `b`-ben került be a OneDrive-észlelés (Power Automate
flow) és a magyar stílusszabály-blokk. A `c`-ben a zárt **Címketár** — a
`kulcsszavak` mező innentől nem szabad szöveg, és a runner validálja.

---

## ⚠ A LEGFONTOSABB MEGKÖTÉS — olvasd el, mielőtt bármit teszel

A dokumentumokban a cikkek linkjei **Word-hyperlinkek**, nem kiírt URL-ek. A
`[URL: ...]` jelölést, amire a feldolgozási szabály hivatkozik, **a runner.py
maga gyártja** a docx hyperlink-relationshipjeiből (`runner.py` `extract_docx`,
88–98. sor). Ez kizárólag a lokális, python-docx alapú `extract` során jön létre.

Ebből következik: **a Microsoft 365 connector (`read_resource`) által
visszaadott szöveg NEM használható cikk-feldolgozásra.** A connector csak a
látható szöveget adja vissza, hyperlink-célok nélkül — abból minden cikk
`url: null` lenne, a `state/processed.json` pedig véglegesen lezárná a fájlt,
tehát a linkek soha nem kerülnének be. Összehasonlításul: a 2026-07-24-i futás
17 cikkéből 14-nek volt URL-je.

**TILOS tehát:** `runner.py apply`-t hívni olyan JSON-ra, amit a connectorból
kapott szövegből állítottál elő. A `read_resource` a OneDrive-mappára
kizárólag **észlelésre** használható (mi vár feldolgozásra), nem tartalom-
kinyerésre.

---

## Lépések

1. `pip install python-docx --break-system-packages` (ha még nincs telepítve).
2. **Észlelés** — mi vár feldolgozásra (2/A: OneDrive, 2/B: lokális).
3. **Feldolgozás** — kizárólag lokálisan meglévő fájlokra (3.).
4. **Zárás** és ellenőrzés.

---

## 2/A — OneDrive-észlelés

A Power Automate flow (`87b404ad-b66e-4df6-8df0-1f5dcc208778`, státusz On) minden
`wshungary.hu` feladótól érkező, mellékletes emailt lement Gellért OneDrive for
Business-ébe a `Sajtofigyelo_Incoming` mappába. Ezt a mappát a Microsoft 365
connectorral olvasod — **de csak azért, hogy megtudd, MELYIK fájl érkezett.**

**driveId (Gellért „Dokumentumok" OneDrive-ja):**
`b!LVNR9gwBukiluIqC2nb9BVwsKM1JEWZAtuAT7MFQR8rEDdopWv8oRq25yJjc4raz`

Ha ez hibát ad, kérd le újra: `read_resource` → `drive:///users/me`, és a
**„Dokumentumok"** nevű drive URI-jából vedd a driveId-t. A
„PersonalCacheLibrary" NEM az — a két driveId az első 44 karakteren azonos,
tehát figyelmesen hasonlítsd össze.

Lépések:

1. `read_resource` → `file:///{driveId}/` — a OneDrive gyökere.
   Keresd a `Sajtofigyelo_Incoming` nevű **folder** bejegyzést.
   - **Figyelem:** a gyökérlistában az elemek elválasztó nélkül futhatnak
     egymásba (`...SQ4AE3AX57Attachments (folder, ...)`). Az itemId a
     `file:///{driveId}/` utáni rész **a következő elem neve előtt** — a
     megnevezés (`Sajtofigyelo_Incoming (folder, ...)`) és az URI párosítására
     figyelj, ne vakon az utolsó szegmenst vedd.
   - **Ha nincs ilyen mappa:** a flow még soha nem futott le élesben
     (2026-07-28-i állapot: valóban nem létezett). Jegyezd fel a záró
     összefoglalóba, és menj a 2/B-re.

2. `read_resource` → `file:///{driveId}/{mappa_itemId}` — a mappa tartalma
   (soronkénti lista). Csak a `.docx` fájlok érdekelnek; a `.doc`/`.xlsx`
   fájlokat nevezd meg `FIGYELEM:` sorként.

3. **Dedupe.** Olvasd be a `<mappa>/state/processed.json` `"processed"` listáját.
   Csak azok a fájlok újak, amelyek **neve** (kiterjesztéssel) nincs a listában.
   A `processed.json`-t kézzel SOHA ne írd — azt a runner `apply` írja.

4. Minden így talált új fájlnévre nézd meg, megvan-e **lokálisan** is:
   `ls -la "<mappa>/incoming/"`
   - **Ha lokálisan is megvan:** semmi külön dolgod, a 2/B `scan` fel fogja
     venni. Menj a 2/B-re.
   - **Ha lokálisan NINCS meg:** ezt a fájlt **nem tudod feldolgozni** (lásd a
     fenti megkötést: a connector szövegéből elvesznének a linkek). Írd fel a
     záró összefoglalóba pontos fájlnévvel:
     „`<fájlnév>` a OneDrive `Sajtofigyelo_Incoming` mappájában vár, de
     lokálisan nincs meg. A linkek csak a lokális python-docx feldolgozásból
     jönnek ki, ezért kézi mentés (vagy a `mklink /J` junction beállítása) kell
     az `incoming/` mappába."
   - **NE hívj `apply`-t erre a fájlra.** Inkább maradjon feldolgozatlan, mint
     hogy linkek nélkül, véglegesen bekerüljön.

5. Az `extract` alparancsot a OneDrive-os itemId-kre soha ne próbáld meghívni —
   az a lokális `incoming/`-ra (vagy abszolút lokális útvonalra) van kötve.

---

## 2/B — Lokális `incoming/` (a tényleges feldolgozási út)

1. `python3 -B <mappa>/cowork/runner.py scan`

2. Ha `NINCS_UJ`:
   - Ha a 2/A talált OneDrive-on várakozó, lokálisan nem meglévő fájlt, akkor
     már tudod a választ: azt jelentsd (4/a lépés a 2/A-ban). Fejezd be.
   - Egyébként ellenőrizd az M365 connectorral (`outlook_email_search`),
     érkezett-e az elmúlt 7 napban **„piaci körkép"** tárgyszavú email
     (mindhárom írásmód előfordult: „piaci körkép", „HR Piaci Körkép",
     „HR-piaci körkép" — a substring-keresés mindet elkapja). **Feladó
     email-címre NE szűrj** — a WS Hungary-s küldő személye változik (volt Laky
     Zoltán és Szökendi Zsuzsanna is), és Major Gábor forwardja is bejön.
   - **HAMIS RIASZTÁS ELKERÜLÉSE (fontos):** a 7 napos ablakban benne lehet egy
     már FELDOLGOZOTT email is (jellemzően a hétfői futásnál az előző pénteki).
     Mielőtt hiányzó mellékletet jelentesz, hasonlítsd össze az email
     mellékletnevét a `processed.json` listájával. Ha a fájl már fel van
     dolgozva, azt írd: „nincs teendő, az utolsó körkép (`<fájlnév>`) már
     feldolgozva". Csak akkor jelezz hiányt, ha a melléklet neve NINCS a
     `processed.json`-ban és lokálisan sem létezik.
   - Fejezd be a futást.

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

## Zárás

A végén foglald össze:
- hány fájlt dolgoztál fel, hány új cikk, push megtörtént-e (`PUSH OK`),
- **a OneDrive-észlelés eredménye:** létezik-e a `Sajtofigyelo_Incoming` mappa
  (ha igen: **a Power Automate flow működik** — ez az első éles visszajelzés
  róla), és van-e ott olyan fájl, ami lokálisan még nincs meg,
- kihagyott cikkek, `FIGYELEM:` sorok, hibák.

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
  alparancsok: `scan`, `extract`, `apply` — más nincs.
- A PAT a projektmappa `github_pat.txt` fájljában van; ha hiányzik, állj le és
  jelezd. (Publikus repónál a klón rossz PAT-tal is sikerül — a PAT-hiba csak
  push-nál derül ki.)
- Ha az `apply` „Nincs új cikk" eredménnyel zárul, az nem hiba (duplikált küldés).
- A `state/processed.json` mondja meg, mi volt már feldolgozva. Ez a fájlszintű
  dedupe. Cikkszinten a runner `merge` függvénye kisbetűsített `cim`+`datum`
  vagy URL egyezésre dedupál — ez másodlagos védelem, ne támaszkodj rá.
- Ha a `work/` vagy a `state/` mappa nem létezik, a runner létrehozza az
  `extract`/`apply` során. Ha te írsz oda kézzel, előbb `mkdir -p`.

## Nyitott ügy — a teljes automatizálás hiányzó fele

A Power Automate flow lementi a mellékletet a OneDrive-ra, de a fájl **bináris
tartalma** onnan nem hozható be a feldolgozásba: a connector csak szöveget ad
(linkek nélkül), a bash sandbox pedig csak a csatolt mappákat látja.

A hiányzó láncszem egy egyszeri Windows-junction, amit Gellértnek kell
lefuttatnia rendszergazdai `cmd`-ben, miután a `Sajtofigyelo_Incoming` mappa
létrejött és lokálisan szinkronizálódott:

```
mklink /J "C:\Users\toth.gellert\Claude\Projects\Sajtófigyelés\incoming\_onedrive" "C:\Users\toth.gellert\<OneDrive-szinkron-mappa>\Sajtofigyelo_Incoming"
```

Ezután a synced docx-ek a projektmappán belül, a bash sandbox számára is
láthatóan megjelennek, és a `runner.py` a saját python-docx útján dolgozza fel
őket — hyperlinkekkel együtt. **A `runner.py` `cmd_scan` viszont jelenleg csak az
`incoming/` mappa közvetlen tartalmát olvassa (nem rekurzívan), tehát a junction
bekötése után a `scan`-t is módosítani kell** (`runner.py` 338–341. sor,
`INCOMING.iterdir()` → `INCOMING.rglob("*")`). Ez még nincs megtéve.
