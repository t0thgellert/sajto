# Sajtófigyelő – heti feldolgozás (scheduled task instrukció)

Ez a referencia-instrukció a hétfő/péntek 16:00-s scheduled taskhoz.
A projektmappa Windows-útvonala: `C:\Users\toth.gellert\Claude\Projects\Sajtófigyelés`
A bash sandboxban: `/sessions/<session>/mnt/Sajtófigyelés/` (a pontos mount-útvonal a session-leírásban).
Lent `<mappa>` = a sandbox-útvonal. A runnert MINDIG teljes útvonallal hívd.

## Lépések

1. `pip install python-docx --break-system-packages` (ha még nincs telepítve).
2. Futtasd: `python3 -B <mappa>/cowork/runner.py scan`
   - Ha `NINCS_UJ`: ellenőrizd az M365 connectorral (outlook_email_search),
     érkezett-e új "piaci körkép" tárgyú email az elmúlt 7 napban. Ha igen,
     jelezd a felhasználónak, hogy a docx még nem került be az incoming/
     mappába. Fejezd be a futást.
   - A `FIGYELEM: ... .doc formátum` sorokat add tovább a felhasználónak
     (kézi .docx-mentés kell).
3. Minden új fájlra (a scan sorrendjében):
   a. `python3 -B <mappa>/cowork/runner.py extract "<fájlnév>"`
      → kiírja: `OK <útvonal> (N karakter)` — ez a work/ mappába kerül,
      docx-enként egyedi néven (pl. work/sajto_input_korkep0703.txt).
   b. Olvasd be ezt a fájlt (bash `cat`-tal a kiírt útvonalról, vagy a Windows
      útvonalon: `...\Sajtófigyelés\work\...`), és dolgozd fel az alábbi
      szabályokkal. Ellenőrizd, hogy a beolvasott hossz nagyságrendileg egyezik
      a kiírt karakterszámmal — ha rövidebb (mount-késés), olvasd bash-sel.
   c. Az eredményt írd ide: `work/sajto_articles_<fájlnév_kiterjesztés_nélkül>.json`
      (docx-enként egyedi név — NE írj felül régi fájlt).
   d. `python3 -B <mappa>/cowork/runner.py apply "<mappa>/work/<json-fájl>" "<fájlnév>"`
      - Az apply kiírja a kihagyott cikkeket (`FIGYELEM:` sorok) — ezeket
        add tovább a felhasználónak.
      - Ha JSON-hibát jelez többszöri próbálkozás után is, írd ki a JSON-t
        bash heredoc-kal (`cat > ... <<'EOF'`) és futtasd újra az apply-t.
4. A végén foglald össze: hány fájl, hány új cikk, push megtörtént-e
   (`PUSH OK` a runner kimenetében), kihagyott cikkek, hibák.
   Ellenőrzésképp kérd le a https://t0thgellert.github.io/sajto/ oldalt
   2-3 perccel a push után: frissült-e a `generated` dátum. (A GitHub Pages
   build 1-2 percet késhet — ha még régi, jelezd, de ne kezeld hibaként.)

## Feldolgozási szabályok (b lépés)

A dokumentum szekciókra osztott. Főkategória-térkép:
- "Prohumanra vonatkozó megjelenések" → `"prohuman"`
- "Munkaerőpiaci hírek és elemzések" → `"piac"`
- "Versenytársakra vonatkozó megjelenések" → `"versenytars"`

Alkategória: a szekción belüli alcím (pl. "Trenkwalder", "Vendégmunka",
"Bérszínvonal / fizetések").

Minden cikknél: cím (néha `[URL: ...]` jelöléssel), forrás és dátum zárójelben
`(Médium, 2026.03.20.)`, majd összefoglaló szöveg.

Kimeneti JSON (CSAK az új cikkek, het/id mező NEM kell — a runner számolja):

```json
{"cikkek":[
  {"cim":"...","url":"https://... vagy null","forras":"...",
   "datum":"YYYY-MM-DD","fokategoria":"piac","alkategoria":"...",
   "osszefoglalo":"...","kulcsszavak":["..."]}
]}
```

Szabályok:
1. cim és osszefoglalo: CSAK sima szöveg, semmi markdown (#, *, [, ])
2. url: a `[URL: ...]` jelölésből; ha nincs, null
3. datum: KÖTELEZŐ, szigorúan YYYY-MM-DD (a magyar `2026.03.20.` formátumot
   konvertáld) — érvénytelen dátumú cikket a runner kihagy
4. fokategoria: KIZÁRÓLAG prohuman | piac | versenytars — mást a runner kihagy
5. Ha "Nem volt releváns megjelenés" áll egy alkategóriánál → hagyd ki
6. kulcsszavak: 2–5 releváns magyar kulcsszó (lista, nem string)
7. A tartalomjegyzék-blokkot (a "Tartalom" rész pontozott soraival) hagyd ki
8. Ne találj ki adatot: ami nincs a szövegben, az null vagy üres

## Fontos technikai megkötések

- A repót a runner egyedi temp mappába klónozza — a projektmappában SOHA ne
  futtass gitet, és fix /tmp-útvonalakat (pl. /tmp/sajto_work) SOHA ne használj:
  a /tmp más sessionök 'nobody' tulajdonú fájljait tartalmazhatja.
- Mindig `python3 -B`-vel és teljes útvonallal futtasd a runnert.
- A PAT a projektmappa `github_pat.txt` fájljában van; ha hiányzik, jelezd.
- Ha az apply "Nincs új cikk" eredménnyel zárul, az nem hiba (duplikált küldés).
- A `state/processed.json` mondja meg, mi volt már feldolgozva — kézzel ne írd.
