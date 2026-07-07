# Sajtófigyelő – heti feldolgozás (scheduled task instrukció)

Ez a referencia-instrukció a hétfő/péntek 16:00-s scheduled taskhoz.

## Lépések

1. `pip install python-docx --break-system-packages` (ha még nincs telepítve).
2. Futtasd: `python3 -B <Sajtófigyelés>/cowork/runner.py scan`
   - Ha `NINCS_UJ`: ellenőrizd az M365 connectorral, érkezett-e új "piaci körkép"
     tárgyú email az elmúlt 7 napban. Ha igen, jelezd a felhasználónak, hogy a
     docx még nem került be az incoming/ mappába. Fejezd be a futást.
3. Minden új fájlra (időrendben):
   a. `python3 -B runner.py extract "<fájlnév>"` → /tmp/sajto_input.txt
   b. Olvasd be a /tmp/sajto_input.txt-t és dolgozd fel az alábbi szabályokkal.
   c. Az eredményt írd a /tmp/sajto_articles.json fájlba.
   d. `python3 -B runner.py apply /tmp/sajto_articles.json "<fájlnév>"`
4. A végén foglald össze: hány fájl, hány új cikk, push megtörtént-e.
   Hiba esetén írd le pontosan, mi hiúsult meg.

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
3. Ha "Nem volt releváns megjelenés" áll egy alkategóriánál → hagyd ki
4. kulcsszavak: 2–5 releváns magyar kulcsszó
5. A tartalomjegyzék-blokkot (a "Tartalom" rész pontozott soraival) hagyd ki
6. Ne találj ki adatot: ami nincs a szövegben, az null vagy üres

## Fontos technikai megkötések

- A repót a runner klónozza /tmp alá — a projektmappában SOHA ne futtass gitet.
- Mindig `python3 -B`-vel futtasd a runnert (pycache-védelem).
- A PAT a projektmappa `github_pat.txt` fájljában van; ha hiányzik, jelezd.
- Ha az apply "Nincs új cikk" eredménnyel zárul, az nem hiba (duplikált küldés).
