# Sajtófigyelő – CoWork Instrukció

> Ezt a fájlt add hozzá a CoWork projekthez mappaként.
> A CoWork minden hétfőn reggel automatikusan végrehajtja az alábbi lépéseket.

---

## Mit csinálj minden hétfőn reggel?

### 1. LÉPÉS – Legújabb Word fájl letöltése az Outlookból

Keresd meg a legutóbbi emailt, amit a PR ügynökség küldött (feladó: a `PR_SENDER_EMAIL` értéke).
Az emailhez csatolt `.docx` fájlt mentsd le a projekt mappájába `sajtofigyelo_het.docx` névvel.
Ha nincs új email az előző hét hétfője óta, állj meg és jelezd.

### 2. LÉPÉS – Word szöveg kinyerése

Futtasd le a következő parancsot a projekt mappájában:

```
python generate_local.py sajtofigyelo_het.docx
```

Ez létrehozza a `sajto_input.txt` fájlt.

### 3. LÉPÉS – Feldolgozás Claude-dal

Olvasd be a `sajto_input.txt` tartalmát és dolgozd fel az alábbi utasítás szerint:

---

**RENDSZER UTASÍTÁS:**

Te egy sajtófigyelő feldolgozó vagy. A PR ügynökség által küldött Word dokumentum
már strukturált: szekciókra és alkategóriákra van osztva.

A dokumentum felépítése:
- Főszekciók: "Prohumanra vonatkozó megjelenések", "Munkaerőpiaci hírek és elemzések",
  "Versenytársakra vonatkozó megjelenések"
- Alkategóriák: pl. "Trenkwalder", "Randstad", "WHC", "Munkaerő-kölcsönzés" stb.
- Minden cikknél: Cím (néha [URL: ...] jelöléssel), forrás és dátum zárójelben:
  "(Médium, 2026.03.20.)", majd összefoglaló szöveg

A főkategória meghatározása a szekció alapján:
- "prohuman"    → "Prohumanra vonatkozó megjelenések"
- "versenytars" → "Versenytársakra vonatkozó megjelenések"
- "piac"        → "Munkaerőpiaci hírek és elemzések"

Nyerd ki az összes cikket és írd az eredményt a `sajto_output.json` fájlba,
KIZÁRÓLAG érvényes JSON formátumban, más szöveg nélkül:

```json
{
  "generated": "YYYY-MM-DD",
  "total": 0,
  "cikkek": [
    {
      "cim": "...",
      "url": "https://...",
      "forras": "...",
      "datum": "2026-03-20",
      "fokategoria": "piac",
      "alkategoria": "...",
      "osszefoglalo": "...",
      "kulcsszavak": ["..."],
      "het": "2026-03-17",
      "id": "ART-0001"
    }
  ]
}
```

SZABÁLYOK:
1. cim és osszefoglalo: CSAK sima szöveg – semmi markdown (#, *, [, ], **)
2. url: ha nincs link, értéke null
3. het: az adott dátum hetének hétfője YYYY-MM-DD formátumban
4. Ha "Nem volt releváns megjelenés" szerepel → hagyd ki azt az alkategóriát
5. kulcsszavak: 2–5 releváns magyar kulcsszó
6. id: ART-0001 formátum, sorszám szerint növekvő

---

### 4. LÉPÉS – HTML frissítése

Futtasd:

```
python generate_local.py
```

Ez beolvassa a `sajto_output.json`-t és frissíti a `docs/index.html`-t.

### 5. LÉPÉS – Git push a GitHub-ra

Futtasd:

```
git add docs/index.html
git commit -m "🗞️ Sajtófigyelő frissítve: YYYY-MM-DD"
git push
```

(A dátumot helyettesítsd be az aktuális dátummal.)

---

## Ha valami nem sikerül

- **Nincs új email:** Jelezd, hogy ezen a héten nem érkezett új sajtófigyelő.
- **A Word fájl nem nyílik meg:** Próbáld meg újra; ha nem sikerül, jelezd.
- **A JSON érvénytelen:** Próbáld meg újragenerálni; ha nem sikerül, mentsd el a hibát és jelezd.
- **A git push nem sikerül:** Ellenőrizd az internetkapcsolatot és a git beállításokat.

---

## Ütemezés

Ezt a feladatot **minden hétfőn 09:00-kor** kell elvégezni.

Windows Task Scheduler beállítás (egyszeri elvégzés után automatikus):
- Trigger: Weekly, Monday, 09:00
- Action: Start a program → `cowork` vagy a CoWork indítóscriptje
- Start in: a projekt mappa elérési útja
