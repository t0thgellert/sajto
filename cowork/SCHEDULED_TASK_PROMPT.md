# A két ütemezett feladat promptja

**Miért ilyen rövid:** eddig a teljes instrukció be volt másolva mindkét task
promptjába, és a `cowork/TASK_PROMPT.md` csak „referencia" volt. Így minden
változást három helyen kellett átírni, és a taskok elavultak (2026-07-28b, míg a
TASK_PROMPT.md már 2026-08-05e). Innentől a task promptja csak *hova nézz* +
*mikor mit* — a *hogyan* egyetlen helyen, a `TASK_PROMPT.md`-ben él.

Másold be az alábbi szövegeket a Claude app → ütemezett feladatok felületén,
a meglévő prompt HELYÉRE.

---

## `sajtofigyelo-heti-frissites` (hétfő 16:01)

```
Sajtófigyelő heti feldolgozás — hétfői futás.

A projektmappa Windows-útvonala:
C:\Users\toth.gellert\Claude\Projects\Sajtófigyelés
A bash sandboxban ez egy csatolt mappa; a pontos mount-útvonalat a
session-leírásból olvasd ki.

1. Olvasd be a `<mappa>/cowork/TASK_PROMPT.md` fájlt, és PONTOSAN azt kövesd.
   Ez az egyetlen érvényes instrukció; ha ellentmond ennek a promptnak,
   a TASK_PROMPT.md nyer.
   Ha a fájl nem olvasható, ÁLLJ MEG és jelezd — ne improvizálj feldolgozást.

2. A hétfői futás sajátja: a 4. lépés (link-pótlás) MINDKÉT körét futtasd le,
   ebben a sorrendben:
   a) friss kör: `relink-list --het=<a most feldolgozott hét hétfője>`
   b) backlog-kör: `relink-list --limit=10`
   A `--limit=10`-et ne emeld meg.

3. Ha van `<mappa>/work/sajto_relink_*.json` (NEM `.done.json` végű), az egy
   előkészített, még ki nem pusholt link-csomag. Így kezeld:
   a) először `relink-apply "<fájl>" --dry-run --fix-sources`
   b) ha a kimenet 1 vagy több „link beírva"-t jelez, futtasd újra `--dry-run`
      NÉLKÜL (de `--fix-sources`-szal), és nevezd át a fájlt `.done.json`-ra
   c) ha 0 link beírva (mert már mind bent van), CSAK nevezd át `.done.json`-ra,
      élesben ne futtasd
   Miért az átnevezés: élesben újra beadva a runner „a cikknek MÁR van URL-je"
   indokkal elutasítja őket, és feleslegesen növeli a `url_probak` számlálót —
   3 után a cikk kiesne a backlog-listából.

4. A záró összefoglalóban a TASK_PROMPT.md „Zárás" szakaszában felsoroltakat
   add meg, és külön emeld ki:
   - hány link került be, hány cikk lett `nincs`-re zárva, hány lett elutasítva
     és milyen indokkal (az `ELUTASÍTVA` sorokat szó szerint),
   - mennyi maradt a backlogban (a `relink-list` stderr-sora),
   - minden `FIGYELEM: <domain> nincs a DOMAIN_MAP-ban` sort.
```

---

## `sajtofigyelo-heti-frissites-pentek` (péntek 16:39)

```
Sajtófigyelő heti feldolgozás — pénteki futás.

A projektmappa Windows-útvonala:
C:\Users\toth.gellert\Claude\Projects\Sajtófigyelés
A bash sandboxban ez egy csatolt mappa; a pontos mount-útvonalat a
session-leírásból olvasd ki.

1. Olvasd be a `<mappa>/cowork/TASK_PROMPT.md` fájlt, és PONTOSAN azt kövesd.
   Ez az egyetlen érvényes instrukció; ha ellentmond ennek a promptnak,
   a TASK_PROMPT.md nyer.
   Ha a fájl nem olvasható, ÁLLJ MEG és jelezd — ne improvizálj feldolgozást.

2. A pénteki futás sajátja: a 4. lépésből CSAK a friss kört futtasd
   (`relink-list --het=<a most feldolgozott hét hétfője>`).
   Backlog-kört pénteken NE — az a hétfői futás dolga.

3. A záró összefoglalóban a TASK_PROMPT.md „Zárás" szakaszában felsoroltakat
   add meg, és külön emeld ki:
   - hány link került be, hány cikk lett `nincs`-re zárva, hány lett elutasítva
     és milyen indokkal (az `ELUTASÍTVA` sorokat szó szerint),
   - minden `FIGYELEM: <domain> nincs a DOMAIN_MAP-ban` sort.
```

---

## Amit ezzel megnyersz

- A `TASK_PROMPT.md` egyetlen igazságforrás. A következő módosításnál nem kell
  ütemezett feladatot szerkeszteni.
- A két task között csak a valódi különbség marad (hétfőn van backlog-kör,
  pénteken nincs) — eddig a teljes szöveg duplikálva volt, tehát bármelyik
  elcsúszhatott a másiktól észrevétlenül.
- A „ha a TASK_PROMPT.md nem olvasható, állj meg" szabály miatt egy elmozdult
  vagy átnevezett projektmappa nem improvizált feldolgozást okoz, hanem
  látható hibát.

## Ellenőrzés az átállás után

Futtasd le kézzel („Run now") mindkettőt. Elvárt kimenet, ha nincs új körkép:

- `NINCS_UJ`, illetve „nincs teendő, az utolsó körkép (`<fájlnév>`) már
  feldolgozva",
- a hétfőin ezen felül a backlog-kör eredménye (jelenleg 10 cikk vár).

Ha az összefoglalóban nem szerepel a link-pótlás egyáltalán, akkor a task nem
olvasta be a `TASK_PROMPT.md`-t — akkor a mappa-útvonal a hibás a promptban.
