# sEMG önkéntes jelentkezés – telepítés (kb. 5 perc)

1. Hozz létre egy új Google Sheetet (sheets.new), nevezd el pl. "sEMG jelentkezések".
2. Bővítmények → Apps Script. Másold be:
   - `Code.gs` tartalmát a Code.gs fájlba,
   - új HTML fájl `Index` néven (Fájl +  → HTML), abba az `Index.html` tartalmát.
3. Üzembe helyezés → Új üzembe helyezés → Webalkalmazás.
   Futtatás: **Én**, hozzáférés: **Bárki**. Engedélyezd a jogosultságokat.
4. A kapott webapp-linket küldd ki a jelentkezőknek.
5. Az első megnyitáskor létrejön a "Jelentkezések" lap, ide kerülnek a jelentkezések.

A lefoglalt időpontok automatikusan eltűnnek a listából.

## Emlékeztető (opcionális)
Apps Script → Eseményindítók → Új: `sendReminders`, időalapú, napi, 7–8 óra között.
Ez az aznapi mérésekre küld emlékeztetőt, és a lapon kitölti az "Emlékeztető elküldve" oszlopot.
