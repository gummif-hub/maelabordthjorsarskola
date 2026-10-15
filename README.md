# Mælaborð Þjórsárskóla

Vefsíða sem les tölur úr Excel-skjali (`gogn/maelabord.xlsx`) og birtir þær sem mælaborð. Hver liður í mælaborðinu hefur sinn flipa í Excel-skjalinu og sinn flipa á síðunni.

Engin gagnagrunnur og enginn netþjónn. Síðan er venjuleg HTML-síða sem keyrir á GitHub Pages, og Excel-skjalið er lesið í vafra þess sem opnar hana.

## Skrárnar

| Skrá | Hlutverk |
|---|---|
| `index.html` | Síðan sjálf |
| `css/style.css` | Útlit |
| `js/app.js` | Les Excel-skjalið, reiknar og teiknar |
| `js/config.js` | Stilling: hvar gagnaskjalið er |
| `lib/` | SheetJS (les Excel) og Chart.js (gröf). Vistað hér svo síðan virki óháð öðrum vefjum |
| `gogn/maelabord.xlsx` | Gögnin. Það eina sem þarf að uppfæra |

## Setja á GitHub Pages (einu sinni)

1. Stofnaðu nýtt repository á GitHub, t.d. `maelabord`.
2. Veldu **Add file → Upload files** og dragðu allt innihald möppunnar inn (líka `lib`, `css`, `js`, `gogn` og skrána `.nojekyll`). Smelltu á **Commit changes**.
3. Farðu í **Settings → Pages**. Undir *Build and deployment* veldu **Deploy from a branch**, greinina `main` og möppuna `/ (root)`. Vistaðu.
4. Eftir 1–2 mínútur er síðan komin á `https://<notandi>.github.io/maelabord/`.

## Uppfæra tölurnar

1. Opnaðu `gogn/maelabord.xlsx` í Excel og skráðu nýjar tölur í gulu reitina.
2. Á GitHub: opnaðu möppuna `gogn`, veldu **Add file → Upload files** og dragðu nýja skjalið inn með sama nafni. **Commit changes**.
3. Síðan sýnir nýju tölurnar eftir 1–2 mínútur.

Til að skoða breytingar áður en þær fara á vefinn: opnaðu síðuna og smelltu á **Opna Excel-skjal**. Þá er skjalið af tölvunni lesið í staðinn, en aðeins í þínum vafra.

## Nýtt skólaár

Ekkert þarf að endurgera. Haltu áfram að bæta línum neðst í flipana. Síðan finnur skólaárin sjálf og sýnir nýjasta árið fyrst. Fyrri ár má velja í fellilistanum efst.

## Nýr mælikvarði

Bættu línu í flipann **Mælikvarðar** með Gagnalind:

- **Skráningar**: talið úr flipanum Skráningar (veldu tegund með sama heiti).
- **Ársmælingar**: gildið er skráð í flipanum Ársmælingar.

Mælikvarðinn birtist þá sjálfkrafa í skorkortinu. Gagnalindin **Reiknað** er aðeins fyrir mælikvarða sem `js/app.js` reiknar úr öðrum flipum (kóðarnir `maeting`, `forfoll_mannad`, `akstur_atvik`, `frettir_heimasida`, `aefingar`, `utinam_argangar`, `list_verk_stig`, `ithrottir_stig`, `art`, `bofs`, `leyfisbref`, `rekstur_fravik`, `innramat_metid`). Mælikvarði sem er hætt að nota fær **Virkur = Nei**.

## Skýrsla fyrir skólanefnd

Veldu skólaár og smelltu á **Prenta skýrslu**. Allir liðir prentast, hver á sinni síðu. Veldu *Vista sem PDF* í prentglugganum og sendu PDF-skjalið.

## Persónuvernd og aðgangur

**GitHub Pages síður eru opnar öllum**, og það á líka við um Excel-skjalið í möppunni `gogn`: hver sem þekkir slóðina getur sótt það. Í skjalinu eru engin nöfn, aðeins heildartölur, en þar eru tölur um stuðning við nemendur og fjárhag.

Ef gögnin eiga ekki að vera opin:

1. Settu `gagnaskra: ""` í `js/config.js` og eyddu `gogn/maelabord.xlsx` úr repository-inu.
2. Geymdu Excel-skjalið á OneDrive eða SharePoint skólans.
3. Hver sem á að sjá mælaborðið opnar síðuna og velur skjalið með **Opna Excel-skjal**. Skjalið er lesið í vafranum og fer hvergi annað.

## Prófa á eigin tölvu

Ef `index.html` er opnuð beint úr möppu getur vafrinn ekki sótt gagnaskjalið sjálfur. Smelltu þá á **Opna Excel-skjal** og veldu `gogn/maelabord.xlsx`.
