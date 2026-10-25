# FÁZE 3 – Prezentace (.pptx) a text pro řečníka

**Soubor:** `OSA_inzulinova_rezistence_kardiolog.pptx` (29 slidů: 22 obsahových včetně 2 rezervovaných + 7 slidů literatury; 16:9, 13,33 × 7,5").
**Generátor:** `build/build.js` (pptxgenjs) + `build/notes.json` (text pro řečníka) + `build/refs.json` (metadata PubMed pro Vancouver). Znovu sestavit: `cd build && NODE_PATH=$PWD/node_modules node build.js`.

Text pro řečníka je vložen do pole poznámek (speaker notes) každého slidu. Každé číslo v poznámkách má v závorce zkrácený zdroj (autor, rok) dohledatelný v `FAZE1_tabulka_evidence.md`. Na slidech je zkrácená citace (autor, časopis, rok) a štítek úrovně evidence vpravo nahoře.

## Časový plán

| # | Slide | Čas (min) | Slov v poznámkách |
|---|---|---|---|
| 1 | Úvod: OSA a inzulinová rezistence (3 čísla) | 1,0 | 182 (úvodní, delší) |
| 2 | OSA u našich pacientů: rezistentní HTN, FS, HF | 1,5 | 149 |
| 3 | Překryv OSA – obezita – DM2: kolik z toho je obezita? | 1,5 | 161 |
| 4 | Mechanismy I: intermitentní hypoxie a fragmentace spánku | 1,5 | 144 |
| 5 | Mechanismy II: sympatikus, HPA, zánět, adipokiny, MASLD | 1,5 | 165 |
| 6 | Obousměrnost – asymetricky | 1,5 | 148 |
| 7 | Observační evidence: MA prospektivních kohort | 1,5 | 139 |
| 8 | Dávka–odpověď, nezávislost na BMI, protiváha | 1,5 | 150 |
| 9 | AHI není všechno: hypoxic burden | 1,0 | 164 |
| 10 | Intervence I: CPAP a IR u nediabetiků | 1,5 | 148 |
| 11 | Intervence II: CPAP a HbA1c u DM2 | 1,5 | 167 |
| 12 | Adherence jako dávka | 1,5 | 157 |
| 13 | Intervence III: redukce hmotnosti vs. CPAP | 1,5 | 170 |
| 14 | Intervence IV: inkretiny (SCALE, SURMOUNT-OSA, TRIUMPH-OSA) | 1,5 | 174 |
| 15 | Tvrdé KV outcomy: SAVE, ISAACC, RICCADSA | 1,5 | 142 |
| 16 | Metaanalýzy a IPD MA; adherence | 1,5 | 159 |
| 17 | Proč jsou RCT neutrální | 1,0 | 135 |
| 18 | Klinická praxe I: koho screenovat | 1,5 | 161 |
| 19 | Klinická praxe II: STOP-BANG, HSAT, kdy referovat | 1,5 | 178 |
| 20 | **Rezervovaný A – Vlastní data CPM (prázdný)** | 1,0 | instrukce 84 |
| 21 | **Rezervovaný B – Vlastní data CPM (prázdný)** | 1,0 | instrukce 86 |
| 22 | Take-home messages a otevřené otázky | 1,5 | 202 (závěrečný, delší) |
| 23–29 | Literatura (Vancouver, PMID/DOI), 113 položek | 0 | – |
| | **Celkem** | **30,5** | ≈ 3 500 slov mluveného textu |

Při tempu ~115–120 slov/min odpovídá objem poznámek ~30 min; slidy 1 a 22 jsou záměrně delší (úvod, závěr), ostatní 135–178 slov (1–1,5 min).

## Rezervované slidy (20, 21)

- Na slidu pouze nadpis „Vlastní data – Centrum preventivní medicíny (doplníme)“, štítek „placeholder A/B“ a prázdný čárkovaný rámeček pro graf/tabulku. Žádná data.
- V poznámkách pouze „DOPLNIT: vlastní data CPM“ + návrh typu dat (A: prevalence vysokého rizika OSA dle STOP-BANG u pacientů s IR/prediabetem/DM2; B: vztah ODI/AHI a HOMA-IR či HbA1c v kohortě CPM, nebo výtěžnost screeningu) a můstek k dalšímu slidu.

## Co je na slidech záměrně označeno jako neověřené

- Třída/úroveň evidence ESC 2024 (hypertenze), úroveň evidence ESC 2024 FS (IIb uvedeno), znění a stupeň ADA 2026 – označeno „k ověření v plném textu“ (slide 18, poznámky).
- Tasali 2008: procento poklesu inzulinové senzitivity se neuvádí (slide 4).
- Adherence u Shaw 2016, Martínez-Cerón 2016, Lam 2017, Banghøj 2021, Mokhlesi 2017: označeno hvězdičkou „dle tab. 1 v MA Herth 2023 / Wu 2026“ (slide 12).
- SAVE propensity-matched analýza: čísla neuvedena (slide 15).

## Technická kontrola

- `validate.py`: All validations PASSED.
- Vizuální kontrola všech 29 slidů po renderu (LibreOffice → PDF → JPG): bez přesahů a kolizí.
- Forest ploty jsou kresleny z hodnot v tabulce FÁZE 1 (bodový odhad + 95 % CI); kde CI chybí, zobrazen jen bod (Foster 2009, SE).
