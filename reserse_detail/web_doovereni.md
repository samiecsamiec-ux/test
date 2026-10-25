# Webové doověření položek bez abstraktu/plného textu v PubMedu

Datum: 2026-10-25
Celkem webových volání: 25 (8× WebFetch – všechna selhala, 17× WebSearch) + 3× PubMed MCP.

## DŮLEŽITÉ OMEZENÍ PROSTŘEDÍ (platí pro všechny položky)

Síťová politika tohoto cloudového prostředí **blokuje přímý přístup ke všem vydavatelským webům** (proxy vrací `CONNECT tunnel failed, response 403`; WebFetch končí `ENOTFOUND`). Konkrétně nedostupné hosty: `academic.oup.com`, `escardio.org`, `diabetesjournals.org`, `pnas.org`, `nejm.org`, `atsjournals.org`, `pmc.ncbi.nlm.nih.gov`, `pubmed.ncbi.nlm.nih.gov`, `europepmc.org`, `doi.org`, `web.archive.org`.

Důsledek: **plný text žádné z 8 položek jsem nemohl otevřít a přečíst.** Jediné, co fungovalo, bylo (a) fulltextové vyhledávání (WebSearch), které vrací úryvky („snippety") indexované z primárních stránek, a (b) PubMed MCP (abstrakty; PMC full text u obou kandidátů vrátil prázdný řetězec).

Proto rozlišuji tři stupně ověření:
- **OVĚŘENO (abstrakt, PubMed MCP)** – text pochází z oficiálního abstraktu vydavatele.
- **ČÁSTEČNĚ – snippet z primární domény** – doslovný úryvek, který vyhledávač vrátil z URL vydavatele (academic.oup.com, nejm.org, pnas.org, diabetesjournals.org). Nejde o přečtení plného textu; čísla tabulek a třídu/úroveň většinou nelze potvrdit. Doporučuji před finálním použitím dohledat v plném textu.
- **NEOVĚŘENO – nedostupné** – nic doslovného z primárního zdroje k dispozici.

Nic níže není domyšleno; kde snippet chyběl, je to výslovně napsáno.

---

## 1. 2024 ESC Guidelines – elevated BP and hypertension (McEvoy et al., EHJ 2024;45:3912–4018; DOI 10.1093/eurheartj/ehae178)

**Stav: ČÁSTEČNĚ – snippety z academic.oup.com; třída/úroveň pro OSA u rezistentní hypertenze NEOVĚŘENA.**

Zdrojová URL (snippety): https://academic.oup.com/eurheartj/article/45/38/3912/7741010 ; PDF: https://academic.oup.com/eurheartj/article-pdf/45/38/3912/59633218/ehae178.pdf

Doslovné úryvky vrácené z primární domény:

- Text sekce o OSA (číslo sekce ze snippetu nelze určit):
  > "OSAS should be suspected in patients with hypertension and suggestive symptoms (see Supplementary data online, Table S2), in all patients with resistant hypertension, and in patients with non-dipping."
  > "A simplified polysomnogram confirms the diagnosis [apnoea–hypopnoea index (AHI) > 5]."
- Tabulka sekundární hypertenze (řádek): „Obstructive sleep apnoea syndrome | Overnight ambulatory polysomnography"
- Doporučení (recommendation table ke screeningu sekundární hypertenze; třída/úroveň ve snippetu chybí):
  > "Comprehensive screening for the main causes of secondary hypertension is recommended in adults diagnosed with hypertension before the age of 40 years, except for obese young adults where it is recommended to start with an obstructive sleep apnoea evaluation."
- Rezistentní hypertenze (Recommendation Table 12 / Evidence Table 18 dle snippetu) – snippet obsahoval pouze:
  > "Patients with resistant hypertension should be considered for referral to clinical centres with expertise in hypertension management for further testing."
  > "Objective evaluation of adherence (either directly observed treatment or detecting prescribed drugs in blood or urine samples) should be considered in the clinical work-up of patients with apparent resistant hypertension, if resources allow."
  > "In patients with resistant hypertension and uncontrolled BP despite use of first-line BP-lowering therapies, the addition of spironolactone to existing treatment should be considered."

**Co NEBYLO nalezeno:** samostatné doporučení s třídou a úrovní evidence ke *screeningu/léčbě OSA u rezistentní hypertenze*. Snippety ukazují, že OSA je u rezistentní HT řešena v narativním textu („should be suspected…"), nikoli (prokazatelně) jako číslovaná rekomendace s Class/LoE. Tvrzení, že 2024 ESC HT doporučuje CPAP u středně těžké/těžké OSA, se objevilo jen v sekundárním přehledu (eurjpc) – **nepřebírat jako znění guideline**.

---

## 2. 2024 ESC Guidelines – atrial fibrillation (Van Gelder et al., EHJ 2024;45:3314–3414; DOI 10.1093/eurheartj/ehae176)

**Stav: ČÁSTEČNĚ – znění dvou doporučení ze snippetů; třída jen u jednoho (IIb), úroveň evidence NEOVĚŘENA; číslo tabulky NEOVĚŘENO.**

Zdrojová URL (snippety): https://academic.oup.com/eurheartj/article/45/36/3314/7738779 (sekce 5.5 „Obstructive sleep apnoea", str. 3340 dle obsahu; dále 10.5.5 v rámci primární prevence)

Doslovné úryvky vrácené z primární domény:

- Doporučení (recommendation table k komorbiditám, AF-CARE „C"):
  > "Management of obstructive sleep apnoea may be considered as part of comprehensive management of risk factors in individuals with AF to reduce recurrence and progression."
  Třída: **IIb** (snippet obrázku guideline: „Management of OSA (Class IIb)"). Úroveň evidence: **ve snippetu chybí – NEOVĚŘENO**.
- Doporučení ke screeningu:
  > "When screening for obstructive sleep apnoea in individuals with AF, using only symptom-based questionnaires is not recommended."
  Třída: ve snippetu chybí (formulace „is not recommended" odpovídá třídě III, ale **třída ani úroveň nebyly potvrzeny**).
- Narativní text sekce 5.5:
  > "it may be reasonable to screen for OSA in patients where a rhythm control strategy is being pursued"
  > polysomnography or home sleep apnoea testing „are suggested in preference to screening questionnaires, which are poor predictors of moderate-to-severe OSA"
  > (sekundárně, PMC11666470): „randomized studies on OSA treatment did not demonstrate clear benefit in patients with AF"

**Pozor:** Nenašel jsem žádné doporučení typu „léčba CPAP před rytmus-kontrolní strategií" jako samostatnou číslovanou rekomendaci – screening při rytmus-kontrolní strategii je v snippetu formulován jako „may be reasonable" v textu, nikoli jako rekomendace s třídou.

---

## 3. 2021 ESC Guidelines – heart failure (McDonagh et al., EHJ 2021;42:3599–3726; DOI 10.1093/eurheartj/ehab368)

**Stav: ČÁSTEČNĚ – jedna věta ze snippetu; třída/úroveň a číslo sekce pro 2021 NEOVĚŘENY.**

Zdrojová URL (snippet): https://academic.oup.com/eurheartj/article/42/36/3599/6358045

Doslovný úryvek (vyhledávač jej atribuoval sekci o plicních chorobách a sleep-disordered breathing 2021 guideline; v druhém vyhledávání se identická věta objevila i u 2026 HF guideline, takže atribuci ke konkrétnímu ročníku nelze zcela garantovat):
> "The use of adaptive servo-ventilation in patients with HFrEF and predominantly CSA is not recommended, based on the results of SERVE-HF"

Pro **2021** verzi jsem nezískal: číslo sekce, tabulku doporučení, Class/LoE, ani text o OSA/CPAP u HF.

Pro orientaci (NE jako náhrada za 2021): u **2026 ESC HF guideline** (https://academic.oup.com/eurheartj/article/47/38/5346/8766302) snippet uvádí:
> "ASV is not recommended in patients with HFrEF and sleep-disordered breathing with predominant central sleep apnoea because of an increased risk of CV and all-cause death" – Class III, Level A
> "ASV may be considered in patients with HFrEF and sleep-disordered breathing with predominant obstructive sleep apnoea to improve sleep quality, health-related QoL, and symptoms" – IIb/C
Pokud zadání cituje 2021 guideline, je nutné Class/LoE pro 2021 ověřit v plném textu.

---

## 4. ADA Standards of Care in Diabetes—2026, Section 4 (Diabetes Care 2026;49(Suppl 1):S61–S88; DOI 10.2337/dc26-S004)

**Stav: NEOVĚŘENO – nedostupné (plný text), a navíc indicie, že číslovaná rekomendace ke sleep health NENÍ v Sekci 4.**

- PubMed MCP: článek má PMCID **PMC12690184** (PMID 41358897), ale `get_full_text_article` vrátil prázdný full_text (jen abstrakt bez doporučení).
- URL primárního zdroje (nedostupná z tohoto prostředí): https://diabetesjournals.org/care/article/49/Supplement_1/S61/163931/4-Comprehensive-Medical-Evaluation-and-Assessment
- Snippety z diabetesjournals.org:
  - Sekce 4 má podsekci „Obstructive Sleep Apnea" (Assessment of Comorbidities) a OSA figuruje v Table 4.1; narativně: „the assessment of sleep pattern and duration, bone health, and sexual health should also be considered as an important part of the comprehensive evaluation." **Žádná číslovaná rekomendace 4.x k OSA s gradem se ve snippetech neobjevila.**
  - Číslovaná rekomendace ke sleep health je dle snippetů (Summary of Revisions 2026, https://diabetesjournals.org/care/article/49/Supplement_1/S6/163930/Summary-of-Revisions-Standards-of-Care-in-Diabetes) v **Sekci 5**: **Recommendation 5.56** – screening for sleep health in people with prediabetes or diabetes and in those at risk, včetně sleep disorders; **5.57** – „Counsel people with diabetes to practice sleep-promoting routines and habits." **Grade (A/B/C/E) nebyl ve snippetu – NEOVĚŘENO.**
  - Pediatrie: Recommendation 14.71 (screening příznaků spánkové apnoe u dětí/adolescentů) – grade neověřen.
- Z ročníku 2024 (Section 4, dc24s004) snippet: „Individuals with symptoms suggestive of obstructive sleep apnea, such as excessive daytime sleepiness, snoring, and witnessed apnea, should be considered for screening." – to je starší ročník, nikoli 2026.

**Doporučení pro autora:** ověřit, zda citace „Section 4" + „doporučení ke screeningu OSA s gradem" není ve skutečnosti Rec. 5.56 (Section 5, Facilitating Positive Health Behaviors…). Grade zatím nelze uvést.

---

## 5. Tasali E et al., PNAS 2008;105:1044–9 (DOI 10.1073/pnas.0706446105; PMC2242689)

**Stav: ČÁSTEČNĚ.** Abstrakt OVĚŘEN (PubMed MCP, PMC2242689); IVGTT OVĚŘENO snippetem z pnas.org; **n účastníků NEOVĚŘENO**; „3 noci" a „25 %" jen ze sekundárního zdroje.

- Abstrakt (dle PubMed, [DOI](https://doi.org/10.1073/pnas.0706446105)), doslovně:
  > "Here we show that, in young healthy adults, all-night selective suppression of SWS, without any change in total sleep time, results in marked decreases in insulin sensitivity without adequate compensatory increase in insulin release, leading to reduced glucose tolerance and increased diabetes risk."
  > "Importantly, the magnitude of the decrease in insulin sensitivity was strongly correlated with the magnitude of the reduction in SWS."
- Metoda (snippet z https://www.pnas.org/doi/10.1073/pnas.0706446105):
  > "glucose tolerance was quantified by the rate of decline in glucose levels after i.v. glucose injection; insulin sensitivity (S.I.) and insulin secretion [the 'acute insulin response to glucose' (AIRg)] by using minimal model analysis of blood glucose and insulin levels measured during the ivGTT"
  → metoda = **IVGTT s minimal model analysis** – potvrzeno.
- Počet nocí a procentuální pokles: pouze sekundárně (PNAS 2011, 10.1073/pnas.1101338108, odkazující na Tasali 2008): „after 3 nights of SWS suppression, insulin sensitivity, glucose tolerance, and the disposition index decreased compared with baseline (by 25%, 23%, and 20%, respectively)"; komentář PNAS 2008;105:1107: „after as few as three nights of disruption of SWS". **Z plného textu Tasali 2008 nepotvrzeno.**
- **n účastníků: NEOVĚŘENO** (v žádném primárním snippetu se číslo neobjevilo).

---

## 6. McEvoy RD et al., SAVE, NEJM 2016;375:919–31 (DOI 10.1056/NEJMoa1606599)

**Stav: ČÁSTEČNĚ – doslovné úryvky vrácené vyhledávačem z nejm.org (plný text/paywall nedostupný).**

Zdrojová URL: https://www.nejm.org/doi/full/10.1056/NEJMoa1606599 (PDF: https://www.nejm.org/doi/pdf/10.1056/NEJMoa1606599)

Doslovné úryvky:
- Primární endpoint ITT: „A primary end-point event was confirmed in 436 participants — 229 (17.0%) in the CPAP group and 207 (15.4%) in the usual-care group" … „hazard ratio with CPAP, 1.10; 95% CI, 0.91 to 1.32; P=0.34"
- Adherence: „the mean duration of adherence to CPAP therapy was 3.3 hours per night"; definice adherence: „an average of 4 hours or more of treatment per night over the first 2 years"
- Propensity-matched analýza:
  > "One-to-one propensity-score matching was performed to compare 561 patients who were adherent to CPAP therapy with 561 patients in the usual-care group."
  > „Among these propensity-score–matched patients, 184 primary end-point events occurred" – primární endpoint **HR 0.80 (95% CI 0.60 to 1.07; P=0.13)**
  > "the patients who were adherent to CPAP therapy had a lower risk of stroke than those in the usual-care group (hazard ratio, 0.56; 95% CI, 0.32 to 1.00; P=0.05)"
  > nepředspecifikovaný kompozit cerebrálních příhod: HR 0.52 (95% CI 0.30–0.90)
  > „neither outcome was adjusted for multiple testing"

Čísla souhlasí se zadáním (primární HR 0.80; CMP HR 0.56). Číslo tabulky/suplementu nelze potvrdit.

---

## 7. Peker Y et al., RICCADSA, AJRCCM 2016;194:613–20 (DOI 10.1164/rccm.201601-0088OC)

**Stav: průměrná adherence CPAP (h/noc) – NEOVĚŘENO – nedostupné.**

Zdrojová URL (nedostupná): https://www.atsjournals.org/doi/full/10.1164/rccm.201601-0088OC

- Z abstraktu (snippet atsjournals.org): 244 pacientů, auto-CPAP n=122 vs. bez CPAP n=122; „adjusted on-treatment analysis … CPAP for >4 versus <4 hours per night or did not receive treatment (hazard ratio, 0.29)"; podrobnější snippet z téže domény: **HR 0.29, 95% CI 0.10–0.86, P = 0.026**; cut-off 3 h/noc nesignifikantní; ≥5 h/noc HR 0.34, P = 0.075.
- Průměrné hodiny/noc: v primárním článku nenalezeno. Hodnota „3.3 h/noc" se objevila pouze v dopisu/komentáři (https://www.atsjournals.org/doi/10.1164/rccm.201701-0239LE) a není jisté, zda se vztahuje k RICCADSA (stejné číslo je známo ze SAVE). Sekundární přehled (ERS Breathe) uvádí odlišný HR 0.17 (0.03–0.81) – pravděpodobně z jiné analýzy/publikace RICCADSA. **Nepřebírat bez ověření v plném textu.**

---

## 8. Kendzerska T et al., AJRCCM 2014;190:218–25 (DOI 10.1164/rccm.201312-2209OC)

**Stav: ČÁSTEČNĚ – bodový odhad z abstraktu (jako „30 % higher hazard"); přesné HR 1.31 (95% CI 1.07–1.61) jen ze sekundárního zdroje.**

Zdrojová URL (nedostupná): https://www.atsjournals.org/doi/10.1164/rccm.201312-2209OC

- Abstrakt (snippet atsjournals.org / ICES): historická kohorta 8 678 pacientů bez diabetu (St. Michael's Hospital, Toronto, 1994–2010), medián sledování 67 měsíců, 1 017 (11.7 %) incidentní diabetes, kumulativní 5letá incidence 9.1 % (95% CI 8.4–9.8 %);
  > "in fully-adjusted models, patients with AHI > 30 had a 30% higher hazard of developing diabetes than those with AHI < 5"
- Přesné HR s CI: **1.31 (95% CI 1.07–1.61)** uvádí pouze tabulka sekundární meta-analýzy (ERJ Open Res 2023;9:00401-2022, https://publications.ersnet.org/content/erjor/9/2/00401-2022). V primárním textu (Table 3) **neověřeno**.
- Doplňkově (editorial AJRCCM 2014;190:133): CPAP léčba nebyla v této studii spojena se snížením rizika diabetu.

---

## Souhrn

| # | Položka | Stav |
|---|---------|------|
| 1 | ESC HT 2024 – OSA | ČÁSTEČNĚ (narativní text „should be suspected… in all patients with resistant hypertension"); Class/LoE pro OSA u RH **neověřeno** – pravděpodobně neexistuje samostatná rekomendace |
| 2 | ESC AF 2024 – OSA | ČÁSTEČNĚ: „Management of OSA may be considered…" = IIb (LoE neověřeno); „using only symptom-based questionnaires is not recommended" (Class/LoE neověřeno) |
| 3 | ESC HF 2021 – SDB/ASV | ČÁSTEČNĚ: věta „ASV… predominantly CSA is not recommended, based on SERVE-HF"; Class/LoE a sekce pro 2021 **neověřeno** |
| 4 | ADA 2026 Section 4 | **NEOVĚŘENO – nedostupné**; indicie, že sleep-health rekomendace je 5.56 (Section 5), nikoli Section 4; grade neznámý |
| 5 | Tasali PNAS 2008 | ČÁSTEČNĚ: abstrakt + IVGTT ověřeny; n **neověřeno**; 3 noci/25 % jen sekundárně |
| 6 | SAVE NEJM 2016 | ČÁSTEČNĚ (snippety nejm.org): PS-matched primární HR 0.80 (0.60–1.07), CMP HR 0.56 (0.32–1.00), 561 vs 561, adherence ≥4 h; plný text nedostupný |
| 7 | RICCADSA 2016 | Průměrná adherence h/noc **NEOVĚŘENO – nedostupné**; HR ≥4 h 0.29 (0.10–0.86) z abstraktu |
| 8 | Kendzerska 2014 | ČÁSTEČNĚ: „30% higher hazard" z abstraktu; HR 1.31 (1.07–1.61) jen sekundárně |
