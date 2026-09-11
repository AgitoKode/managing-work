# Metodedokumentasjon – AGA-analyse av Ført arbeid – utvidet

Denne filen beskriver hvordan `aga_analyse.py` (og pakken `aga_lib/`) behandler
RecMan-eksporten og bygger `AGA_Rapport_2026.xlsx`. Se også
`documentation/Kildeoversikt.md` for kildehierarkiet og
`README.md` for hvordan løsningen kjøres.

## 1. Hvilken fil som ble analysert

Løsningen leter etter kildefilen i denne rekkefølgen (se `aga_lib/file_discovery.py`):

1. `.\input\Ført arbeid - utvidet.xlsx` (eller enhver fil i `.\input` som inneholder
   både "ført arbeid" og "utvidet" i filnavnet – filnavnet kan ha suffikser som dato).
2. Samme filnavnmønster i prosjektets rotmappe.
3. Enhver annen `.xlsx`/`.xlsm`-fil i `.\input`.
4. Enhver annen `.xlsx`/`.xlsm`-fil i prosjektets rotmappe.

Filer med navn som inneholder `AGA_Rapport`, `Prosjektregister`, `Avviksrapport`,
`Kilderegister`, `Kontrollrapport` eller `AGA_oppslagsbehov` ekskluderes alltid,
slik at løsningen aldri leser sin egen forrige rapport som ny input. Finnes ingen
egnet fil, stoppes programmet med en norsk feilmelding som lister de kontrollerte
mappene.

Kildefilen åpnes **read-only** (`openpyxl.load_workbook(..., read_only=True)`) og
skrives aldri tilbake til – dette er verifisert med en automatisk test som
sammenligner SHA-256 av kildefilen før og etter kjøring.

## 2. Hvilke ark som ble brukt

Alle synlige ark leses (`aga_lib/excel_io.py`). For hvert ark:

- Headerraden søkes automatisk opp blant de 10 første radene, ved å telle hvor
  mange celler som matcher et kjent feltalias (se kolonnemapping under).
- Arket regnes som et "dataark" dersom minst 5 standardfelt kunne kobles til
  kolonneoverskriftene. Rene metadata-ark (f.eks. "Søkekriterier" i RecMan-
  eksporten, som beskriver hvilket filter eksporten ble kjørt med) faller
  dermed automatisk utenfor og brukes ikke som datagrunnlag.
- Blant dataarkene brukes det arket med flest rader som hovedgrunnlag.

## 3. Kolonnemapping

`aga_lib/column_mapping.py` normaliserer kolonnenavn (små bokstaver, punktum og
bindestrek fjernet, ekstra mellomrom kollapset) og slår dem opp mot en liste av
kjente alias per standardfelt (se `FELT_ALIASER`). Resultatet – hvilken original
kolonne som ble koblet til hvilket standardfelt, og hvilke kolonner som IKKE ble
gjenkjent – skrives til arket **Kolonnemapping** i rapporten.

Kritiske feltgrupper (dato/år, ansattidentifikator, prosjekt-/jobb-id,
prosjekt/jobb/bedrift-navn, geografisk informasjon, lønnsbeløp) må ha minst ett
gjenkjent felt hver. Mangler en hel gruppe, stopper programmet med en tydelig
feilmelding i stedet for å gjette videre.

## 4. Datarensing og datokonvertering

- **Arbeidsdato** = kolonnen `Dato` (arbeidsdatoen slik RecMan har registrert
  den), med `Første dag` som reserve dersom `Dato` mangler. `Dato registrert`
  og en eventuell fakturadato brukes ALDRI som arbeidsdato.
- Excel-serienummer, tekst-datoer (`dd.mm.åååå`, ISO, m.fl.) og `datetime`-
  objekter konverteres alle via `aga_lib/dates.til_dato`, som returnerer `None`
  i stedet for å gjette dersom verdien ikke lar seg tolke.
- Radene grupperes i **Termin** 1–6 (jan/feb, mar/apr, mai/jun, jul/aug,
  sep/okt, nov/des) via `config.json` sin `termin_maaneder`.
- Rader for andre år enn `config.json` sin `analyseaar` (2026) beholdes i det
  innleste datasettet (de slettes aldri), men ekskluderes fra hovedrapportens
  beregninger. Antallet ekskluderte rader rapporteres i Kjøremetadata og
  Lederoppsummering.

## 5. Prosjektidentifikasjon

`aga_lib/projects.py` grupperer rader til ett prosjektregister:

- **Prosjektnr.** er primær nøkkel. **Jobbnr.** brukes som sekundær nøkkel når
  Prosjektnr. mangler. Prosjektnavn brukes aldri alene som nøkkel når et
  nummer finnes.
- For hvert prosjekt kontrolleres: flere postnumre, flere prosjektnavn for
  samme nummer, samme navn med flere numre, og manglende/ugyldig postnummer.
  Alt dette vises i arket **Prosjektregister**.

## 6. Kommuneidentifikasjon

`aga_lib/municipality.py` implementerer prioriteringen fra oppdraget:

1. Gyldig postnummer (slått opp i en kuratert referanse,
   `aga_lib/data/postnummer_referanse.json`, bygget og verifisert manuelt for
   nettopp de postnumrene som forekommer i datagrunnlaget – IKKE et
   fullstendig postnummerregister).
2–4. Kommunenavn funnet i henholdsvis Bedrift, Prosjekt og Jobb-feltet (mot
   navnene i Kommunekatalog 2026, inkl. sammensatte samiske/norske navn som
   splittes på " - ").
5. Et lite, eksplisitt kuratert unntak for institusjoner der verken
   postnummer eller kommunenavn i teksten identifiserer riktig kommune
   (`aga_lib/data/manuelle_unntak.json`).
6. Manuell kontroll.

**Viktig presisering (lagt til etter observasjon i faktiske data):** metode
2–5 samler ALLE treff (fra Bedrift, Prosjekt, Jobb og det manuelle unntaket)
før konklusjon, i stedet for å returnere på første treff. Dette avdekket at
ett prosjekt ("Sonjatun Sykestue") har prosjektnavnet
"Sonjatun Sykestue Kvænangen kommune" i RecMan, mens institusjonen selv
fysisk ligger i Nordreisa kommune – et reelt motstridende signal som nå
korrekt flagges `MOTSTRIDENDE_OPPLYSNINGER` i stedet for at det ene eller det
andre velges automatisk.

Dersom et prosjekt har mer enn ett gyldig postnummer, avbrytes automatikken
helt (heller ikke navnesøk forsøkes) og raden får
`MÅ_KVALITETSSIKRES – prosjekt kan omfatte flere arbeidssteder`.

## 7. AGA-klassifisering – IKKE endelig AGA

`aga_lib/aga_classification.py` slår opp kommunens AGA-sone i Skatteetatens
Kommunekatalog 2026 (`source_archive/skatteetaten_kommunekatalog_2026.csv`).
Kommuner som er delt mellom flere soner (f.eks. Sunnfjord: "Gamle Førde" =
sone 1a, "Gamle Gaular, Jølster og Naustdal" = sone 2) løses via prosjektets
postnummer der dette er mulig; ellers settes sone/sats tomt med
`MÅ_KVALITETSSIKRES`.

**Alle** funn i denne rapporten er merket
`GEOGRAFISK ANALYSEGRUNNLAG - MÅ AVKLARES MED LØNN`, ikke `KILDEVERIFISERT` –
se `Regelverksgrunnlag`-arket for hvorfor (kort: det juridiske grunnlaget for
å bruke arbeidsstedets kommune som AGA-basis for et bemanningsforetak som
leier ut arbeidskraft, er ikke bekreftet mot primærkilde i denne kjøringen,
og skatteetaten.no/lovdata.no/regjeringen.no var ikke nåbare fra
kjøremiljøet – se nettverksprobe i Datakvalitet).

Dersom kommunekatalogen mangler lokalt OG ingen av de offisielle
kildedomenene er nåbare (`aga_lib/oppslagsbehov.py`), fylles IKKE AGA-sone/
-sats ut i det hele tatt. I stedet skrives `.\output\AGA_oppslagsbehov.xlsx`
med kontrollstatus `MANGLER_OFFISIELT_KILDEGRUNNLAG`, mens resten av
dataanalysen (prosjektregister, lønnsgrunnlag, kommuneidentifikasjon der
mulig) fortsatt gjennomføres.

## 8. Kildehierarki

Skatteetaten > Lovdata > Regjeringen.no > SSB/Kartverket (kun for offentlig
geografisk informasjon). Se `Kilderegister`-arket og
`documentation/Kildeoversikt.md`.

## 9. Behandling av tilleggslinjer

`aga_lib/lines.py` klassifiserer hver rad som **arbeidstid**, **tillegg**
eller **korreksjon**, styrt av tre lister i `config.json`
(`linjeklassifisering`): `arbeidstid_artikler`, `korreksjon_artikler`/
`korreksjon_artikkeltyper`, og `kjente_tillegg_artikler`. En artikkel som
ikke finnes i noen av de tre listene flagges `ukjent_klassifisering` (og
behandles som tillegg) – dette vises som en avviksrad slik at lønn kan
bekrefte at listene fortsatt er komplette neste lønnsperiode.

**Arbeidstimer** telles KUN for rader klassifisert som arbeidstid (f.eks.
Artikkel "Arbeidstimer") – tilleggslinjer for samme vakt (nattillegg,
kveldstillegg, helg, overtid osv.) legges ikke oppå.

**Viktig unntak oppdaget i faktiske data:** På enkelte helge-/helligdagsvakter
mangler RecMan en egen "Arbeidstimer"-rad helt - hele vaktens timer ligger i
stedet på artikkelen "Helg" og/eller "Helligdag 133,33%" (bekreftet: disse
radene har reelle verdier i `Timer`/`Timer ekskl. pause` når ingen
"Arbeidstimer"-rad finnes for samme ansattnr/arbeidsdato/jobbnr/prosjektnr).
Uten en korreksjon ville disse vaktenes arbeidstid feilaktig blitt talt som 0.
`aga_lib.lines.korriger_arbeidstimer_for_manglende_arbeidstidsrad` retter dette:
for hver vakt UTEN egen arbeidstidsrad brukes den STØRSTE enkeltverdien blant
kandidatene i `config.json` sin `fallback_arbeidstid_artikler` (ikke summen -
"Helg" og "Helligdag 133,33%" kan representere DEN SAMME vakten registrert på
to parallelle lønnsartikler). Berørte rader flagges i den nye kolonnen
`arbeidstid_fallback_brukt` og telles i Datakvalitet-arket - i denne kjøringen
gjaldt det 83 vakter.

## 9b. Én prosess, alle leveranser

`aga_analyse.kjoer_analyse()` leser og analyserer kildefilen KUN ÉN gang per
kjøring (`aga_analyse.beregn_grunnlag()`), og skriver deretter ALLE
resultatfilene fra det samme, delte grunnlaget - ingen fil krever et eget
programkall eller leser filen på nytt:

1. `AGA_Rapport_<år>.xlsx` - hovedrapporten med 15 ark.
2. Seks separate STEG4-9-filer (forenklet kolonneoppsett per oppdragets
   STEG 4-9): `Prosjektregister.xlsx`, `AGA-detaljrapport.xlsx`,
   `AGA-oppsummering.xlsx`, `Kilderegister.xlsx`, `Avviksrapport.xlsx` og
   `Lederoppsummering.md`. Se `aga_lib/leveranser_steg4_9.py`
   (`bygg_leveranse_data()` bygger dataene rent, `bygg_og_skriv_leveranser()`
   skriver dem til fil).
3. `AGA-analyse-<år>-alle-steg.xlsx` - alle ni steg samlet i én arbeidsbok
   (12 ark), se `aga_lib/felles_arbeidsbok.py`.

Dersom offisiell AGA-kilde mangler helt (jf. det kritiske kravet i STEG 4/9),
skrives `Kilderegister.xlsx`/STEG5-arket TOMT (kun kolonneoverskrifter) og
AGA-sone/-status settes til "MANGLER OFFISIELT KILDEGRUNNLAG" i de andre
filene, uten å gjette - se `AGA_oppslagsbehov.xlsx`.

Etter kjøringen skrives et tydelig sammendrag til konsoll/logg
(`aga_analyse._skriv_tydelig_sammendrag()`): nøkkeltall og en full liste over
alle filene som ble produsert, slik at man ikke trenger å åpne Excel for å se
hva analysen fant.

## 10. Konstruert vakt-ID

`aga_lib/lines.bygg_vakt_id` lager en teknisk ID fra
Ansattnr + Arbeidsdato + Jobbnr + Prosjektnr + Fra kl. + Til kl. Dette er
**ikke** en original RecMan-ID – den brukes kun til å telle unike vakter og
oppdage duplikater. Siden tilleggslinjer for samme vakt har samme
klokkeslett, får de automatisk samme vakt-ID som arbeidstid-linjen, uten at
det regnes som en egen vakt.

## 11. Økonomiske beregninger

- `Lønn_uten_sosial_kost` = kolonnen `Total lønn` (aldri blandet med
  `Total lønn inkl. sosial kost`).
- `Sosial_kost` = kolonnen `Sosial kost`.
- `Lønn_inkl_sosial_kost` = `Total lønn inkl. sosial kost` (faller tilbake til
  `Lønn_uten_sosial_kost + Sosial_kost` dersom feltet mangler).
- **Foreløpig AGA-beløp** = `Lønn_uten_sosial_kost × ordinær sats for sonen`,
  kun beregnet der sonen er kjent. Alltid merket
  "FORELØPIG BEREGNING – MÅ KONTROLLERES AV LØNN".
- Sammenligningssatsen (for differanseberegningen i AGA_per_*-arkene) hentes
  fra `config.json` (`sammenligningssats_prosent`) – aldri hardkodet i koden.
- Alle beløp beregnes med full flyttallspresisjon og avrundes kun ved
  presentasjon i Excel (`#,##0.00`-format).

## 12. Usikkerheter og begrensninger

- Nettverkstilgang til skatteetaten.no, lovdata.no, regjeringen.no, ssb.no og
  kartverket.no var ikke tilgjengelig fra kjøremiljøet da denne rapporten ble
  generert (bekreftet av `aga_lib/sources.test_nettverkstilgang` – se
  Datakvalitet-arket for detaljer per domene).
- Det er ikke bekreftet mot primærkilde om arbeidsstedets kommune (i
  motsetning til f.eks. registrert underenhet) er riktig AGA-basis for et
  bemanningsforetak som leier ut helsepersonell. Se `Regelverksgrunnlag`.
- Postnummer-referansen dekker kun postnumrene observert i datagrunnlaget på
  analysetidspunktet – et NYTT postnummer i en fremtidig kjøring vil IKKE bli
  gjettet, men vil kreve manuell kontroll eller utvidelse av referansefilen.

## 13. Manuelle kontrollpunkter

Gå gjennom, i prioritert rekkefølge:

1. Alle rader i **Avvik**.
2. Alle rader i **Kommuneoppslag**/**Prosjektregister** uten status
   `KOMMUNE_VERIFISERT`.
3. **Regelverksgrunnlag**-arket, særlig spørsmålet om utleie av arbeidskraft.
4. Eventuell **AGA_oppslagsbehov.xlsx** dersom offisiell kilde manglet ved
   kjøretidspunktet.

## 14. Gjenkjøring neste lønnsperiode

1. Oppdater `config.json` (`analyseaar` dersom det endres).
2. Legg den nye RecMan-eksporten i `.\input`.
3. Oppdater `source_archive\skatteetaten_kommunekatalog_2026.csv` og
   `source_archive\skatteetaten_satser_2026_brukeroppgitt.txt` dersom det er
   et nytt år, eller dersom Skatteetaten har publisert en oppdatert
   sone-/satstabell for samme år.
4. Kjør `.\run_aga_analyse.ps1` fra prosjektmappen (eller
   `python aga_analyse.py` direkte i et aktivert virtuelt miljø).
5. Kontroller `output\AGA_Rapport_<år>.xlsx` – spesielt Avvik, Datakvalitet og
   Regelverksgrunnlag – før tallene brukes videre av lønn.
