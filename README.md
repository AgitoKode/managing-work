# AGA-analyse av Ført arbeid – utvidet

Python-løsning som leser RecMan-eksporten `Ført arbeid - utvidet.xlsx`,
kobler hver arbeidspost til kommune og dokumentert
arbeidsgiveravgiftssone (AGA), og bygger et kvalitetssikret
rapportgrunnlag for lønnsavdelingen: `output\AGA_Rapport_2026.xlsx`.

**Viktig:** AGA-sone-/satsfunnene i rapporten er dokumentert som
**"GEOGRAFISK ANALYSEGRUNNLAG - MÅ AVKLARES MED LØNN"**, ikke som endelig
korrekt AGA. Se arket `Regelverksgrunnlag` i rapporten og
`documentation\Kildeoversikt.md` for hvorfor.

## Forutsetninger

- Windows med PowerShell.
- Python 3 installert og tilgjengelig i `PATH` (`python` eller `py`).
- `Ført arbeid - utvidet.xlsx` liggende enten i prosjektmappen eller i
  `.\input`. (Filnavnet kan ha et suffiks, f.eks. en dato - løsningen finner
  filen så lenge navnet inneholder "Ført arbeid" og "utvidet".)

Du trenger ikke flytte filen manuelt dersom den allerede ligger i
prosjektmappen eller i `.\input`.

## Rask start (PowerShell)

Åpne PowerShell i prosjektmappen og kjør:

```powershell
.\run_aga_analyse.ps1
```

Skriptet vil:

1. Kontrollere at Python er tilgjengelig.
2. Opprette et virtuelt miljø i `.\.venv` dersom det ikke finnes.
3. Aktivere miljøet.
4. Installere pakkene i `requirements.txt`.
5. Kjøre `aga_analyse.py`.
6. Skrive ut hvor resultatfilen (`output\AGA_Rapport_2026.xlsx`) ble lagret.
7. Stoppe med en forståelig norsk feilmelding dersom noe går galt (f.eks.
   manglende kildefil, manglende Python, eller feil under installasjon).

## Manuell kjøring

```powershell
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python aga_analyse.py
```

## Mappestruktur

Løsningen oppretter disse mappene automatisk hvis de ikke finnes:

| Mappe | Innhold |
|---|---|
| `.\input` | Kildefilen(e) fra RecMan. |
| `.\output` | Genererte rapporter (`AGA_Rapport_2026.xlsx`, evt. `AGA_oppslagsbehov.xlsx`). |
| `.\cache` | Reservert for fremtidig mellomlagring (brukes ikke av dagens versjon). |
| `.\logs` | Kjørelogger (persondata er maskert, se Personvern under). |
| `.\documentation` | Denne filen, `Metodedokumentasjon.md` og `Kildeoversikt.md`. |
| `.\source_archive` | Offisielle kildefiler (Skatteetatens Kommunekatalog 2026 og satstabell) med SHA-256-sum. |

Alle stier er relative og bygget med `pathlib.Path`, og fungerer selv om
Windows-brukernavnet, prosjektmappen eller filnavnet inneholder mellomrom
eller bokstavene æ/ø/å.

## Konfigurasjon (`config.json`)

- `analyseaar`: hvilket år som analyseres (2026).
- `termin_maaneder`: hvilke måneder som tilhører hvilken termin.
- `sammenligningssats_prosent`: satsen AGA_per_*-arkene sammenligner den
  foreløpige beregningen mot (aldri hardkodet i Python-koden).
- `kildefil_sok`: søkerekkefølge og eksklusjonsmønstre for kildefilen.
- `linjeklassifisering`: hvilke Artikkel-/Artikkeltype-verdier som regnes som
  arbeidstid, korreksjon eller kjent tillegg (se
  `documentation\Metodedokumentasjon.md` punkt 9).

## Resultatfiler

Hver kjøring produserer, i `.\output`:

- `AGA_Rapport_2026.xlsx` - den samlede rapporten med 15 ark (se under).
- Seks frittstående STEG4-9-leveranser (samme datagrunnlag, andre kolonneoppsett):
  `Prosjektregister.xlsx`, `AGA-detaljrapport.xlsx`, `AGA-oppsummering.xlsx`,
  `Kilderegister.xlsx`, `Avviksrapport.xlsx`, `Lederoppsummering.md`.
- `AGA_oppslagsbehov.xlsx` - kun dersom offisiell AGA-kilde mangler helt (se under).

Ingen av disse overskrives stille - finnes filen fra før, lagres den nye med et
tidsstempel i filnavnet i stedet.

`output\AGA_Rapport_2026.xlsx` inneholder 15 ark: Lederoppsummering,
AGA_per_termin, AGA_per_kommune, AGA_per_prosjekt, AGA_per_ansatt,
Detaljgrunnlag, Prosjektregister, Kommuneoppslag, Kilderegister,
Regelverksgrunnlag, Avvik, Datakvalitet, Kolonnemapping, Metode og
Kjøremetadata.

Dersom en rapport med samme navn allerede finnes fra før, lagres den nye
rapporten med et tidsstempel i filnavnet i stedet for å overskrive - den
forrige rapporten røres aldri.

Dersom den offisielle AGA-kilden (Kommunekatalog/satstabell) mangler lokalt
OG de offisielle kildedomenene ikke er nåbare, fylles ikke AGA-sone/-sats ut.
I stedet skrives `output\AGA_oppslagsbehov.xlsx` med hva som må slås opp
manuelt, mens resten av analysen (prosjektregister, lønnsgrunnlag,
kommuneidentifikasjon der mulig) fortsatt gjennomføres.

## Personvern og sikkerhet

- Kildefilen inneholder ansattnavn og lønnsopplysninger. Den lastes aldri opp
  til eksterne tjenester, og eksterne oppslag (nettverksprobe mot de
  offisielle kildedomenene) sender kun domenenavn - aldri ansattdata.
- Loggfilene i `.\logs` maskerer automatisk e-postadresser og 11-sifrede tall
  (mulig fødselsnummer). Tekniske feilmeldinger som må referere en ansatt,
  bruker initialer (`aga_lib.logging_setup.masker_navn`), ikke fullt navn.
- `.gitignore` ekskluderer Excel-filer, `.venv`, `cache`, `logs`, `input`,
  `output` og `source_archive` fra git, slik at persondata og genererte
  rapporter aldri havner i et delt repository.

## Tester

```powershell
python -m unittest discover -s tests -v
```

Dekker blant annet: postnummer med ledende null, konvertering av
Excel-seriedatoer, at tilleggslinjer ikke dobler arbeidstimer, at negative
korreksjoner beholdes, unik prosjektidentifikator i prosjektregisteret, at
rader uten dokumentert AGA-kilde ikke får status `KILDEVERIFISERT`, at en
AGA-kilde for feil år avvises, at resultatfilen inneholder alle obligatoriske
ark, og at kildefilen ikke endres av en kjøring.

## Kjent begrensning i denne kjøringen

Dette prosjektet ble utviklet og verifisert i et Linux-basert kjøremiljø
(Claude Code på nett), ikke lokalt på Windows. All Python-logikk er testet
og kjørt reelt mot den faktiske kildefilen i dette miljøet, men selve
`run_aga_analyse.ps1`-skriptet har **ikke** kunnet kjøres i PowerShell her -
det er skrevet og gjennomgått etter beste evne for Windows/PowerShell 5.1+,
men bør verifiseres på en faktisk Windows-maskin ved første bruk.
