# Kildeoversikt

Fullstendige, strukturerte kilderader ligger i arket **Kilderegister** i
`output\AGA_Rapport_2026.xlsx` (kilde-ID, kildetittel, direkte URL, offentlig
utgiver, dokumenttype, relevant år, kontrollert tidspunkt, sitat/henvisning,
kontrollstatus og merknad). Denne filen oppsummerer kildehierarkiet og de
viktigste forbeholdene.

## Kildehierarki (prioritert)

1. **Skatteetaten** – primær kilde for AGA-sone og -sats.
2. **Lovdata** – forskrift/Stortingsvedtak som fastsetter satsene.
3. **Regjeringen.no** – brukes der relevant informasjon ikke finnes hos
   Skatteetaten.
4. **SSB / Kartverket** – kun for offentlig geografisk informasjon
   (kommunenummer, kommunegrenser), ikke for selve AGA-regelverket.

## Kilder brukt i denne kjøringen

| Kilde-ID | Hva den underbygger | Status og forbehold |
|---|---|---|
| `KK-2026` | AGA-sone per kommunenummer for 2026, inkl. delte kommuner | Skatteetatens "Kommunekatalog 2026", mottatt som CSV fra oppdragsgiver 2026-09-11. `MÅ_KVALITETSSIKRES`: den levende kildesiden er ikke selv åpnet i denne kjøringen (nettverksblokkering, se under). |
| `SATS-2026` | Ordinær AGA-sats per sone og fribeløpsregelen for sone Ia | Samme kilde/side som KK-2026, satstabellen mottatt som tekst fra oppdragsgiver. Samme forbehold. |
| `PN-REF-2026` | Postnummer -> kommunenummer for postnumrene i datagrunnlaget | Bring/Posten Norge sitt offentlige postnummerregister, kuratert manuelt per postnummer (IKKE et fullstendig register). |
| `REGELVERK-AGA-2026` | Om arbeidsstedets kommune er riktig AGA-basis ved utleie av arbeidskraft | **Ikke funnet/åpnet** i denne kjøringen. Se Regelverksgrunnlag-arket. |

## Hvorfor ingen kilde har status KILDEVERIFISERT

Programmet tester nettverkstilgangen til de fem offisielle domenene ved hver
kjøring (`aga_lib/sources.test_nettverkstilgang`, seksjon 11 i oppdraget). I
kjøremiljøet denne rapporten ble generert i, svarte alle fem domenene med en
`ProxyError`/`403 Forbidden` fra miljøets nettverks-egress-policy – se
Datakvalitet-arket for feilmeldingen per domene og tidspunkt.

Siden ingen av kildene dermed er direkte åpnet og lest i *denne* kjøringen
(jf. kravet i seksjon 10: "Åpne den faktiske kilden og kontroller at
opplysningen finnes der"), og det juridiske spørsmålet i seksjon 8 heller
ikke er avklart, får ingen rad status `KILDEVERIFISERT` (seksjon 15). Alle
sone-/satsfunn er i stedet merket
`GEOGRAFISK ANALYSEGRUNNLAG - MÅ AVKLARES MED LØNN`.

## Anbefalt oppfølging

1. Kjør programmet på nytt fra et miljø med tilgang til skatteetaten.no,
   lovdata.no og regjeringen.no, slik at kildesidene faktisk kan åpnes og
   sitatene i Kilderegister kan verifiseres ord for ord.
2. Få en jurist/lønnsansvarlig til å avklare spørsmålene i
   `Regelverksgrunnlag`-arket, særlig om reglene for utleie av arbeidskraft.
3. Oppdater `source_archive`-filene når Skatteetaten publiserer endelige
   satser/soner for et nytt år, og noter publiserings-/oppdateringsdato i
   Kilderegister (feltet er tomt i denne kjøringen fordi det ikke fremkom i
   det mottatte uttrekket).
