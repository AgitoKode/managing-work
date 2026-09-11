"""Datakvalitet- og avvikskontroller (seksjon 15).

Returnerer to lister av ordbøker (rader) klare til å skrives til Excel-arkene
"Avvik" og "Datakvalitet". Hver kontroll er uavhengig og teller/lister opp
konkrete rader/prosjekter - det gjettes aldri på årsak.
"""
from __future__ import annotations

import pandas as pd

from aga_lib.sources import NettverksproveResultat


def kjoer_kontroller(
    df: pd.DataFrame,
    prosjektregister: list,
    aga_klassifiseringer: list,
    analyseaar: int,
    kildeaar_kommunekatalog: int,
    kildeaar_satstabell: int,
    nettverksprobe: list[NettverksproveResultat],
) -> tuple[list[dict], list[dict]]:
    avvik: list[dict] = []
    datakvalitet: list[dict] = []

    def legg_til_avvik(kategori: str, antall: int, beskrivelse: str, alvorlighet: str = "Middels") -> None:
        if antall > 0:
            avvik.append(
                {
                    "Kategori": kategori,
                    "Antall": antall,
                    "Alvorlighet": alvorlighet,
                    "Beskrivelse": beskrivelse,
                }
            )

    manglende_ansattnr = df["ansattnr"].isna().sum()
    legg_til_avvik("Manglende ansattnummer", int(manglende_ansattnr), "Rader uten Ansattnr.", "Høy")

    manglende_prosjekt_jobb = df.apply(
        lambda r: (pd.isna(r.get("prosjektnr")) or str(r.get("prosjektnr")) in ("", "nan"))
        and (pd.isna(r.get("jobbnr")) or str(r.get("jobbnr")) in ("", "nan")),
        axis=1,
    ).sum()
    legg_til_avvik(
        "Manglende prosjekt- og jobbnummer", int(manglende_prosjekt_jobb),
        "Rader uten både Prosjektnr. og Jobbnr. - kan ikke kobles til prosjektregisteret.", "Høy",
    )

    manglende_dato = df["arbeidsdato"].isna().sum()
    legg_til_avvik("Manglende arbeidsdato", int(manglende_dato), "Rader uten tolkbar arbeidsdato.", "Høy")

    ugyldige_postnumre = (~df["postnummer_normalisert"].isna() & ~df["postnummer_gyldig"]).sum()
    legg_til_avvik(
        "Ugyldige postnumre", int(ugyldige_postnumre),
        "Postnummerfeltet inneholder en verdi som ikke er et gyldig 4-sifret postnummer.",
    )

    linjer_uten_belop = (
        df["total_lonn"].isna()
        & df["loenn_paa_loennsgrunnlag"].isna()
        & df["total_lonn_inkl_sosial_kost"].isna()
    ).sum()
    legg_til_avvik(
        "Lønnslinjer uten beløp", int(linjer_uten_belop),
        "Ingen av feltene Total lønn / Lønn på lønnsgrunnlag / Total lønn inkl. sosial kost har verdi.",
    )

    negative_belop = (pd.to_numeric(df["total_lonn"], errors="coerce") < 0).sum()
    legg_til_avvik(
        "Negative lønnsbeløp (korreksjoner beholdt)", int(negative_belop),
        "Rader med negativt beløp i Total lønn. Beholdt i grunnlaget som korreksjon/fratrekk, ikke slettet.",
        "Lav",
    )

    ukjent_klassifisering = df.get("ukjent_klassifisering", pd.Series(dtype=bool)).sum()
    legg_til_avvik(
        "Artikkel med ukjent linjeklassifisering", int(ukjent_klassifisering),
        "Artikkelverdi som ikke er registrert i config.json sine lister over arbeidstid-/korreksjonsartikler "
        "(behandlet som tilleggslinje). Bør gjennomgås ved neste periode.",
        "Lav",
    )

    duplikater = df.duplicated(subset=["vakt_id", "artikkel", "total_lonn"], keep=False).sum()
    legg_til_avvik(
        "Mulige dupliserte linjer", int(duplikater),
        "Rader med identisk vakt-ID, Artikkel og Total lønn - kan være reelle dubletter eller "
        "legitime like tilleggslinjer på samme vakt. Krever manuell vurdering.",
    )

    manglende_kommune = sum(1 for p in prosjektregister if p.kommune_resultat.kommunenummer is None)
    legg_til_avvik(
        "Prosjekter uten identifisert kommune", manglende_kommune,
        "Prosjekter der kommune ikke kunne fastslås automatisk (se Kommuneoppslag-arket).", "Høy",
    )

    kun_fritekst = sum(
        1 for p in prosjektregister if p.kommune_resultat.kontrollstatus == "KOMMUNE_SANNSYNLIG"
    )
    legg_til_avvik(
        "Kommune identifisert kun fra fritekst", kun_fritekst,
        "Kommune er utledet fra navn i Bedrift/Prosjekt/Jobb-feltet (ikke fra postnummer) - "
        "bør kvalitetssikres, se advarsel i seksjon 7 om at kundenavn ikke garanterer arbeidssted.",
    )

    flere_postnumre = sum(1 for p in prosjektregister if p.flere_postnumre)
    legg_til_avvik(
        "Prosjekter koblet til flere postnumre", flere_postnumre,
        "Prosjektet har mer enn ett gyldig postnummer i datagrunnlaget - kan ikke automatisk "
        "tilordnes én kommune/sone.", "Høy",
    )

    flere_kommuner = sum(1 for p in prosjektregister if p.flere_kommuner_mulig)
    legg_til_avvik(
        "Prosjekter med motstridende/uklar kommunekobling", flere_kommuner,
        "Prosjektet har flertydige eller motstridende signaler om hvilken kommune arbeidet gjelder.", "Høy",
    )

    uten_sone = sum(1 for a in aga_klassifiseringer if not a.sone)
    legg_til_avvik(
        "Prosjekter uten avklart AGA-sone", uten_sone,
        "AGA-sone er tom fordi kommunen ikke er identifisert, kommunenummeret ikke finnes i "
        "Kommunekatalogen, eller kommunen er delt mellom flere soner uten at riktig del kunne fastslås.",
        "Høy",
    )

    delt_sone = sum(
        1 for a in aga_klassifiseringer
        if a.kommunenummer and "delt mellom flere soner" in a.kontrollkommentar
        or (a.kommunenummer and not a.sone and "flere AGA-soner" in a.kontrollkommentar)
    )
    legg_til_avvik(
        "Prosjekter i kommuner med mulig delt AGA-sone", delt_sone,
        "Kommunen forekommer på flere rader i Kommunekatalog 2026 (ulik sone for ulike deler av kommunen).",
    )

    samme_navn_flere_nr = sum(1 for p in prosjektregister if p.samme_navn_flere_nr)
    legg_til_avvik(
        "Samme prosjektnavn har flere prosjekt-/jobbnumre", samme_navn_flere_nr,
        "Kan være separate oppdrag med likt navn, eller en navnedublett - krever manuell vurdering.", "Lav",
    )

    samme_nr_flere_navn = sum(1 for p in prosjektregister if p.samme_nr_flere_navn)
    legg_til_avvik(
        "Samme prosjektnummer har flere prosjektnavn", samme_nr_flere_navn,
        "Kan skyldes at prosjektet er omdøpt i perioden - bør bekreftes at det er samme oppdrag.", "Lav",
    )

    if analyseaar != kildeaar_kommunekatalog:
        avvik.append(
            {
                "Kategori": "AGA-kilde gjelder ikke analyseåret (kommunekatalog)",
                "Antall": 1,
                "Alvorlighet": "Høy",
                "Beskrivelse": (
                    f"Kommunekatalogen er for {kildeaar_kommunekatalog}, men analyseåret er {analyseaar}."
                ),
            }
        )
    if analyseaar != kildeaar_satstabell:
        avvik.append(
            {
                "Kategori": "AGA-kilde gjelder ikke analyseåret (satstabell)",
                "Antall": 1,
                "Alvorlighet": "Høy",
                "Beskrivelse": f"Satstabellen er for {kildeaar_satstabell}, men analyseåret er {analyseaar}.",
            }
        )

    for probe in nettverksprobe:
        if not probe.naaadd:
            avvik.append(
                {
                    "Kategori": "Nettverks-/kildefeil",
                    "Antall": 1,
                    "Alvorlighet": "Høy",
                    "Beskrivelse": (
                        f"Fikk ikke koblet til {probe.vert} for å verifisere kilden direkte "
                        f"i denne kjøringen ({probe.detaljer})."
                    ),
                }
            )

    avvik.append(
        {
            "Kategori": "Regelverksgrunnlag for geografisk AGA-basis er ikke bekreftet",
            "Antall": 1,
            "Alvorlighet": "Høy",
            "Beskrivelse": (
                "Det er ikke bekreftet mot primærkilde i denne kjøringen at prosjektets arbeidskommune er "
                "riktig juridisk grunnlag for AGA-sone for et bemanningsforetak (se Regelverksgrunnlag-arket). "
                "Derfor har INGEN rad i denne rapporten status KILDEVERIFISERT - alle geografiske "
                "sonefunn er merket 'GEOGRAFISK ANALYSEGRUNNLAG - MÅ AVKLARES MED LØNN'."
            ),
        }
    )

    # --- Datakvalitet (tellende oversikt, ikke nødvendigvis avvik) ---
    datakvalitet = [
        {"Kontroll": "Totalt antall innleste rader", "Verdi": len(df)},
        {"Kontroll": "Rader med gyldig postnummer", "Verdi": int(df["postnummer_gyldig"].sum())},
        {"Kontroll": "Rader klassifisert som arbeidstid", "Verdi": int((df["lonnskomponent"] == "arbeidstid").sum())},
        {"Kontroll": "Rader klassifisert som tillegg", "Verdi": int((df["lonnskomponent"] == "tillegg").sum())},
        {"Kontroll": "Rader klassifisert som korreksjon", "Verdi": int((df["lonnskomponent"] == "korreksjon").sum())},
        {"Kontroll": "Unike vakt-ID-er", "Verdi": int(df["vakt_id"].nunique())},
        {"Kontroll": "Unike ansatte", "Verdi": int(df["ansattnr"].nunique())},
        {"Kontroll": "Unike prosjekter i prosjektregisteret", "Verdi": len(prosjektregister)},
    ]

    return avvik, datakvalitet
