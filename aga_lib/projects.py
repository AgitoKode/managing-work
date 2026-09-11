"""Prosjektregister: én rad per unikt prosjekt/oppdrag, med avvikskontroller.

Prosjektnr er primær koblingsnøkkel. Jobbnr brukes som sekundær nøkkel når
Prosjektnr mangler. Prosjektnavn brukes aldri alene som nøkkel når et nummer
finnes.
"""
from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from aga_lib.municipality import (
    KONTROLLSTATUS_MAA_KVALITETSSIKRES,
    KommuneOppslag,
    KommuneResultat,
)
from aga_lib.postnummer import er_gyldig_postnummer


def _tekst_eller_none(verdi) -> str | None:
    """Rensker en celleverdi til en tekststreng, eller None dersom verdien
    mangler. pandas kan levere manglende verdier som NaN (float) i stedet for
    None etter en rundtur gjennom en "str"-dtype-kolonne, så vi sjekker
    eksplisitt for float/NaN i tillegg til None/tom streng."""
    if verdi is None:
        return None
    if isinstance(verdi, float) and verdi != verdi:  # NaN
        return None
    tekst = str(verdi).strip()
    return tekst if tekst and tekst not in ("nan", "None") else None


def _prosjektnokkel(rad) -> tuple[str, str] | None:
    prosjektnr = _tekst_eller_none(rad.get("prosjektnr"))
    if prosjektnr is not None:
        return ("prosjektnr", prosjektnr)
    jobbnr = _tekst_eller_none(rad.get("jobbnr"))
    if jobbnr is not None:
        return ("jobbnr", jobbnr)
    return None


@dataclass
class ProsjektregisterRad:
    noekkeltype: str
    noekkelverdi: str
    prosjektnr: str
    jobbnr: str
    prosjektnavn: list[str]
    jobbnavn: list[str]
    bedriftnavn: list[str]
    org_nr: list[str]
    postnumre: list[str]
    avdelinger: list[str]
    antall_rader: int
    flere_postnumre: bool
    flere_kommuner_mulig: bool
    samme_nr_flere_navn: bool
    samme_navn_flere_nr: bool
    postnummer_mangler_eller_ugyldig: bool
    kommune_resultat: KommuneResultat


def bygg_prosjektregister(df: pd.DataFrame, kommuneoppslag: KommuneOppslag) -> list[ProsjektregisterRad]:
    grupper: dict[tuple[str, str], list[dict]] = {}
    for _, rad in df.iterrows():
        noekkel = _prosjektnokkel(rad)
        if noekkel is None:
            noekkel = ("ukjent", f"UKJENT-{rad.get('bedrift', 'ukjent bedrift')}")
        grupper.setdefault(noekkel, []).append(rad.to_dict())

    # Kontroller "samme prosjektnavn har flere prosjektnumre": bygg et
    # navn -> sett av nøkler-indeks på tvers av alle grupper.
    navn_til_nokler: dict[str, set[tuple[str, str]]] = {}
    for noekkel, rader in grupper.items():
        for navn in {_tekst_eller_none(r.get("prosjekt")) for r in rader} - {None}:
            navn_til_nokler.setdefault(navn, set()).add(noekkel)

    register: list[ProsjektregisterRad] = []
    for (noekkeltype, noekkelverdi), rader in grupper.items():
        prosjektnr = next((_tekst_eller_none(r.get("prosjektnr")) for r in rader if _tekst_eller_none(r.get("prosjektnr"))), "")
        jobbnr = next((_tekst_eller_none(r.get("jobbnr")) for r in rader if _tekst_eller_none(r.get("jobbnr"))), "")
        prosjektnavn = sorted({_tekst_eller_none(r.get("prosjekt")) for r in rader} - {None})
        jobbnavn = sorted({_tekst_eller_none(r.get("jobb")) for r in rader} - {None})
        bedriftnavn = sorted({_tekst_eller_none(r.get("bedrift")) for r in rader} - {None})
        org_nr = sorted({_tekst_eller_none(r.get("org_nr")) for r in rader} - {None})
        avdelinger = sorted({_tekst_eller_none(r.get("avdeling")) for r in rader} - {None})

        # pandas' "str"-dtype-kolonner kan gi NaN (float) i stedet for None for
        # manglende verdier ved rundtrip via DataFrame - filtrer strengt på str.
        raw_postnumre = [r.get("postnummer_normalisert") for r in rader]
        tekst_postnumre = [p for p in raw_postnumre if isinstance(p, str) and p]
        gyldige_postnumre = sorted({p for p in tekst_postnumre if er_gyldig_postnummer(p)})
        alle_postnumre = sorted(set(tekst_postnumre))

        flere_postnumre = len(gyldige_postnumre) > 1
        postnummer_mangler_eller_ugyldig = len(gyldige_postnumre) == 0
        samme_nr_flere_navn = len(prosjektnavn) > 1
        samme_navn_flere_nr = any(len(navn_til_nokler.get(navn, set())) > 1 for navn in prosjektnavn)

        if flere_postnumre:
            kommune_resultat = KommuneResultat(
                kommune=None,
                kommunenummer=None,
                benyttet_postnummer=", ".join(gyldige_postnumre),
                identifikasjonsmetode="Avbrutt - flere postnumre på samme prosjekt",
                identifikasjonskilde="RecMan-datagrunnlag",
                kontrollstatus=KONTROLLSTATUS_MAA_KVALITETSSIKRES,
                kontrollkommentar="MÅ_KVALITETSSIKRES – prosjekt kan omfatte flere arbeidssteder",
            )
            flere_kommuner_mulig = True
        else:
            postnummer = gyldige_postnumre[0] if gyldige_postnumre else None
            kommune_resultat = kommuneoppslag.identifiser(
                postnummer_raw=postnummer,
                bedrift="; ".join(bedriftnavn),
                prosjekt="; ".join(prosjektnavn),
                jobb="; ".join(jobbnavn),
            )
            flere_kommuner_mulig = kommune_resultat.kontrollstatus == "MOTSTRIDENDE_OPPLYSNINGER"

        register.append(
            ProsjektregisterRad(
                noekkeltype=noekkeltype,
                noekkelverdi=noekkelverdi,
                prosjektnr=prosjektnr,
                jobbnr=jobbnr,
                prosjektnavn=prosjektnavn,
                jobbnavn=jobbnavn,
                bedriftnavn=bedriftnavn,
                org_nr=org_nr,
                postnumre=alle_postnumre,
                avdelinger=avdelinger,
                antall_rader=len(rader),
                flere_postnumre=flere_postnumre,
                flere_kommuner_mulig=flere_kommuner_mulig,
                samme_nr_flere_navn=samme_nr_flere_navn,
                samme_navn_flere_nr=samme_navn_flere_nr,
                postnummer_mangler_eller_ugyldig=postnummer_mangler_eller_ugyldig,
                kommune_resultat=kommune_resultat,
            )
        )

    return register
