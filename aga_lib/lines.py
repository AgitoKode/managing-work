"""Klassifisering av lønnslinjer og konstruksjon av teknisk vakt-ID.

RecMan-eksporten inneholder flere rader per vakt (ordinære timer,
kveldstillegg, nattillegg, helgetillegg, overtid, etterbetaling,
korreksjon). Vi skiller mellom:

  A. Arbeidstimer         - fysisk arbeidstid (kun ordinære arbeidstime-rader)
  B. Lønnskomponent/tillegg - påslag som gjelder timer som allerede er talt i A
  C. Korreksjon/etterbetaling - rettelser, kan være negative

Den tekniske vakt-ID-en brukes KUN til å telle unike vakter og oppdage
duplikater - den er konstruert av oss og er ikke en original RecMan-ID.
"""
from __future__ import annotations

import hashlib

import pandas as pd


def _tekst(verdi) -> str:
    """Som str(verdi), men behandler None/NaN som tom streng i stedet for
    å skrive teksten "nan" (pandas kan gi NaN i stedet for None for
    manglende verdier i en "str"-dtype-kolonne)."""
    if verdi is None:
        return ""
    if isinstance(verdi, float) and verdi != verdi:
        return ""
    return str(verdi)


def klassifiser_linjer(df: pd.DataFrame, klassifiseringsregler: dict) -> pd.DataFrame:
    df = df.copy()
    arbeidstid_artikler = set(klassifiseringsregler["arbeidstid_artikler"])
    korreksjon_artikler = set(klassifiseringsregler["korreksjon_artikler"])
    korreksjon_artikkeltyper = set(klassifiseringsregler["korreksjon_artikkeltyper"])
    kjente_tillegg_artikler = set(klassifiseringsregler.get("kjente_tillegg_artikler", []))

    def klassifiser_rad(rad) -> str:
        artikkel = _tekst(rad.get("artikkel"))
        artikkeltype = _tekst(rad.get("artikkeltype"))
        if artikkel in korreksjon_artikler or artikkeltype in korreksjon_artikkeltyper:
            return "korreksjon"
        if artikkel in arbeidstid_artikler:
            return "arbeidstid"
        return "tillegg"

    df["lonnskomponent"] = df.apply(klassifiser_rad, axis=1)
    df["er_tilleggslinje"] = df["lonnskomponent"] == "tillegg"
    df["er_korreksjon"] = (df["lonnskomponent"] == "korreksjon") | (
        pd.to_numeric(df.get("total_lonn"), errors="coerce").fillna(0) < 0
    )

    kjente_artikler = arbeidstid_artikler | korreksjon_artikler | kjente_tillegg_artikler
    # "ukjent_klassifisering": artikkelen sto ikke i NOEN av de tre konfigurerte
    # listene i config.json (arbeidstid/korreksjon/kjente tillegg) og havnet
    # dermed i default-kategorien "tillegg" uten at det er bekreftet at det er
    # riktig. Flagges i Datakvalitet slik at lønn kan bekrefte at
    # klassifiseringen fortsatt er komplett neste lønnsperiode.
    df["ukjent_klassifisering"] = ~df["artikkel"].isin(kjente_artikler)
    return df


def arbeidstimer_for_rad(rad) -> float:
    """Arbeidstimer telles kun for rader klassifisert som 'arbeidstid', og da
    brukes primært 'Timer ekskl. pause' -> 'Timer' -> 'Timer inkl. pause'
    (i den prioriteringen), slik at vi ikke dobbeltteller pause som arbeidstid
    når et mer presist felt finnes."""
    if rad.get("lonnskomponent") != "arbeidstid":
        return 0.0
    for felt in ("timer_ekskl_pause", "timer", "timer_inkl_pause"):
        verdi = rad.get(felt)
        if verdi is not None and verdi == verdi:  # ikke NaN
            try:
                return float(verdi)
            except (TypeError, ValueError):
                continue
    return 0.0


def bygg_vakt_id(rad) -> str:
    """Teknisk, konstruert vakt-ID - IKKE en original RecMan-ID. Brukes kun til
    å telle unike vakter og oppdage duplikater, aldri som fasit-nøkkel mot
    RecMan."""
    deler = [
        str(rad.get("ansattnr", "")),
        str(rad.get("arbeidsdato", "")),
        str(rad.get("jobbnr", "")),
        str(rad.get("prosjektnr", "")),
        str(rad.get("fra_kl", "")),
        str(rad.get("til_kl", "")),
    ]
    grunnlag = "|".join(deler)
    return hashlib.sha1(grunnlag.encode("utf-8")).hexdigest()[:16]
