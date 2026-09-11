"""Oppsummeringer per ansatt, prosjekt, kunde, kommune, sone, måned, termin, år.

Alle beløp beregnes med full presisjon og avrundes kun i presentasjonslaget
(report.py). Sammenligningssatsen hentes fra config.json og er aldri
hardkodet her.
"""
from __future__ import annotations

import pandas as pd


def _sum_unike_vakter(gruppe: pd.DataFrame) -> int:
    return gruppe["vakt_id"].nunique()


def _foreloepig_aga(gruppe: pd.DataFrame, sammenligningssats: float) -> tuple[float, float]:
    grunnlag = gruppe["lonn_uten_sosial_kost"].sum()
    foreloepig_aga_sats = gruppe["ordinaer_sats"].dropna()
    if foreloepig_aga_sats.empty or gruppe["sone"].isna().all() or (gruppe["sone"] == "").all():
        beregnet = None
    else:
        # Vektet AGA-beregning: hver rad bruker sin egen sone-sats der den er kjent.
        rader_med_sats = gruppe.dropna(subset=["ordinaer_sats"])
        beregnet = (rader_med_sats["lonn_uten_sosial_kost"] * rader_med_sats["ordinaer_sats"]).sum()
    differanse = None
    if beregnet is not None:
        sammenligning = grunnlag * (sammenligningssats / 100.0)
        differanse = beregnet - sammenligning
    return beregnet, differanse


def bygg_oppsummering(df: pd.DataFrame, grupperingsfelt: list[str], sammenligningssats: float) -> pd.DataFrame:
    rader = []
    for noekkel, gruppe in df.groupby(grupperingsfelt, dropna=False):
        if not isinstance(noekkel, tuple):
            noekkel = (noekkel,)
        beregnet_aga, differanse = _foreloepig_aga(gruppe, sammenligningssats)
        rad = dict(zip(grupperingsfelt, noekkel))
        rad.update(
            {
                "antall_unike_ansatte": gruppe["ansattnr"].nunique(),
                "antall_unike_vakter": _sum_unike_vakter(gruppe),
                "arbeidstimer": gruppe["arbeidstimer"].sum(),
                "lonn_uten_sosial_kost": gruppe["lonn_uten_sosial_kost"].sum(),
                "sosial_kost": gruppe["sosial_kost"].sum(),
                "lonn_inkl_sosial_kost": gruppe["lonn_inkl_sosial_kost"].sum(),
                "foreloepig_aga_grunnlag": gruppe["lonn_uten_sosial_kost"].sum(),
                "foreloepig_beregnet_aga": beregnet_aga,
                "differanse_mot_sammenligningssats": differanse,
                "antall_rader": len(gruppe),
            }
        )
        rader.append(rad)
    return pd.DataFrame(rader)
