"""Geografisk AGA-klassifisering per prosjekt/arbeidssted - IKKE endelig AGA.

VIKTIG (jf. oppdragets seksjon 8 og 15): Denne modulen slår opp hvilken sone
Skatteetatens Kommunekatalog 2026 knytter til arbeidsstedets kommune (og der
det er mulig, hvilken del av en delt kommune). Den avgjør IKKE om
prosjektkommunen faktisk er riktig juridisk grunnlag for AGA-sone for et
bemanningsforetak som leier ut arbeidskraft - det spørsmålet er ikke avklart
mot primærkilde i denne kjøringen (se Regelverksgrunnlag-arket og
aga_lib/sources.py). Derfor får ingen rad status KILDEVERIFISERT her; i
stedet brukes betegnelsen GEOGRAFISK_ANALYSEGRUNNLAG for geografisk
resonnerte funn, og MÅ_KVALITETSSIKRES der selve sonen ikke lar seg fastslå.
"""
from __future__ import annotations

from dataclasses import dataclass

from aga_lib.aga_rules import SATSTABELL_KILDE_AAR, Sonekatalog

KONTROLLSTATUS_GEOGRAFISK_GRUNNLAG = "GEOGRAFISK ANALYSEGRUNNLAG - MÅ AVKLARES MED LØNN"
KONTROLLSTATUS_MAA_KVALITETSSIKRES = "MÅ_KVALITETSSIKRES"


@dataclass
class AgaKlassifisering:
    kommunenummer: str | None
    kommune: str | None
    sone: str
    ordinaer_sats: float | None
    saerregler: str
    geografisk_presisering: str
    kilde_id: str
    kildeaar: int
    kontrollstatus: str
    kontrollkommentar: str


def klassifiser(
    sonekatalog: Sonekatalog,
    kommunenummer: str | None,
    kommune: str | None,
    postnummer: str | None,
    delkommune_merknad: str,
    kilde_id: str,
) -> AgaKlassifisering:
    if kommunenummer is None:
        return AgaKlassifisering(
            kommunenummer=None,
            kommune=kommune,
            sone="",
            ordinaer_sats=None,
            saerregler="",
            geografisk_presisering="",
            kilde_id=kilde_id,
            kildeaar=SATSTABELL_KILDE_AAR,
            kontrollstatus=KONTROLLSTATUS_MAA_KVALITETSSIKRES,
            kontrollkommentar="Ingen kommune er identifisert for arbeidsstedet, kan derfor ikke sonebestemmes.",
        )

    soner = sonekatalog.soner_for_kommune(kommunenummer)
    if not soner:
        return AgaKlassifisering(
            kommunenummer=kommunenummer,
            kommune=kommune,
            sone="",
            ordinaer_sats=None,
            saerregler="",
            geografisk_presisering="",
            kilde_id=kilde_id,
            kildeaar=SATSTABELL_KILDE_AAR,
            kontrollstatus=KONTROLLSTATUS_MAA_KVALITETSSIKRES,
            kontrollkommentar=(
                f"Kommunenummer {kommunenummer} finnes ikke i Kommunekatalog 2026. "
                "Kan skyldes en kommune-/nummerendring som ikke er reflektert i kildefilen."
            ),
        )

    if len(soner) == 1:
        rad = soner[0]
        sats = sonekatalog.sats_for_sone(rad.sone)
        saerregel = _sats_merknad(rad.sone)
        return AgaKlassifisering(
            kommunenummer=kommunenummer,
            kommune=rad.kommunenavn,
            sone=rad.sone,
            ordinaer_sats=sats,
            saerregler=saerregel,
            geografisk_presisering=rad.kommentar,
            kilde_id=kilde_id,
            kildeaar=SATSTABELL_KILDE_AAR,
            kontrollstatus=KONTROLLSTATUS_GEOGRAFISK_GRUNNLAG,
            kontrollkommentar="Hele kommunen tilhører én AGA-sone i Kommunekatalog 2026.",
        )

    # Kommunen er delt mellom flere soner - krev postnummerbasert presisering.
    if postnummer:
        for rad in soner:
            if delkommune_merknad and delkommune_merknad.lower() in rad.kommentar.lower():
                sats = sonekatalog.sats_for_sone(rad.sone)
                saerregel = _sats_merknad(rad.sone)
                return AgaKlassifisering(
                    kommunenummer=kommunenummer,
                    kommune=rad.kommunenavn,
                    sone=rad.sone,
                    ordinaer_sats=sats,
                    saerregler=saerregel,
                    geografisk_presisering=rad.kommentar,
                    kilde_id=kilde_id,
                    kildeaar=SATSTABELL_KILDE_AAR,
                    kontrollstatus=KONTROLLSTATUS_GEOGRAFISK_GRUNNLAG,
                    kontrollkommentar=(
                        f"Kommunen er delt mellom flere soner ({[r.sone for r in soner]}). "
                        f"Løst via postnummer {postnummer}, som er tilordnet delområdet "
                        f"'{rad.kommentar}'."
                    ),
                )

    alle_soner = ", ".join(sorted({r.sone for r in soner}))
    return AgaKlassifisering(
        kommunenummer=kommunenummer,
        kommune=kommune,
        sone="",
        ordinaer_sats=None,
        saerregler="",
        geografisk_presisering="; ".join(f"{r.sone}: {r.kommentar}" for r in soner if r.kommentar),
        kilde_id=kilde_id,
        kildeaar=SATSTABELL_KILDE_AAR,
        kontrollstatus=KONTROLLSTATUS_MAA_KVALITETSSIKRES,
        kontrollkommentar=(
            f"Kommune eller område kan omfattes av flere AGA-soner ({alle_soner}), og "
            "arbeidsstedet lar seg ikke presist tilordne en av dem ut fra tilgjengelig "
            "postnummer/geografisk informasjon i datagrunnlaget."
        ),
    )


def _sats_merknad(sone: str) -> str:
    from aga_lib.aga_rules import SATSTABELL_2026

    return SATSTABELL_2026.get(sone, {}).get("merknad", "")
