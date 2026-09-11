"""Kilderegister, nettverksprobe og SHA-256-arkivering av kildefiler.

Kun geografiske/offentlige data (kommunenavn, kommunenummer, postnummer,
domenenavn) sendes i nettverksforespørsler herfra - aldri ansattnavn,
e-postadresser, lønnsopplysninger eller arbeidsposter (jf. oppdragets
personvernkrav i seksjon 2).
"""
from __future__ import annotations

import datetime as dt
import hashlib
from dataclasses import dataclass
from pathlib import Path

import requests


@dataclass
class KildeRad:
    kilde_id: str
    kildetittel: str
    kildeadresse: str
    offentlig_utgiver: str
    dokumenttype: str
    relevant_aar: int
    publisert_dato: str
    sist_oppdatert_dato: str
    kontrollert_tidspunkt: str
    underbygger: str
    sitat_henvisning: str
    kontrollstatus: str
    merknad: str
    sha256: str = ""


@dataclass
class NettverksproveResultat:
    vert: str
    naaadd: bool
    detaljer: str


def test_nettverkstilgang(offisielle_domener: list[str], timeout: float) -> list[NettverksproveResultat]:
    """Tester om de offisielle kildedomenene er nåbare FØR noe forsøkes hentet.
    Sender kun en enkel HTTPS-forespørsel til domenets forside - ingen
    persondata, ingen filinnhold, involveres i disse kallene."""
    resultater: list[NettverksproveResultat] = []
    for domene in offisielle_domener:
        url = f"https://www.{domene}" if not domene.startswith("www.") else f"https://{domene}"
        try:
            svar = requests.get(url, timeout=timeout)
            resultater.append(
                NettverksproveResultat(vert=domene, naaadd=svar.ok, detaljer=f"HTTP {svar.status_code}")
            )
        except requests.RequestException as e:
            resultater.append(
                NettverksproveResultat(vert=domene, naaadd=False, detaljer=f"{type(e).__name__}: {e}")
            )
    return resultater


def sha256_for_fil(sti: Path) -> str:
    return hashlib.sha256(sti.read_bytes()).hexdigest()


def naa_iso() -> str:
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
