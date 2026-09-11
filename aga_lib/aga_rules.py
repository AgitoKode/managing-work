"""Innlasting av Skatteetatens Kommunekatalog 2026 og satstabell for AGA 2026.

Begge datasettene er hentet fra Skatteetatens side "Satser for
arbeidsgiveravgift - soneinndeling" (kommunekatalog-2026-seksjonen) og lagt i
.\\source_archive av oppdragsgiver, siden dette kjøremiljøet ikke har utgående
nettverkstilgang til skatteetaten.no (se aga_lib.sources for nettverksprobe og
dokumentasjon av kildekjeden). Se documentation/Kildeoversikt.md for full
sporbarhet.
"""
from __future__ import annotations

import csv
import io
from dataclasses import dataclass
from pathlib import Path

# Satstabell for arbeidsgiveravgift 2026 (Stortingsvedtak, gjengitt av
# Skatteetaten). "ordinaer_sats" er satsen som gjelder for ordinære
# næringer (inkl. helse-/omsorgstjenester) - IKKE landbruk/fiske-satsen.
SATSTABELL_2026: dict[str, dict[str, object]] = {
    "1": {"ordinaer_sats": 0.141, "landbruk_fiske_sats": 0.141, "fribeloep": None},
    "1a": {
        "ordinaer_sats": 0.141,
        "landbruk_fiske_sats": 0.106,
        "fribeloep": 850_000,
        "merknad": (
            "Sone 1a: 10,6 % inntil differansen mot 14,1 %-satsen når fribeløpet "
            "(kr 850 000 per foretak i 2026), deretter 14,1 % på det overskytende "
            "avgiftsgrunnlaget."
        ),
    },
    "2": {"ordinaer_sats": 0.106, "landbruk_fiske_sats": 0.106, "fribeloep": None},
    "3": {"ordinaer_sats": 0.064, "landbruk_fiske_sats": 0.064, "fribeloep": None},
    "4": {"ordinaer_sats": 0.051, "landbruk_fiske_sats": 0.051, "fribeloep": None},
    "4a": {"ordinaer_sats": 0.079, "landbruk_fiske_sats": 0.051, "fribeloep": None},
    "5": {"ordinaer_sats": 0.0, "landbruk_fiske_sats": 0.0, "fribeloep": None},
}

SATSTABELL_KILDE_AAR = 2026


@dataclass
class SoneRad:
    kommunenummer: str
    kommunenavn: str
    fylke: str
    sone: str
    kommentar: str


def _reparerer_feilkvotert_rad(felt: list[str]) -> list[str] | None:
    """Kommunekatalog-CSV-en har én kjent feilkvotert rad der hele raden ved
    en feil ble pakket inn som ett enkelt CSV-felt (dobbel koding). Denne
    funksjonen forsøker å tolke et slikt enkeltfelt på nytt som en egen
    CSV-rad. Returnerer None dersom det ikke lar seg reparere, slik at
    kallende kode kan flagge raden i stedet for å gjette."""
    if len(felt) != 1:
        return None
    try:
        reparert = next(csv.reader(io.StringIO(felt[0])))
    except csv.Error:
        return None
    if len(reparert) in (4, 5):
        return reparert
    return None


def les_kommunekatalog(csv_sti: Path) -> tuple[list[SoneRad], list[str]]:
    """Leser Skatteetatens Kommunekatalog 2026. Returnerer (rader, advarsler).

    En kommune kan forekomme på flere rader dersom den er delt mellom flere
    AGA-soner (se Kommentar-feltet, f.eks. "Gamle Førde" for Sunnfjord)."""
    advarsler: list[str] = []
    rader: list[SoneRad] = []
    with csv_sti.open("r", encoding="utf-8-sig", newline="") as f:
        leser = csv.reader(f)
        header = next(leser)
        for linjenr, felt in enumerate(leser, start=2):
            if len(felt) not in (4, 5):
                reparert = _reparerer_feilkvotert_rad(felt)
                if reparert is None:
                    advarsler.append(
                        f"Linje {linjenr} i kommunekatalogen kunne ikke tolkes "
                        f"({len(felt)} felt funnet) og ble hoppet over: {felt!r}"
                    )
                    continue
                advarsler.append(
                    f"Linje {linjenr} i kommunekatalogen var feilkvotert i kildefilen "
                    "(hele raden var pakket inn som ett CSV-felt) og ble automatisk reparert."
                )
                felt = reparert
            kommentar = felt[4] if len(felt) == 5 else ""
            rader.append(
                SoneRad(
                    kommunenummer=felt[0].strip(),
                    kommunenavn=felt[1].strip(),
                    fylke=felt[2].strip(),
                    sone=felt[3].strip(),
                    kommentar=kommentar.strip(),
                )
            )
    return rader, advarsler


class Sonekatalog:
    """Oppslag av AGA-sone(r) per kommunenummer, med støtte for kommuner som
    er delt mellom flere soner."""

    def __init__(self, rader: list[SoneRad]):
        self.rader = rader
        self._per_kommunenummer: dict[str, list[SoneRad]] = {}
        for rad in rader:
            self._per_kommunenummer.setdefault(rad.kommunenummer, []).append(rad)

    def soner_for_kommune(self, kommunenummer: str) -> list[SoneRad]:
        return self._per_kommunenummer.get(kommunenummer, [])

    def er_delt_kommune(self, kommunenummer: str) -> bool:
        return len(self.soner_for_kommune(kommunenummer)) > 1

    def sats_for_sone(self, sone: str, sektor: str = "ordinaer") -> float | None:
        rad = SATSTABELL_2026.get(sone)
        if rad is None:
            return None
        noekkel = "ordinaer_sats" if sektor == "ordinaer" else "landbruk_fiske_sats"
        return rad[noekkel]
