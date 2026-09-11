"""Normalisering og validering av norske postnummer.

Postnummer skal alltid bevares som tekst med fire sifre - ledende null skal
aldri fjernes (0469 skal forbli "0469", ikke bli til 469 eller 469.0).
"""
from __future__ import annotations

import re

GYLDIG_POSTNUMMER_MONSTER = re.compile(r"^\d{4}$")


def normaliser_postnummer(verdi) -> str | None:
    """Returnerer et 4-sifret postnummer som tekst, eller None hvis verdien
    ikke lar seg tolke som et gyldig postnummer. Håndterer at Excel/pandas
    kan ha lest verdien som tall (469, 469.0) eller som tekst med eller uten
    ledende null."""
    if verdi is None:
        return None
    if isinstance(verdi, float):
        if verdi != verdi:  # NaN
            return None
        verdi = int(verdi)
    if isinstance(verdi, int):
        tekst = str(verdi)
    else:
        tekst = str(verdi).strip()
    if not tekst:
        return None
    tekst = tekst.replace(" ", "")
    if not tekst.isdigit():
        return None
    tekst = tekst.zfill(4)
    if len(tekst) != 4:
        return None
    return tekst


def er_gyldig_postnummer(postnummer) -> bool:
    """Robust mot at pandas (str-dtype-serier) kan levere manglende verdier
    som NaN (float) i stedet for None ved rundtrip gjennom en Series."""
    if not isinstance(postnummer, str):
        return False
    return bool(GYLDIG_POSTNUMMER_MONSTER.match(postnummer))
