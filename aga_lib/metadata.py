"""Kjøremetadata og revisjonsspor (seksjon 20). Ingen sensitiv informasjon lagres her."""
from __future__ import annotations

import platform
import sys
from dataclasses import dataclass, field
from importlib.metadata import PackageNotFoundError, version


@dataclass
class Kjoremetadata:
    kjoredato_tidspunkt: str
    python_versjon: str
    kildefil_navn: str
    kildefil_sha256: str
    antall_innleste_rader: int
    antall_analyserte_rader: int
    antall_ekskluderte_rader: int
    antall_unike_ansatte: int
    antall_unike_prosjekter: int
    antall_identifiserte_kommuner: int
    antall_kildeverifiserte_kommune_sonekoblinger: int
    programversjon: str
    pakkeversjoner: dict[str, str] = field(default_factory=dict)


def hent_pakkeversjoner() -> dict[str, str]:
    pakker = ["pandas", "openpyxl", "requests", "beautifulsoup4", "lxml"]
    resultat = {}
    for pakke in pakker:
        try:
            resultat[pakke] = version(pakke)
        except PackageNotFoundError:
            resultat[pakke] = "ikke installert"
    return resultat


def python_versjon_streng() -> str:
    return f"{platform.python_implementation()} {sys.version.split()[0]}"
