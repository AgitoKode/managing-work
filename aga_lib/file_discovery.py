"""Finn kildefilen ("Ført arbeid - utvidet.xlsx") i prosjektmappen.

Søkerekkefølge (jf. oppdragets seksjon 3):
1. .\\input\\Ført arbeid - utvidet.xlsx (eksakt eller inneholder begge nøkkeltekster)
2. .\\Ført arbeid - utvidet.xlsx i prosjektets rotmappe
3. Andre .xlsx-filer i .\\input
4. Andre .xlsx-filer i prosjektets rotmappe

Filer hvis navn inneholder tidligere genererte resultatnavn
(AGA_Rapport, Prosjektregister, Avviksrapport, Kilderegister, Kontrollrapport,
AGA_oppslagsbehov) ekskluderes alltid, slik at løsningen aldri leser sin egen
forrige rapport som ny input.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


class KildefilIkkeFunnet(Exception):
    pass


@dataclass
class KildefilResultat:
    sti: Path
    begrunnelse: str
    kontrollerte_mapper: list[Path]


def _er_ekskludert(filnavn: str, ekskluder_tekster: list[str]) -> bool:
    lav = filnavn.lower()
    return any(tekst.lower() in lav for tekst in ekskluder_tekster)


def _matcher_noekkeltekster(filnavn: str, noekkeltekster: list[str]) -> bool:
    lav = filnavn.lower()
    return all(tekst.lower() in lav for tekst in noekkeltekster)


def finn_kildefil(
    prosjektmappe: Path,
    input_mappe: Path,
    noekkeltekster: list[str],
    ekskluder_tekster: list[str],
    tillatte_endelser: list[str],
) -> KildefilResultat:
    kontrollerte_mapper = [input_mappe, prosjektmappe]

    def kandidater(mappe: Path) -> list[Path]:
        if not mappe.exists():
            return []
        funn = []
        for p in sorted(mappe.iterdir()):
            if not p.is_file():
                continue
            if p.suffix.lower() not in tillatte_endelser:
                continue
            if p.name.startswith("~$"):  # Excel-låsefil
                continue
            if _er_ekskludert(p.name, ekskluder_tekster):
                continue
            funn.append(p)
        return funn

    # Steg 1: .\input, fil som matcher begge nøkkeltekster
    input_kandidater = kandidater(input_mappe)
    treff = [p for p in input_kandidater if _matcher_noekkeltekster(p.name, noekkeltekster)]
    if treff:
        return KildefilResultat(
            sti=treff[0],
            begrunnelse=f"Funnet i {input_mappe} med treff på nøkkeltekstene {noekkeltekster}.",
            kontrollerte_mapper=kontrollerte_mapper,
        )

    # Steg 2: rotmappe, fil som matcher begge nøkkeltekster
    rot_kandidater = kandidater(prosjektmappe)
    treff = [p for p in rot_kandidater if _matcher_noekkeltekster(p.name, noekkeltekster)]
    if treff:
        return KildefilResultat(
            sti=treff[0],
            begrunnelse=f"Funnet i prosjektets rotmappe med treff på nøkkeltekstene {noekkeltekster}.",
            kontrollerte_mapper=kontrollerte_mapper,
        )

    # Steg 3: andre xlsx-filer i .\input
    if input_kandidater:
        return KildefilResultat(
            sti=input_kandidater[0],
            begrunnelse=(
                f"Ingen fil i {input_mappe} matchet nøkkeltekstene {noekkeltekster}. "
                f"Falt tilbake til første gyldige Excel-fil i mappen: {input_kandidater[0].name}."
            ),
            kontrollerte_mapper=kontrollerte_mapper,
        )

    # Steg 4: andre xlsx-filer i rotmappen
    if rot_kandidater:
        return KildefilResultat(
            sti=rot_kandidater[0],
            begrunnelse=(
                "Ingen fil i input-mappen. Falt tilbake til første gyldige Excel-fil i "
                f"prosjektets rotmappe: {rot_kandidater[0].name}."
            ),
            kontrollerte_mapper=kontrollerte_mapper,
        )

    raise KildefilIkkeFunnet(
        "Fant ingen egnet kildefil ('Ført arbeid - utvidet.xlsx' eller annen .xlsx). "
        f"Kontrollerte mapper: {[str(m) for m in kontrollerte_mapper]}. "
        "Legg filen i .\\input eller i prosjektets rotmappe og prøv igjen."
    )
