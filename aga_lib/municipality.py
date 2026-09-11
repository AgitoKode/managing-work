"""Kommuneidentifikasjon for et enkelt prosjekt/arbeidssted.

Prioritering (jf. oppdragets seksjon 7):
1. Gyldig postnummer fra RecMan (slått opp i kuratert postnummer-referanse)
2. Kommunenavn eksplisitt angitt i Bedrift
3. Kommunenavn eksplisitt angitt i Prosjekt
4. Kommunenavn eksplisitt angitt i Jobb
5. Offentlig postnummerregister (generisk oppslag - denne løsningen har kun
   en kuratert delmengde, se postnummer_referanse.json; postnummer som ikke
   finnes der faller videre til manuell kontroll i stedet for å bli gjettet)
6. Manuell kontroll

Postnummer er hovedmetoden når det finnes og gir status KOMMUNE_VERIFISERT.
Navn funnet i fritekst gir kun KOMMUNE_SANNSYNLIG, siden et firmanavn med
"kommune" i seg ikke er en garanti for at ALLE prosjekter for kunden utføres
i den kommunen.
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path

from aga_lib.aga_rules import Sonekatalog
from aga_lib.postnummer import er_gyldig_postnummer, normaliser_postnummer

_DATA_MAPPE = Path(__file__).resolve().parent / "data"

KONTROLLSTATUS_VERIFISERT = "KOMMUNE_VERIFISERT"
KONTROLLSTATUS_SANNSYNLIG = "KOMMUNE_SANNSYNLIG"
KONTROLLSTATUS_MAA_KVALITETSSIKRES = "MÅ_KVALITETSSIKRES"
KONTROLLSTATUS_IKKE_FUNNET = "KOMMUNE_IKKE_FUNNET"
KONTROLLSTATUS_MOTSTRIDENDE = "MOTSTRIDENDE_OPPLYSNINGER"


@dataclass
class KommuneResultat:
    kommune: str | None
    kommunenummer: str | None
    benyttet_postnummer: str | None
    identifikasjonsmetode: str
    identifikasjonskilde: str
    kontrollstatus: str
    kontrollkommentar: str
    delkommune_merknad: str = ""


def _last_json(filnavn: str) -> dict:
    with (_DATA_MAPPE / filnavn).open("r", encoding="utf-8") as f:
        return json.load(f)


class KommuneOppslag:
    def __init__(self, sonekatalog: Sonekatalog):
        self._sonekatalog = sonekatalog
        self._postnummer_referanse = _last_json("postnummer_referanse.json")["oppslag"]
        self._manuelle_unntak = _last_json("manuelle_unntak.json")["bedrift_inneholder"]
        self._kommunenavn_index = self._bygg_kommunenavn_index(sonekatalog)

    @staticmethod
    def _bygg_kommunenavn_index(sonekatalog: Sonekatalog) -> dict[str, set[str]]:
        """Bygger navn(-del) -> {kommunenummer} fra Kommunekatalogen. Sammensatte
        navn som "Ráisa - Raisi - Nordreisa" splittes på " - " slik at hver
        stavemåte kan gjenkjennes i fritekst. Parentetiske presiseringer som
        "Herøy (Nordland)" beholdes som eget navn i tillegg til "Herøy"."""
        index: dict[str, set[str]] = {}
        for rad in sonekatalog.rader:
            navn_varianter = {rad.kommunenavn}
            for del_ in rad.kommunenavn.split(" - "):
                navn_varianter.add(del_.strip())
            uten_parentes = re.sub(r"\s*\([^)]*\)", "", rad.kommunenavn).strip()
            if uten_parentes:
                navn_varianter.add(uten_parentes)
            for navn in navn_varianter:
                if len(navn) < 3:
                    continue
                index.setdefault(navn.lower(), set()).add(rad.kommunenummer)
        return index

    def _soek_navn_i_tekst(self, tekst: str | None) -> set[str]:
        if not tekst:
            return set()
        lav = tekst.lower()
        treff: set[str] = set()
        for navn, kommunenumre in self._kommunenavn_index.items():
            if re.search(rf"\b{re.escape(navn)}\b", lav):
                treff |= kommunenumre
        return treff

    def _manuelt_unntak(self, *tekster: str | None) -> tuple[str, str, str] | None:
        for tekst in tekster:
            if not tekst:
                continue
            lav = tekst.lower()
            for noekkel, data in self._manuelle_unntak.items():
                if noekkel in lav:
                    return data["kommunenummer"], data["kommune"], data["begrunnelse"]
        return None

    def kommunenavn_for(self, kommunenummer: str) -> str | None:
        soner = self._sonekatalog.soner_for_kommune(kommunenummer)
        return soner[0].kommunenavn if soner else None

    def identifiser(
        self,
        postnummer_raw,
        bedrift: str | None,
        prosjekt: str | None,
        jobb: str | None,
    ) -> KommuneResultat:
        postnummer = normaliser_postnummer(postnummer_raw)

        if postnummer is not None and not er_gyldig_postnummer(postnummer):
            return KommuneResultat(
                kommune=None,
                kommunenummer=None,
                benyttet_postnummer=postnummer,
                identifikasjonsmetode="Postnummer (ugyldig format)",
                identifikasjonskilde="RecMan-datagrunnlag",
                kontrollstatus=KONTROLLSTATUS_MAA_KVALITETSSIKRES,
                kontrollkommentar=f"Postnummerfeltet inneholder en ugyldig verdi ({postnummer_raw!r}).",
            )

        # Metode 1: gyldig postnummer i kuratert referanse
        if postnummer is not None:
            ref = self._postnummer_referanse.get(postnummer)
            if ref is not None:
                return KommuneResultat(
                    kommune=ref["kommune"],
                    kommunenummer=ref["kommunenummer"],
                    benyttet_postnummer=postnummer,
                    identifikasjonsmetode="Postnummer (metode 1)",
                    identifikasjonskilde=ref["kilde"],
                    kontrollstatus=KONTROLLSTATUS_VERIFISERT,
                    kontrollkommentar="",
                    delkommune_merknad=ref.get("delkommune_merknad", ""),
                )

        # Metode 2-4 (kommunenavn i Bedrift/Prosjekt/Jobb) og det kuraterte
        # institusjonsunntaket samles ALLE opp før vi konkluderer, i stedet for
        # å returnere på første treff. Det er nødvendig fordi et prosjekt- eller
        # jobbnavn kan nevne én kommune (f.eks. et interkommunalt samarbeid) mens
        # institusjonen faktisk ligger i en annen - da skal det flagges som
        # motstridende, ikke stille velges den ene eller den andre.
        signaler: list[tuple[str, str, str]] = []  # (kommunenummer, metode, kilde)
        for feltnavn, tekst in (("Bedrift", bedrift), ("Prosjekt", prosjekt), ("Jobb", jobb)):
            treff = self._soek_navn_i_tekst(tekst)
            if len(treff) == 1:
                signaler.append((
                    next(iter(treff)),
                    f"Kommunenavn funnet i {feltnavn}",
                    f"Fritekst i feltet {feltnavn!r} i datagrunnlaget",
                ))
            elif len(treff) > 1:
                return KommuneResultat(
                    kommune=None,
                    kommunenummer=None,
                    benyttet_postnummer=postnummer,
                    identifikasjonsmetode=f"Kommunenavn funnet i {feltnavn} (flertydig)",
                    identifikasjonskilde=f"Fritekst i feltet {feltnavn!r} i datagrunnlaget",
                    kontrollstatus=KONTROLLSTATUS_MOTSTRIDENDE,
                    kontrollkommentar=(
                        f"Teksten i {feltnavn} matcher flere ulike kommuner ({sorted(treff)}) - "
                        "kan ikke avgjøres automatisk."
                    ),
                )

        manuelt = self._manuelt_unntak(bedrift, prosjekt, jobb)
        if manuelt is not None:
            kommunenummer, _kommune, begrunnelse = manuelt
            signaler.append((
                kommunenummer,
                "Manuelt kuratert unntak (institusjonskunnskap)",
                "documentation/Metodedokumentasjon.md - manuelle_unntak.json: " + begrunnelse,
            ))

        distinkte_kommuner = {s[0] for s in signaler}
        if len(distinkte_kommuner) == 1:
            kommunenummer = next(iter(distinkte_kommuner))
            metoder = ", ".join(sorted({s[1] for s in signaler}))
            kilder = "; ".join(sorted({s[2] for s in signaler}))
            return KommuneResultat(
                kommune=self.kommunenavn_for(kommunenummer),
                kommunenummer=kommunenummer,
                benyttet_postnummer=postnummer,
                identifikasjonsmetode=metoder,
                identifikasjonskilde=kilder,
                kontrollstatus=KONTROLLSTATUS_SANNSYNLIG,
                kontrollkommentar=(
                    "Kommune er utledet fra tekst/institusjonskunnskap, ikke fra postnummer. "
                    "Bør kontrolleres dersom kunden kan ha oppdrag i flere kommuner."
                ),
            )
        if len(distinkte_kommuner) > 1:
            detaljer = "; ".join(f"{s[1]} -> {self.kommunenavn_for(s[0])} ({s[0]})" for s in signaler)
            return KommuneResultat(
                kommune=None,
                kommunenummer=None,
                benyttet_postnummer=postnummer,
                identifikasjonsmetode="Flere kilder gir ulike kommuner",
                identifikasjonskilde="Bedrift/Prosjekt/Jobb-tekst og/eller manuelt kuratert unntak",
                kontrollstatus=KONTROLLSTATUS_MOTSTRIDENDE,
                kontrollkommentar=(
                    "Ulike signaler i datagrunnlaget peker mot ulike kommuner og kan ikke avgjøres "
                    f"automatisk: {detaljer}."
                ),
            )

        # Metode 6: manuell kontroll
        if postnummer is not None:
            kommentar = (
                f"Postnummer {postnummer} er gyldig formatert, men finnes ikke i den kuraterte "
                "postnummer-referansen. Krever oppslag i et fullstendig offentlig "
                "postnummerregister (metode 5) eller manuell kontroll (metode 6)."
            )
        else:
            kommentar = (
                "Ingen postnummer, og ingen kommune funnet i Bedrift/Prosjekt/Jobb-tekst. "
                "Krever manuell kontroll (metode 6)."
            )
        return KommuneResultat(
            kommune=None,
            kommunenummer=None,
            benyttet_postnummer=postnummer,
            identifikasjonsmetode="Ingen automatisk metode ga treff",
            identifikasjonskilde="",
            kontrollstatus=KONTROLLSTATUS_IKKE_FUNNET,
            kontrollkommentar=kommentar,
        )
