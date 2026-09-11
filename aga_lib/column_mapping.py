"""Robust kobling fra originale RecMan-kolonnenavn til standardiserte feltnavn.

Håndterer store/små bokstaver, punktum, bindestrek, ekstra mellomrom og noen
vanlige alternative skrivemåter. Mapping-resultatet (original -> standardfelt)
skal alltid dokumenteres i rapportens "Kolonnemapping"-ark - se aga_lib.report.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass

# Standardfelt -> liste av kjente/alternative originalnavn (normalisert form
# sammenlignes, se _normaliser()).
FELT_ALIASER: dict[str, list[str]] = {
    "dato": ["dato"],
    "maaned": ["maned", "måned"],
    "aar": ["ar", "år"],
    "foerste_dag": ["forste dag", "første dag"],
    "siste_dag": ["siste dag"],
    "dato_registrert": ["dato registrert", "registreringsdato"],
    "ansattnr": ["ansattnr", "ansatt nr", "ansattnummer", "employee no", "emp nr"],
    "ansatt": ["ansatt", "ansattnavn", "navn"],
    "epost": ["e post", "epost", "email"],
    "artikkeltype": ["artikkeltype", "artikkel type"],
    "artikkel": ["artikkel"],
    "fra_kl": ["fra kl", "fra"],
    "til_kl": ["til kl", "til"],
    "pause": ["pause"],
    "timer_inkl_pause": ["timer inkl pause", "timer inkludert pause"],
    "timer_ekskl_pause": ["timer ekskl pause", "timer eksklusiv pause", "timer ekskludert pause"],
    "timer": ["timer"],
    "timelonn": ["timelonn", "timelønn"],
    "fakturapris": ["fakturapris"],
    "faktor": ["faktor"],
    "total_lonn": ["total lonn", "total lønn"],
    "sosial_kost": ["sosial kost"],
    "sosial_kost_pct": ["sosial kost prosent", "sosial kost %"],
    "total_lonn_inkl_sosial_kost": [
        "total lonn inkl sosial kost",
        "total lønn inkl sosial kost",
        "total lonn inkludert sosial kost",
    ],
    "total_faktura": ["total faktura"],
    "dekningsbidrag": ["dekningsbidrag"],
    "dekningsgrad": ["dekningsgrad"],
    "mva_sats": ["mva sats"],
    "mva": ["mva"],
    "status": ["status"],
    "jobbnr": ["jobbnr", "jobb nr", "jobbnummer"],
    "jobb": ["jobb"],
    "org_nr": ["org nr", "organisasjonsnummer", "orgnr"],
    "bedriftsnummer": ["bedriftsnummer", "bedrifts nummer"],
    "bedrift": ["bedrift", "kunde"],
    "postnummer": ["postnummer", "post nr", "postnr"],
    "prosjektnr": ["prosjektnr", "prosjekt nr", "prosjektnummer"],
    "prosjekt": ["prosjekt"],
    "avdelingsnr": ["avdelingsnr", "avdeling nr", "avdelingsnummer"],
    "avdeling": ["avdeling"],
    "loennsgrunnlagnr": ["lonnsgrunnlagnr", "lønnsgrunnlagnr", "lonnsgrunnlag nr"],
    "loenn_paa_loennsgrunnlag": [
        "lonn pa lonnsgrunnlag",
        "lønn på lønnsgrunnlag",
        "lonn pa lonnsgrunnlaget",
    ],
    "stillingsgruppe_nr": ["stillingsgruppe nr"],
    "stillingsgruppe": ["stillingsgruppe"],
    "stillingstype": ["stillingstype"],
}

# Kritiske standardfelt (jf. oppdragets seksjon 4). Minst ett felt fra hver
# gruppe må være til stede, ellers stoppes klassifiseringen for den gruppen
# og mangelen dokumenteres i kontrollrapporten/datakvalitet i stedet for at
# løsningen gjetter.
KRITISKE_FELTGRUPPER: dict[str, list[str]] = {
    "dato_eller_aar": ["dato", "aar"],
    "ansattidentifikator": ["ansattnr"],
    "prosjekt_eller_jobb_id": ["prosjektnr", "jobbnr"],
    "prosjekt_jobb_eller_bedrift_navn": ["prosjekt", "jobb", "bedrift"],
    "geografisk_info": ["postnummer", "bedrift", "prosjekt", "jobb"],
    "loennsbeloep": ["total_lonn", "loenn_paa_loennsgrunnlag", "total_lonn_inkl_sosial_kost"],
}


def _normaliser(tekst: str) -> str:
    """Normaliserer et kolonnenavn for sammenligning: små bokstaver, trimmet,
    punktum/bindestrek/understrek -> mellomrom, og en aksent-fri variant lagt
    til implisitt via alias-listene (både "år" og "ar", "lønn" og "lonn" osv.
    er registrert som aliaser - selve normaliseringen gjør IKKE aksentstripping,
    for å unngå å slå sammen ord som faktisk betyr noe annet uten æøå)."""
    tekst = unicodedata.normalize("NFC", tekst)
    tekst = tekst.strip().lower()
    tekst = tekst.replace(".", " ")
    tekst = tekst.replace("-", " ")
    tekst = tekst.replace("_", " ")
    tekst = re.sub(r"\s+", " ", tekst)
    return tekst.strip()


@dataclass
class KolonnemappingResultat:
    original_til_standard: dict[str, str]
    standard_til_original: dict[str, str]
    umappede_originalkolonner: list[str]
    manglende_kritiske_grupper: list[str]

    def dokumentasjonstabell(self) -> list[dict[str, str]]:
        rader = []
        for original, standard in self.original_til_standard.items():
            rader.append({"Original kolonne": original, "Standardfelt": standard})
        for original in self.umappede_originalkolonner:
            rader.append({"Original kolonne": original, "Standardfelt": "(ikke koblet)"})
        return rader


def bygg_kolonnemapping(originale_kolonnenavn: list[str]) -> KolonnemappingResultat:
    normalisert_til_alias: dict[str, str] = {}
    for standard, aliaser in FELT_ALIASER.items():
        for alias in aliaser:
            normalisert_til_alias[_normaliser(alias)] = standard

    original_til_standard: dict[str, str] = {}
    umappede: list[str] = []
    for kolonne in originale_kolonnenavn:
        if kolonne is None:
            continue
        norm = _normaliser(str(kolonne))
        standard = normalisert_til_alias.get(norm)
        if standard is None:
            umappede.append(str(kolonne))
            continue
        original_til_standard[str(kolonne)] = standard

    standard_til_original = {v: k for k, v in original_til_standard.items()}

    manglende_grupper = []
    for gruppe, felter in KRITISKE_FELTGRUPPER.items():
        if not any(felt in standard_til_original for felt in felter):
            manglende_grupper.append(gruppe)

    return KolonnemappingResultat(
        original_til_standard=original_til_standard,
        standard_til_original=standard_til_original,
        umappede_originalkolonner=umappede,
        manglende_kritiske_grupper=manglende_grupper,
    )
