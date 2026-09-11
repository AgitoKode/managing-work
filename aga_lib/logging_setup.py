"""Logging med personvernmaskering.

Kravet fra oppdraget er at loggfilen aldri skal inneholde fullstendige
persondata (ansattnavn, e-post, lønnsbeløp per person). Vi løser dette med
et logging.Filter som maskerer gjenkjennelige mønstre (e-postadresser og
tall som ligner personnummer/kontonummer) og ved disiplin i kallende kode:
det logges kun aggregerte tall, tekniske identifikatorer (ansattnr som tall,
ikke navn) og statusmeldinger - aldri hele datarader.
"""
from __future__ import annotations

import logging
import re
from pathlib import Path

_EPOST_MONSTER = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")
_MULIG_FNR_MONSTER = re.compile(r"\b\d{11}\b")


class PersonvernFilter(logging.Filter):
    """Maskerer e-postadresser og 11-sifrede tall (mulig fødselsnummer) i loggmeldinger."""

    def filter(self, record: logging.LogRecord) -> bool:
        msg = record.getMessage()
        maskert = _EPOST_MONSTER.sub("[e-post maskert]", msg)
        maskert = _MULIG_FNR_MONSTER.sub("[tall maskert]", maskert)
        if maskert != msg:
            record.msg = maskert
            record.args = ()
        return True


def sett_opp_logging(logg_mappe: Path, kjoretidspunkt: str) -> logging.Logger:
    logg_mappe.mkdir(parents=True, exist_ok=True)
    logg_fil = logg_mappe / f"aga_analyse_{kjoretidspunkt}.log"

    logger = logging.getLogger("aga_analyse")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    filfilter = PersonvernFilter()

    fil_handler = logging.FileHandler(logg_fil, encoding="utf-8")
    fil_handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    fil_handler.addFilter(filfilter)

    konsoll_handler = logging.StreamHandler()
    konsoll_handler.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
    konsoll_handler.addFilter(filfilter)

    logger.addHandler(fil_handler)
    logger.addHandler(konsoll_handler)
    logger.propagate = False

    logger.info("Logging initialisert. Loggfil: %s", logg_fil)
    return logger


def masker_navn(navn: str | None) -> str:
    """Erstatter et ansattnavn med initialer, til bruk i tekniske feilmeldinger/unntak."""
    if not navn:
        return "[ukjent]"
    deler = [d for d in str(navn).replace(",", " ").split() if d]
    if not deler:
        return "[ukjent]"
    return "".join(d[0].upper() for d in deler) + "."
