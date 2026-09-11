"""Innlasting av config.json og oppsett av standard mappestruktur."""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class Konfigurasjon:
    raw: dict[str, Any]
    prosjektmappe: Path

    @property
    def analyseaar(self) -> int:
        return int(self.raw["analyseaar"])

    @property
    def sammenligningssats_prosent(self) -> float:
        return float(self.raw["sammenligningssats_prosent"])

    @property
    def termin_maaneder(self) -> dict[int, list[int]]:
        return {int(k): v for k, v in self.raw["termin_maaneder"].items()}

    def sti(self, navn: str) -> Path:
        """Returnerer en absolutt Path for en navngitt undermappe (input/output/cache/logs/...)."""
        rel = self.raw["stier"][navn]
        p = self.prosjektmappe / rel
        p.mkdir(parents=True, exist_ok=True)
        return p

    @property
    def offisielle_kildedomener(self) -> list[str]:
        return list(self.raw["offisielle_kildedomener"])

    @property
    def nettverk_timeout(self) -> float:
        return float(self.raw["nettverk_probe_timeout_sekunder"])

    @property
    def kildefil_sok(self) -> dict[str, Any]:
        return self.raw["kildefil_sok"]

    @property
    def linjeklassifisering(self) -> dict[str, Any]:
        return self.raw["linjeklassifisering"]

    @property
    def kildearkiv_filer(self) -> dict[str, str]:
        return self.raw["kildearkiv_filer"]


def last_konfigurasjon(prosjektmappe: Path | None = None, filnavn: str = "config.json") -> Konfigurasjon:
    """Leser config.json fra prosjektmappen (default: mappen der dette kjøres fra)."""
    base = prosjektmappe or Path(__file__).resolve().parent.parent
    sti = base / filnavn
    if not sti.exists():
        raise FileNotFoundError(
            f"Fant ikke konfigurasjonsfilen '{filnavn}' i {base}. "
            "Kjør programmet fra prosjektmappen, eller opprett config.json på nytt."
        )
    with sti.open("r", encoding="utf-8") as f:
        raw = json.load(f)
    return Konfigurasjon(raw=raw, prosjektmappe=base)
