"""Fallback per seksjon 11: dersom offisiell kildefil for AGA-sone mangler og
nettverkstilgang til de offisielle domenene heller ikke er tilgjengelig, skal
AGA-sone/-sats IKKE fylles ut og IKKE gjettes. I stedet skrives
.\\output\\AGA_oppslagsbehov.xlsx med hva som må slås opp manuelt - resten av
dataanalysen (prosjektregister, kommuneidentifikasjon, lønnsgrunnlag mv.)
gjennomføres som normalt.
"""
from __future__ import annotations

import pandas as pd

STATUS_MANGLER_KILDE = "MANGLER_OFFISIELT_KILDEGRUNNLAG"


def bygg_oppslagsbehov(prosjektregister: list, analyseaar: int) -> pd.DataFrame:
    rader = []
    for p in prosjektregister:
        rader.append({
            "År": analyseaar,
            "Prosjektnr.": p.prosjektnr,
            "Prosjekt": "; ".join(p.prosjektnavn),
            "Postnummer": "; ".join(p.postnumre),
            "Foreslått kommune": p.kommune_resultat.kommune or "",
            "Kommunenummer": p.kommune_resultat.kommunenummer or "",
            "Oppslag som kreves": (
                "AGA-sone og -sats for kommunen (og evt. presis sonedeling) fra Skatteetaten/Lovdata/"
                "Regjeringen.no for gjeldende arbeidsår."
            ),
            "Kontrollstatus": STATUS_MANGLER_KILDE,
            "Kommentar": (
                "Offisiell kommunekatalog/satstabell var ikke tilgjengelig lokalt, og de offisielle "
                "kildedomenene var ikke nåbare fra dette kjøremiljøet i denne kjøringen."
            ),
        })
    return pd.DataFrame(rader)
