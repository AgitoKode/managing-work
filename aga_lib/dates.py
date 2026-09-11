"""Dato- og periodehåndtering: Excel-serienummer -> dato, og terminberegning.

Arbeidsdato skal alltid være datoen arbeidet faktisk ble utført - ikke
fakturadato eller registreringsdato. Vi bruker kolonnen "Dato" (RecMans
arbeidsdato) som primærkilde, og faller tilbake til "Første dag" dersom
"Dato" mangler. "Dato registrert" og en eventuell fakturadato brukes aldri
som arbeidsdato.
"""
from __future__ import annotations

import datetime as dt

import pandas as pd

EXCEL_EPOCH = dt.datetime(1899, 12, 30)  # Excel sin (feilaktige) dag-0, inkl. 1900-skuddårsbug


def excel_serial_til_dato(verdi: float) -> dt.datetime | None:
    try:
        return EXCEL_EPOCH + dt.timedelta(days=float(verdi))
    except (ValueError, OverflowError, TypeError):
        return None


def til_dato(verdi) -> pd.Timestamp | None:
    """Konverterer en celleverdi (datetime, pandas Timestamp, Excel-serienummer
    som tall/tekst, eller tekstdato) til en pandas Timestamp. Returnerer None
    dersom verdien ikke lar seg tolke - vi gjetter aldri en dato."""
    if verdi is None:
        return None
    if isinstance(verdi, pd.Timestamp):
        return verdi
    if isinstance(verdi, (dt.datetime, dt.date)):
        return pd.Timestamp(verdi)
    if isinstance(verdi, (int, float)):
        # Excel lagrer noen ganger datoer som serienummer selv om cellen
        # formateres som dato ved visning i Excel.
        if 0 < verdi < 600000:
            konvertert = excel_serial_til_dato(verdi)
            return pd.Timestamp(konvertert) if konvertert else None
        return None
    if isinstance(verdi, str):
        tekst = verdi.strip()
        if not tekst:
            return None
        # Ren serienummer-tekst, f.eks. "46023"
        if tekst.replace(",", ".").replace(".", "", 1).isdigit():
            tallverdi = float(tekst.replace(",", "."))
            konvertert = excel_serial_til_dato(tallverdi)
            return pd.Timestamp(konvertert) if konvertert else None
        for format_str in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                return pd.Timestamp(dt.datetime.strptime(tekst, format_str))
            except ValueError:
                continue
        try:
            return pd.Timestamp(tekst)
        except (ValueError, TypeError):
            return None
    return None


def termin_for_maaned(maaned: int, termin_maaneder: dict[int, list[int]]) -> int | None:
    for termin, maaneder in termin_maaneder.items():
        if maaned in maaneder:
            return termin
    return None
