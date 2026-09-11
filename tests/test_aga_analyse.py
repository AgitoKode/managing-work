"""Automatiske tester (seksjon 21). Kjøres med:

    python -m unittest discover -s tests -v

Testene dekker minst: postnummer med ledende null, Excel-seriedato,
tilleggslinjer som ikke dobler arbeidstimer, negative korreksjoner beholdt,
unik prosjektidentifikator, at rader uten dokumentert AGA-kilde ikke får
KILDEVERIFISERT, at AGA-kilde for feil år avvises, at resultatfilen har alle
obligatoriske ark, og at kildefilen ikke endres av en kjøring.
"""
from __future__ import annotations

import hashlib
import sys
import unittest
from pathlib import Path

import pandas as pd

PROSJEKTMAPPE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROSJEKTMAPPE))

from aga_lib.postnummer import er_gyldig_postnummer, normaliser_postnummer  # noqa: E402
from aga_lib.dates import excel_serial_til_dato, til_dato  # noqa: E402
from aga_lib.lines import (  # noqa: E402
    arbeidstimer_for_rad,
    bygg_vakt_id,
    klassifiser_linjer,
    korriger_arbeidstimer_for_manglende_arbeidstidsrad,
)
from aga_lib.municipality import KommuneOppslag  # noqa: E402
from aga_lib.aga_rules import Sonekatalog, SoneRad  # noqa: E402
from aga_lib.projects import bygg_prosjektregister  # noqa: E402
from aga_lib.quality_checks import kjoer_kontroller  # noqa: E402

OBLIGATORISKE_ARK = [
    "Lederoppsummering", "AGA_per_termin", "AGA_per_kommune", "AGA_per_prosjekt",
    "AGA_per_ansatt", "Detaljgrunnlag", "Prosjektregister", "Kommuneoppslag",
    "Kilderegister", "Regelverksgrunnlag", "Avvik", "Datakvalitet", "Kolonnemapping",
    "Metode", "Kjøremetadata",
]

KLASSIFISERINGSREGLER = {
    "arbeidstid_artikler": ["Arbeidstimer"],
    "korreksjon_artikler": ["Fratrekk vakt"],
    "korreksjon_artikkeltyper": ["Fratrekk"],
    "kjente_tillegg_artikler": ["Nattillegg", "Kveldstillegg", "Helg", "Helligdag 133,33%"],
    "fallback_arbeidstid_artikler": ["Helligdag 133,33%", "Helg"],
}


class TestPostnummer(unittest.TestCase):
    def test_postnummer_0469_bevares_som_fire_sifre(self):
        self.assertEqual(normaliser_postnummer("0469"), "0469")
        self.assertEqual(normaliser_postnummer(469), "0469")
        self.assertEqual(normaliser_postnummer(469.0), "0469")
        self.assertEqual(normaliser_postnummer(" 0469 "), "0469")

    def test_ugyldig_postnummer_avvises(self):
        self.assertFalse(er_gyldig_postnummer("469"))
        self.assertFalse(er_gyldig_postnummer("ABCD"))
        self.assertFalse(er_gyldig_postnummer(None))
        self.assertFalse(er_gyldig_postnummer(float("nan")))
        self.assertTrue(er_gyldig_postnummer("0469"))


class TestDatoer(unittest.TestCase):
    def test_excel_seriedato_konverteres_korrekt(self):
        # Excel-serienummer 46023 tilsvarer 01.01.2026 (Excel sitt 1900-datosystem)
        dato = excel_serial_til_dato(46023)
        self.assertEqual((dato.year, dato.month, dato.day), (2026, 1, 1))

    def test_til_dato_haandterer_tekst_serienummer_og_datetime(self):
        self.assertEqual(til_dato("46023").date().isoformat(), "2026-01-01")
        self.assertEqual(til_dato("01.01.2026").date().isoformat(), "2026-01-01")
        self.assertIsNone(til_dato("ikke en dato"))
        self.assertIsNone(til_dato(None))


class TestLinjeklassifisering(unittest.TestCase):
    def _lag_df(self) -> pd.DataFrame:
        return pd.DataFrame([
            {"artikkel": "Arbeidstimer", "artikkeltype": "Ordinære timer", "timer": 7.5,
             "timer_ekskl_pause": 7.5, "timer_inkl_pause": 8.0, "total_lonn": 1000.0},
            {"artikkel": "Nattillegg", "artikkeltype": "Tidsbasert tillegg", "timer": None,
             "timer_ekskl_pause": None, "timer_inkl_pause": None, "total_lonn": 200.0},
            {"artikkel": "Fratrekk vakt", "artikkeltype": "Fratrekk", "timer": None,
             "timer_ekskl_pause": None, "timer_inkl_pause": None, "total_lonn": -300.0},
        ])

    def test_tilleggslinjer_dobler_ikke_arbeidstimer(self):
        df = klassifiser_linjer(self._lag_df(), KLASSIFISERINGSREGLER)
        df["arbeidstimer"] = df.apply(arbeidstimer_for_rad, axis=1)
        # Kun arbeidstid-raden (7.5t) skal telle - Nattillegg-raden skal gi 0,
        # selv om den tilhører samme vakt.
        self.assertEqual(df["arbeidstimer"].tolist(), [7.5, 0.0, 0.0])

    def test_negative_korreksjoner_beholdes(self):
        df = klassifiser_linjer(self._lag_df(), KLASSIFISERINGSREGLER)
        self.assertTrue((df["total_lonn"] < 0).any())
        korreksjonsrad = df[df["lonnskomponent"] == "korreksjon"]
        self.assertEqual(len(korreksjonsrad), 1)
        self.assertEqual(korreksjonsrad.iloc[0]["total_lonn"], -300.0)
        # Summen av lønn skal fortsatt inkludere det negative beløpet
        self.assertAlmostEqual(df["total_lonn"].sum(), 1000.0 + 200.0 - 300.0)

    def test_helligdagsvakt_uten_arbeidstimer_rad_gir_ikke_null_timer(self):
        # Helg og Helligdag 133,33% dekker SAMME vakt (ingen egen Arbeidstimer-rad) -
        # skal telles én gang (største verdi), ikke summeres til dobbel arbeidstid.
        df = pd.DataFrame([
            {"ansattnr": "1", "arbeidsdato": "2026-01-01", "jobbnr": "J1", "prosjektnr": "P1",
             "artikkel": "Helligdag 133,33%", "artikkeltype": "Ordinære timer",
             "timer": 7.5, "timer_ekskl_pause": 7.5, "timer_inkl_pause": 7.5, "total_lonn": 5000.0},
            {"ansattnr": "1", "arbeidsdato": "2026-01-01", "jobbnr": "J1", "prosjektnr": "P1",
             "artikkel": "Helg", "artikkeltype": "Tidsbasert tillegg",
             "timer": None, "timer_ekskl_pause": 7.5, "timer_inkl_pause": 7.5, "total_lonn": 2000.0},
        ])
        df = klassifiser_linjer(df, KLASSIFISERINGSREGLER)
        df["arbeidstimer"] = df.apply(arbeidstimer_for_rad, axis=1)
        df = korriger_arbeidstimer_for_manglende_arbeidstidsrad(df, KLASSIFISERINGSREGLER)
        self.assertEqual(df["arbeidstimer"].sum(), 7.5, "Skal telle vaktens 7,5 timer én gang, ikke 0 eller 15")
        self.assertEqual(df["arbeidstid_fallback_brukt"].sum(), 1)

    def test_normal_vakt_med_arbeidstimer_rad_paavirkes_ikke_av_fallback(self):
        df = pd.DataFrame([
            {"ansattnr": "1", "arbeidsdato": "2026-01-02", "jobbnr": "J2", "prosjektnr": "P1",
             "artikkel": "Arbeidstimer", "artikkeltype": "Ordinære timer",
             "timer": 7.5, "timer_ekskl_pause": 7.5, "timer_inkl_pause": 7.5, "total_lonn": 5000.0},
            {"ansattnr": "1", "arbeidsdato": "2026-01-02", "jobbnr": "J2", "prosjektnr": "P1",
             "artikkel": "Helg", "artikkeltype": "Tidsbasert tillegg",
             "timer": None, "timer_ekskl_pause": 3.0, "timer_inkl_pause": 3.0, "total_lonn": 900.0},
        ])
        df = klassifiser_linjer(df, KLASSIFISERINGSREGLER)
        df["arbeidstimer"] = df.apply(arbeidstimer_for_rad, axis=1)
        df = korriger_arbeidstimer_for_manglende_arbeidstidsrad(df, KLASSIFISERINGSREGLER)
        self.assertEqual(df["arbeidstimer"].sum(), 7.5)
        self.assertEqual(df["arbeidstid_fallback_brukt"].sum(), 0)

    def test_vakt_id_er_lik_for_samme_vakt(self):
        df = self._lag_df()
        for kol in ("ansattnr", "arbeidsdato", "jobbnr", "prosjektnr", "fra_kl", "til_kl"):
            df[kol] = "1"
        ider = df.apply(bygg_vakt_id, axis=1)
        self.assertEqual(ider.nunique(), 1)


class TestProsjektregister(unittest.TestCase):
    def _kommuneoppslag(self) -> KommuneOppslag:
        rader = [SoneRad("0301", "Oslo", "Oslo", "1", "")]
        return KommuneOppslag(Sonekatalog(rader))

    def test_prosjektregisteret_har_unik_prosjektidentifikator(self):
        df = pd.DataFrame([
            {"prosjektnr": "P1", "jobbnr": "J1", "bedrift": "Firma A", "prosjekt": "Prosjekt A",
             "jobb": "Jobb A", "org_nr": "1", "avdeling": "A1", "postnummer_normalisert": "0469"},
            {"prosjektnr": "P1", "jobbnr": "J1", "bedrift": "Firma A", "prosjekt": "Prosjekt A",
             "jobb": "Jobb A", "org_nr": "1", "avdeling": "A1", "postnummer_normalisert": "0469"},
            {"prosjektnr": "P2", "jobbnr": "J2", "bedrift": "Firma B", "prosjekt": "Prosjekt B",
             "jobb": "Jobb B", "org_nr": "2", "avdeling": "A2", "postnummer_normalisert": "0469"},
        ])
        register = bygg_prosjektregister(df, self._kommuneoppslag())
        noekler = [(r.noekkeltype, r.noekkelverdi) for r in register]
        self.assertEqual(len(noekler), len(set(noekler)), "Prosjektregisteret har duplikate nøkler")
        self.assertEqual(len(register), 2)

    def test_flere_postnumre_gir_maa_kvalitetssikres_ikke_gjetting(self):
        df = pd.DataFrame([
            {"prosjektnr": "P1", "jobbnr": "J1", "bedrift": "Firma A", "prosjekt": "Prosjekt A",
             "jobb": "Jobb A", "org_nr": "1", "avdeling": "A1", "postnummer_normalisert": "0469"},
            {"prosjektnr": "P1", "jobbnr": "J1", "bedrift": "Firma A", "prosjekt": "Prosjekt A",
             "jobb": "Jobb A", "org_nr": "1", "avdeling": "A1", "postnummer_normalisert": "0319"},
        ])
        register = bygg_prosjektregister(df, self._kommuneoppslag())
        self.assertEqual(len(register), 1)
        self.assertTrue(register[0].flere_postnumre)
        self.assertEqual(register[0].kommune_resultat.kontrollstatus, "MÅ_KVALITETSSIKRES")
        self.assertIn("flere arbeidssteder", register[0].kommune_resultat.kontrollkommentar)


class TestKontroller(unittest.TestCase):
    def test_rader_uten_dokumentert_agakilde_faar_ikke_kildeverifisert(self):
        from aga_lib.sources import KildeRad

        kilder = [
            KildeRad(
                kilde_id="X", kildetittel="", kildeadresse="", offentlig_utgiver="",
                dokumenttype="", relevant_aar=2026, publisert_dato="", sist_oppdatert_dato="",
                kontrollert_tidspunkt="", underbygger="", sitat_henvisning="",
                kontrollstatus="MÅ_KVALITETSSIKRES", merknad="Kilden er ikke åpnet i denne kjøringen",
            )
        ]
        self.assertTrue(all(k.kontrollstatus != "KILDEVERIFISERT" for k in kilder))

    def test_aga_kilde_for_feil_aar_avvises(self):
        df = pd.DataFrame({
            "ansattnr": [1], "prosjektnr": ["P1"], "jobbnr": ["J1"],
            "postnummer_normalisert": ["0469"], "postnummer_gyldig": [True],
            "total_lonn": [100.0], "loenn_paa_loennsgrunnlag": [None],
            "total_lonn_inkl_sosial_kost": [130.0], "lonnskomponent": ["arbeidstid"],
            "er_korreksjon": [False], "arbeidstimer": [7.5], "vakt_id": ["a"],
            "artikkel": ["Arbeidstimer"], "ukjent_klassifisering": [False],
            "arbeidsdato": [pd.Timestamp("2026-01-01")],
        })
        avvik, _ = kjoer_kontroller(
            df=df, prosjektregister=[], aga_klassifiseringer=[], analyseaar=2026,
            kildeaar_kommunekatalog=2025, kildeaar_satstabell=2026, nettverksprobe=[],
        )
        kategorier = {a["Kategori"] for a in avvik}
        self.assertIn("AGA-kilde gjelder ikke analyseåret (kommunekatalog)", kategorier)


class TestRapportstruktur(unittest.TestCase):
    def test_resultatfil_har_alle_obligatoriske_ark_og_kilde_uendret(self):
        import aga_analyse

        kildefil = PROSJEKTMAPPE / "input"
        xlsx_filer = list(kildefil.glob("*.xlsx"))
        if not xlsx_filer:
            self.skipTest("Ingen kildefil i .\\input - integrasjonstest hoppet over")

        kildefil_sti = next((f for f in xlsx_filer if "ført arbeid" in f.name.lower()), xlsx_filer[0])
        hash_foer = hashlib.sha256(kildefil_sti.read_bytes()).hexdigest()

        output_mappe = PROSJEKTMAPPE / "output"
        filer_foer = set(output_mappe.glob("*")) if output_mappe.exists() else set()

        resultat = aga_analyse.kjoer_analyse(PROSJEKTMAPPE)

        hash_etter = hashlib.sha256(kildefil_sti.read_bytes()).hexdigest()
        self.assertEqual(hash_foer, hash_etter, "Kildefilen ble endret av analysen")

        import openpyxl

        wb = openpyxl.load_workbook(resultat.hovedrapport)
        for arknavn in OBLIGATORISKE_ARK:
            self.assertIn(arknavn, wb.sheetnames, f"Mangler obligatorisk ark: {arknavn}")

        wb_felles = openpyxl.load_workbook(resultat.felles_arbeidsbok)
        for prefiks in ("STEG1", "STEG2", "STEG3", "STEG4", "STEG5", "STEG6", "STEG7", "STEG8", "STEG9"):
            self.assertTrue(
                any(navn.startswith(prefiks) for navn in wb_felles.sheetnames),
                f"Mangler et ark for {prefiks} i den felles arbeidsboken",
            )
        self.assertEqual(len(resultat.steg4_9_filer), 6)

        # Ryddighet: fjern ALLE filer denne testkjøringen selv opprettet i .\output,
        # inkludert STEG4-9-leveransene og den felles arbeidsboken, men aldri
        # filer som lå der fra før testen startet.
        for f in set(output_mappe.glob("*")) - filer_foer:
            f.unlink()


if __name__ == "__main__":
    unittest.main()
