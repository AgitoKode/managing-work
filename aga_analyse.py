#!/usr/bin/env python3
"""aga_analyse.py - AGA-analyse av RecMan-eksporten "Ført arbeid - utvidet.xlsx".

Kjøres fra prosjektmappen:

    python aga_analyse.py

Se README.md for full bruksanvisning (inkl. run_aga_analyse.ps1 for Windows)
og documentation/Metodedokumentasjon.md for metodikk. Dette skriptet endrer
ALDRI kildefilen - det åpnes read-only og resultatet skrives til en ny fil
i .\\output.
"""
from __future__ import annotations

import datetime as dt
import sys
import traceback
from pathlib import Path

import pandas as pd

from aga_lib.aga_classification import klassifiser as klassifiser_aga
from aga_lib.aga_rules import Sonekatalog, les_kommunekatalog
from aga_lib.column_mapping import bygg_kolonnemapping
from aga_lib.config import last_konfigurasjon
from aga_lib.dates import termin_for_maaned, til_dato
from aga_lib.excel_io import les_arbeidsbok
from aga_lib.file_discovery import KildefilIkkeFunnet, finn_kildefil
from aga_lib.leveranser_steg4_9 import bygg_og_skriv_leveranser
from aga_lib.lines import (
    arbeidstimer_for_rad,
    bygg_vakt_id,
    klassifiser_linjer,
    korriger_arbeidstimer_for_manglende_arbeidstidsrad,
)
from aga_lib.logging_setup import sett_opp_logging
from aga_lib.metadata import Kjoremetadata, hent_pakkeversjoner, python_versjon_streng
from aga_lib.municipality import KommuneOppslag
from aga_lib.oppslagsbehov import STATUS_MANGLER_KILDE, bygg_oppslagsbehov
from aga_lib.postnummer import er_gyldig_postnummer, normaliser_postnummer
from aga_lib.projects import bygg_prosjektregister
from aga_lib.quality_checks import kjoer_kontroller
from aga_lib.aggregations import bygg_oppsummering
from aga_lib import report
from aga_lib.sources import KildeRad, naa_iso, sha256_for_fil, test_nettverkstilgang

PROGRAMVERSJON = "1.0.0"
FORELOPIG_BEREGNING_BANNER = (
    "FORELØPIG BEREGNING - MÅ KONTROLLERES AV LØNN. Alle AGA-sone-/satsfunn i denne rapporten er "
    "GEOGRAFISK ANALYSEGRUNNLAG, ikke bekreftet endelig AGA - se ark 'Regelverksgrunnlag'."
)


class KritiskFeil(Exception):
    pass


def _les_kilde_dataframe(prosjektmappe: Path, konfig) -> tuple[pd.DataFrame, dict, Path, str]:
    kildefil_resultat = finn_kildefil(
        prosjektmappe=prosjektmappe,
        input_mappe=konfig.sti("input_mappe"),
        noekkeltekster=konfig.kildefil_sok["krev_alle_tekster_i_navn"],
        ekskluder_tekster=konfig.kildefil_sok["ekskluder_tekster_i_navn"],
        tillatte_endelser=konfig.kildefil_sok["tillatte_filendelser"],
    )
    sha256 = sha256_for_fil(kildefil_resultat.sti)
    ark = les_arbeidsbok(kildefil_resultat.sti)

    dataark_kandidater = {navn: a for navn, a in ark.items() if a.er_dataark}
    if not dataark_kandidater:
        raise KritiskFeil(
            "Fant ingen ark i kildefilen som ligner på et RecMan-arbeidsdataark "
            f"(sjekket arkene: {list(ark.keys())})."
        )
    hoved_arknavn = max(dataark_kandidater, key=lambda n: len(dataark_kandidater[n].df))
    hovedark = dataark_kandidater[hoved_arknavn]

    mapping = hovedark.kolonnemapping
    if mapping.manglende_kritiske_grupper:
        raise KritiskFeil(
            "Kritiske felt mangler i kildefilen og løsningen gjetter ikke: "
            f"{mapping.manglende_kritiske_grupper}. Kontroller kolonneoverskriftene i "
            f"arket '{hoved_arknavn}'."
        )

    df = hovedark.df.rename(columns=mapping.original_til_standard)
    df = df[[c for c in df.columns if c in mapping.standard_til_original]]
    return df, {"mapping": mapping, "arknavn": hoved_arknavn, "alle_ark": ark}, kildefil_resultat.sti, sha256


def kjoer_analyse(prosjektmappe: Path) -> Path:
    konfig = last_konfigurasjon(prosjektmappe)
    kjoretidspunkt = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    logger = sett_opp_logging(konfig.sti("logg_mappe"), kjoretidspunkt)

    logger.info("Starter AGA-analyse (programversjon %s)", PROGRAMVERSJON)

    try:
        df, lesekontekst, kildefil_sti, kildefil_sha256 = _les_kilde_dataframe(prosjektmappe, konfig)
    except KildefilIkkeFunnet as e:
        logger.error("Fant ikke kildefil: %s", e)
        raise
    except KritiskFeil as e:
        logger.error("Kritisk feil ved lesing av kildefil: %s", e)
        raise

    logger.info("Leste %s rader fra ark '%s' i %s", len(df), lesekontekst["arknavn"], kildefil_sti.name)
    antall_innleste = len(df)

    # --- Datorensing ---
    df["arbeidsdato"] = df["dato"].apply(til_dato) if "dato" in df.columns else pd.NaT
    if "foerste_dag" in df.columns:
        mangler_dato = df["arbeidsdato"].isna()
        df.loc[mangler_dato, "arbeidsdato"] = df.loc[mangler_dato, "foerste_dag"].apply(til_dato)
    df["aar"] = df["arbeidsdato"].apply(lambda d: d.year if pd.notna(d) else None)
    df["maaned"] = df["arbeidsdato"].apply(lambda d: d.month if pd.notna(d) else None)
    df["termin"] = df["maaned"].apply(
        lambda m: termin_for_maaned(int(m), konfig.termin_maaneder) if m is not None else None
    )

    analyseaar = konfig.analyseaar
    antall_totalt_med_aar = df["aar"].notna().sum()
    df_analyseaar = df[df["aar"] == analyseaar].copy()
    antall_ekskludert_andre_aar = int((df["aar"].notna() & (df["aar"] != analyseaar)).sum())
    logger.info(
        "%s rader er for analyseåret %s. %s rader fra andre år er beholdt i rådatagrunnlaget, men "
        "ekskludert fra hovedrapporten.",
        len(df_analyseaar), analyseaar, antall_ekskludert_andre_aar,
    )

    df = df_analyseaar

    # --- Postnummer ---
    df["postnummer_normalisert"] = df["postnummer"].apply(normaliser_postnummer) if "postnummer" in df.columns else None
    df["postnummer_gyldig"] = df["postnummer_normalisert"].apply(er_gyldig_postnummer)

    # --- Linjeklassifisering og vakt-ID ---
    df = klassifiser_linjer(df, konfig.linjeklassifisering)
    df["arbeidstimer"] = df.apply(arbeidstimer_for_rad, axis=1)
    df = korriger_arbeidstimer_for_manglende_arbeidstidsrad(df, konfig.linjeklassifisering)
    antall_fallback_vakter = int(df["arbeidstid_fallback_brukt"].sum())
    if antall_fallback_vakter:
        logger.info(
            "%s vakter manglet en egen 'Arbeidstimer'-rad - arbeidstid hentet fra Helg/Helligdag-artikkel "
            "i stedet (se config.json sin fallback_arbeidstid_artikler og Datakvalitet-arket).",
            antall_fallback_vakter,
        )
    df["vakt_id"] = df.apply(bygg_vakt_id, axis=1)

    for felt in ("total_lonn", "sosial_kost", "total_lonn_inkl_sosial_kost", "loenn_paa_loennsgrunnlag"):
        if felt in df.columns:
            df[felt] = pd.to_numeric(
                df[felt].astype(str).str.replace(" ", "").str.replace(",", "."), errors="coerce"
            )
    df["lonn_uten_sosial_kost"] = df.get("total_lonn", pd.Series(0.0, index=df.index)).fillna(0.0)
    df["sosial_kost"] = df.get("sosial_kost", pd.Series(0.0, index=df.index)).fillna(0.0)
    df["lonn_inkl_sosial_kost"] = df.get(
        "total_lonn_inkl_sosial_kost", df["lonn_uten_sosial_kost"] + df["sosial_kost"]
    ).fillna(df["lonn_uten_sosial_kost"] + df["sosial_kost"])

    # --- Kilder: kommunekatalog + satstabell ---
    kildearkiv = konfig.sti("kildearkiv_mappe")
    kk_sti = kildearkiv / konfig.kildearkiv_filer["kommunekatalog_csv"]

    nettverksprobe = test_nettverkstilgang(konfig.offisielle_kildedomener, konfig.nettverk_timeout)
    for probe in nettverksprobe:
        logger.info("Nettverksprobe %s: naaadd=%s (%s)", probe.vert, probe.naaadd, probe.detaljer)
    noen_kilder_naaadde = any(p.naaadd for p in nettverksprobe)

    kilde_tilgjengelig = kk_sti.exists()
    if not kilde_tilgjengelig:
        logger.warning(
            "Fant ikke kommunekatalogen (%s), og ingen av de offisielle kildedomenene var nåbare "
            "(%s). Jf. seksjon 11: AGA-sone/-sats fylles IKKE ut - se AGA_oppslagsbehov.xlsx. "
            "Resten av dataanalysen gjennomføres som normalt.",
            kk_sti, noen_kilder_naaadde,
        )
        kommunekatalog_rader: list = []
        kommunekatalog_advarsler: list[str] = []
    else:
        kommunekatalog_rader, kommunekatalog_advarsler = les_kommunekatalog(kk_sti)
    for adv in kommunekatalog_advarsler:
        logger.warning(adv)
    sonekatalog = Sonekatalog(kommunekatalog_rader)
    kommuneoppslag = KommuneOppslag(sonekatalog)

    # --- Prosjektregister og kommuneidentifikasjon ---
    prosjektregister = bygg_prosjektregister(df, kommuneoppslag)
    logger.info("Bygget prosjektregister med %s unike prosjekter/jobber.", len(prosjektregister))

    kilde_id_geografi = "KK-2026+SATS-2026"
    aga_per_prosjektnoekkel: dict[tuple[str, str], object] = {}
    for p in prosjektregister:
        if not kilde_tilgjengelig:
            klass = klassifiser_aga(
                sonekatalog=sonekatalog,  # tom katalog
                kommunenummer=None,
                kommune=p.kommune_resultat.kommune,
                postnummer=None,
                delkommune_merknad="",
                kilde_id="",
            )
            klass.kontrollstatus = STATUS_MANGLER_KILDE
            klass.kontrollkommentar = (
                "Offisiell kommunekatalog/satstabell var ikke tilgjengelig lokalt, og de offisielle "
                "kildedomenene var ikke nåbare fra dette kjøremiljøet. Se .\\output\\AGA_oppslagsbehov.xlsx."
            )
            aga_per_prosjektnoekkel[(p.noekkeltype, p.noekkelverdi)] = klass
            continue
        postnummer_for_klassifisering = None
        if len(p.postnumre) == 1 and er_gyldig_postnummer(p.postnumre[0]):
            postnummer_for_klassifisering = p.postnumre[0]
        klass = klassifiser_aga(
            sonekatalog=sonekatalog,
            kommunenummer=p.kommune_resultat.kommunenummer,
            kommune=p.kommune_resultat.kommune,
            postnummer=postnummer_for_klassifisering,
            delkommune_merknad=p.kommune_resultat.delkommune_merknad,
            kilde_id=kilde_id_geografi,
        )
        aga_per_prosjektnoekkel[(p.noekkeltype, p.noekkelverdi)] = klass

    def _prosjektnoekkel_for_rad(rad) -> tuple[str, str]:
        prosjektnr = rad.get("prosjektnr")
        if prosjektnr is not None and str(prosjektnr).strip() not in ("", "nan", "None"):
            return ("prosjektnr", str(prosjektnr).strip())
        jobbnr = rad.get("jobbnr")
        if jobbnr is not None and str(jobbnr).strip() not in ("", "nan", "None"):
            return ("jobbnr", str(jobbnr).strip())
        return ("ukjent", f"UKJENT-{rad.get('bedrift', 'ukjent bedrift')}")

    noekler = df.apply(_prosjektnoekkel_for_rad, axis=1)
    df["kommune"] = [aga_per_prosjektnoekkel[n].kommune for n in noekler]
    df["kommunenummer"] = [aga_per_prosjektnoekkel[n].kommunenummer for n in noekler]
    df["sone"] = [aga_per_prosjektnoekkel[n].sone for n in noekler]
    df["ordinaer_sats"] = [aga_per_prosjektnoekkel[n].ordinaer_sats for n in noekler]
    df["aga_kontrollstatus"] = [aga_per_prosjektnoekkel[n].kontrollstatus for n in noekler]
    df["aga_kontrollkommentar"] = [aga_per_prosjektnoekkel[n].kontrollkommentar for n in noekler]
    df["aga_kilde_id"] = [aga_per_prosjektnoekkel[n].kilde_id for n in noekler]

    df["foreloepig_aga_belop"] = df.apply(
        lambda r: r["lonn_uten_sosial_kost"] * r["ordinaer_sats"] if pd.notna(r["ordinaer_sats"]) else None,
        axis=1,
    )

    # --- Kontroller ---
    avvik, datakvalitet = kjoer_kontroller(
        df=df,
        prosjektregister=prosjektregister,
        aga_klassifiseringer=list(aga_per_prosjektnoekkel.values()),
        analyseaar=analyseaar,
        kildeaar_kommunekatalog=2026,
        kildeaar_satstabell=2026,
        nettverksprobe=nettverksprobe,
    )

    # --- Aggregeringer ---
    sammenligningssats = konfig.sammenligningssats_prosent
    agg_termin = bygg_oppsummering(df, ["aar", "termin", "maaned"], sammenligningssats)
    agg_kommune = bygg_oppsummering(df, ["kommune", "kommunenummer", "sone"], sammenligningssats)
    agg_prosjekt = bygg_oppsummering(df, ["prosjektnr", "jobbnr"], sammenligningssats)
    agg_ansatt = bygg_oppsummering(df, ["ansattnr"], sammenligningssats)

    def _sha256_hvis_finnes(sti: Path) -> str:
        return sha256_for_fil(sti) if sti.exists() else ""

    kk_status = "MÅ_KVALITETSSIKRES" if kilde_tilgjengelig else STATUS_MANGLER_KILDE
    kk_merknad = (
        "Skatteetaten.no er ikke nåbar fra dette kjøremiljøet (se nettverksprobe i Datakvalitet). "
        "Datasettet er mottatt som CSV-eksport fra oppdragsgiver 2026-09-11 og lagt i "
        "source_archive med SHA-256-sum. Siden løsningen ikke selv har åpnet den levende "
        "kildesiden i denne kjøringen, gis status MÅ_KVALITETSSIKRES i tråd med seksjon 15."
    ) if kilde_tilgjengelig else (
        f"Fant ikke kildefilen ({kk_sti}), og de offisielle kildedomenene var heller ikke nåbare. "
        "AGA-sone/-sats er derfor ikke fylt ut - se .\\output\\AGA_oppslagsbehov.xlsx."
    )

    # --- Kilderegister ---
    kontrolltidspunkt = naa_iso()
    kilderader = [
        KildeRad(
            kilde_id="KK-2026",
            kildetittel="Satser for arbeidsgiveravgift - soneinndeling (Kommunekatalog 2026)",
            kildeadresse="https://www.skatteetaten.no/satser/arbeidsgiveravgift-soneinndeling/#kommunekatalog-2026",
            offentlig_utgiver="Skatteetaten",
            dokumenttype="Nettside / CSV-eksport av offisiell kommune-sone-tabell",
            relevant_aar=2026,
            publisert_dato="",
            sist_oppdatert_dato="Ikke oppgitt i mottatt uttrekk",
            kontrollert_tidspunkt=kontrolltidspunkt,
            underbygger="AGA-sone per kommunenummer for 2026, inkl. kommuner delt mellom flere soner",
            sitat_henvisning="Kolonnene Kommunenr./Kommunenavn/Fylke/Sone/Kommentar i kildens fullstendige tabell",
            kontrollstatus=kk_status,
            merknad=kk_merknad,
            sha256=_sha256_hvis_finnes(kk_sti),
        ),
        KildeRad(
            kilde_id="SATS-2026",
            kildetittel="Satser for arbeidsgiveravgift - soneinndeling (satstabell og fribeløpsregel)",
            kildeadresse="https://www.skatteetaten.no/satser/arbeidsgiveravgift-soneinndeling/#kommunekatalog-2026",
            offentlig_utgiver="Skatteetaten",
            dokumenttype="Nettside (satstabell)",
            relevant_aar=2026,
            publisert_dato="",
            sist_oppdatert_dato="Ikke oppgitt i mottatt uttrekk",
            kontrollert_tidspunkt=kontrolltidspunkt,
            underbygger="Ordinær AGA-sats per sone (I, Ia, II, III, IV, IVa, V) og fribeløpsregelen for sone Ia i 2026",
            sitat_henvisning=(
                "\"I sone 1a skal det betales arbeidsgiveravgift med en sats på 10,6 prosent inntil "
                "differansen ... er lik fribeløpet. I 2026 er fribeløpet 850 000 kroner per foretak.\""
            ),
            kontrollstatus=kk_status,
            merknad="Samme nettverksbegrensning som KK-2026. Se source_archive/skatteetaten_satser_2026_brukeroppgitt.txt." if kilde_tilgjengelig else kk_merknad,
            sha256=_sha256_hvis_finnes(kildearkiv / konfig.kildearkiv_filer["satser_kilde_tekst"]),
        ),
        KildeRad(
            kilde_id="PN-REF-2026",
            kildetittel="Postnummerregister (kuratert delmengde for observerte postnumre)",
            kildeadresse="https://www.bring.no/tjenester/adressetjenester/postnummer",
            offentlig_utgiver="Bring/Posten Norge AS (offentlig postnummerregister)",
            dokumenttype="Oppslagstjeneste, kuratert manuelt per postnummer",
            relevant_aar=2026,
            publisert_dato="",
            sist_oppdatert_dato="",
            kontrollert_tidspunkt="2026-09-11",
            underbygger="Kobling postnummer -> kommunenummer for postnumrene observert i datagrunnlaget",
            sitat_henvisning="Se aga_lib/data/postnummer_referanse.json for postnummer-for-postnummer-kilder",
            kontrollstatus="MÅ_KVALITETSSIKRES",
            merknad=(
                "Dekker kun postnumrene i denne kjøringens datagrunnlag - IKKE et fullstendig "
                "postnummerregister. Ikke selv slått opp direkte mot bring.no i denne kjøringen."
            ),
        ),
        KildeRad(
            kilde_id="REGELVERK-AGA-2026",
            kildetittel=(
                "Uavklart: rettslig grunnlag for geografisk AGA-sone ved utleie av arbeidskraft (forskrift/"
                "retningslinjer om differensiert arbeidsgiveravgift for 2026)"
            ),
            kildeadresse="",
            offentlig_utgiver="Skatteetaten / Lovdata / Regjeringen.no (ikke bekreftet)",
            dokumenttype="Ikke funnet/åpnet i denne kjøringen",
            relevant_aar=2026,
            publisert_dato="",
            sist_oppdatert_dato="",
            kontrollert_tidspunkt=kontrolltidspunkt,
            underbygger="Hvorvidt arbeidsstedets kommune (fremfor registrert underenhet) er riktig AGA-basis",
            sitat_henvisning="",
            kontrollstatus="MÅ_KVALITETSSIKRES",
            merknad=(
                "Se ark 'Regelverksgrunnlag'. Nettverkstilgang til skatteetaten.no/lovdata.no/regjeringen.no "
                "var ikke tilgjengelig fra dette kjøremiljøet (se nettverksprobe)."
            ),
        ),
    ]

    metadata = Kjoremetadata(
        kjoredato_tidspunkt=kontrolltidspunkt,
        python_versjon=python_versjon_streng(),
        kildefil_navn=kildefil_sti.name,
        kildefil_sha256=kildefil_sha256,
        antall_innleste_rader=antall_innleste,
        antall_analyserte_rader=len(df),
        antall_ekskluderte_rader=antall_ekskludert_andre_aar,
        antall_unike_ansatte=int(df["ansattnr"].nunique()),
        antall_unike_prosjekter=len(prosjektregister),
        antall_identifiserte_kommuner=len({p.kommune_resultat.kommunenummer for p in prosjektregister if p.kommune_resultat.kommunenummer}),
        antall_kildeverifiserte_kommune_sonekoblinger=sum(
            1 for r in kilderader if r.kontrollstatus == "KILDEVERIFISERT"
        ),
        programversjon=PROGRAMVERSJON,
        pakkeversjoner=hent_pakkeversjoner(),
    )

    if not kilde_tilgjengelig:
        oppslagsbehov_df = bygg_oppslagsbehov(prosjektregister, analyseaar)
        oppslagsbehov_sti = konfig.sti("output_mappe") / "AGA_oppslagsbehov.xlsx"
        wb_oppslag = report.lag_arbeidsbok()
        report.skriv_tabellark(wb_oppslag, "Oppslagsbehov", oppslagsbehov_df, status_kolonner=["Kontrollstatus"])
        report.lagre(wb_oppslag, oppslagsbehov_sti)
        logger.warning("Skrev %s (offisiell AGA-kilde mangler - se seksjon 11).", oppslagsbehov_sti)

    output_sti = konfig.sti("output_mappe") / f"AGA_Rapport_{analyseaar}.xlsx"
    if output_sti.exists():
        tidsstempel = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_sti = konfig.sti("output_mappe") / f"AGA_Rapport_{analyseaar}_{tidsstempel}.xlsx"
        logger.info("Tidligere rapport finnes fra før - lagrer med tidsstempel: %s", output_sti.name)

    wb = report.lag_arbeidsbok()
    _skriv_rapport(
        wb=wb,
        df=df,
        prosjektregister=prosjektregister,
        agg_termin=agg_termin,
        agg_kommune=agg_kommune,
        agg_prosjekt=agg_prosjekt,
        agg_ansatt=agg_ansatt,
        kilderader=kilderader,
        avvik=avvik,
        datakvalitet=datakvalitet,
        kolonnemapping=lesekontekst["mapping"],
        metadata=metadata,
        analyseaar=analyseaar,
        antall_ekskludert_andre_aar=antall_ekskludert_andre_aar,
        nettverksprobe=nettverksprobe,
    )
    report.lagre(wb, output_sti)
    logger.info("Rapport lagret: %s", output_sti)

    # --- STEG 4-9: seks frittstående leveransefiler i tillegg til hovedrapporten ---
    kk_kilderad = next(r for r in kilderader if r.kilde_id == "KK-2026")
    leveranser = bygg_og_skriv_leveranser(
        output_mappe=konfig.sti("output_mappe"),
        df=df,
        prosjektregister=prosjektregister,
        aga_per_prosjektnoekkel=aga_per_prosjektnoekkel,
        kilde_id_geografi=kilde_id_geografi,
        kilde_tilgjengelig=kilde_tilgjengelig,
        kk_url=kk_kilderad.kildeadresse,
        kk_tittel=kk_kilderad.kildetittel,
        kk_utgiver=kk_kilderad.offentlig_utgiver,
        kontrolltidspunkt=kontrolltidspunkt,
        analyseaar=analyseaar,
        logger=logger,
    )
    for f in leveranser:
        logger.info("STEG 4-9-leveranse skrevet: %s", f)

    return output_sti


def _skriv_rapport(
    wb, df, prosjektregister, agg_termin, agg_kommune, agg_prosjekt, agg_ansatt,
    kilderader, avvik, datakvalitet, kolonnemapping, metadata, analyseaar,
    antall_ekskludert_andre_aar, nettverksprobe,
) -> None:
    kildeverifiserte = sum(1 for r in kilderader if r.kontrollstatus == "KILDEVERIFISERT")
    maa_kvalitetssikres_kommuner = sum(
        1 for p in prosjektregister if p.kommune_resultat.kontrollstatus != "KOMMUNE_VERIFISERT"
    )
    unike_soner = sorted({p.kommune_resultat.kommune for p in prosjektregister if p.kommune_resultat.kommune})

    lederlinjer = [
        ("DATAGRUNNLAG", f"{metadata.kildefil_navn} (SHA-256: {metadata.kildefil_sha256[:16]}...)"),
        ("PERIODE I FILEN", f"Analyseår: {analyseaar}. Rader utenfor analyseåret ekskludert fra hovedrapport: {antall_ekskludert_andre_aar}."),
        ("SAMLET LØNNSGRUNNLAG (uten sosial kost)", f"{df['lonn_uten_sosial_kost'].sum():,.2f} kr".replace(",", " ").replace(".", ",")),
        ("ANTALL UNIKE ANSATTE", str(df["ansattnr"].nunique())),
        ("ANTALL UNIKE PROSJEKTER", str(len(prosjektregister))),
        ("ANTALL IDENTIFISERTE KOMMUNER", str(len(unike_soner))),
        ("KOMMUNER I DATAGRUNNLAGET", ", ".join(unike_soner) if unike_soner else "Ingen identifisert"),
        ("KILDEVERIFISERTE KILDEREGISTER-RADER", f"{kildeverifiserte} av {len(kilderader)} - se merknad under"),
        (
            "HVORFOR INGEN RADER ER KILDEVERIFISERT",
            "Skatteetaten.no/lovdata.no/regjeringen.no var ikke nåbare fra dette kjøremiljøet i denne "
            "kjøringen (se Datakvalitet/nettverksprobe), og det juridiske grunnlaget for å bruke "
            "arbeidsstedets kommune som AGA-basis for et bemanningsforetak er ikke bekreftet mot "
            "primærkilde (se Regelverksgrunnlag). Alle sone-/satsfunn er derfor merket "
            "'GEOGRAFISK ANALYSEGRUNNLAG - MÅ AVKLARES MED LØNN', ikke som endelig AGA.",
        ),
        ("PROSJEKTER SOM KREVER MANUELL KOMMUNEKONTROLL", str(maa_kvalitetssikres_kommuner)),
        (
            "FORELØPIG BEREGNING",
            "Alle AGA-beløp i denne rapporten er foreløpige beregninger basert på geografisk "
            "analysegrunnlag og MÅ kontrolleres av lønn før de legges til grunn.",
        ),
        (
            "VERIFISERTE FUNN",
            "Kolonnemapping, datokonvertering, linjeklassifisering, prosjektregister og "
            "postnummer-basert kommuneidentifikasjon er utført deterministisk fra RecMan-dataene "
            "og kan etterprøves i arkene Kolonnemapping, Prosjektregister og Kommuneoppslag.",
        ),
        (
            "FORELØPIGE BEREGNINGER",
            "Foreløpig AGA-grunnlag og foreløpig beregnet AGA per ansatt/prosjekt/kommune/sone/"
            "termin, se AGA_per_*-arkene.",
        ),
        (
            "UAVKLARTE FORHOLD",
            "(1) Om arbeidsstedets kommune er riktig AGA-basis for et bemanningsforetak (utleie av "
            "arbeidskraft). (2) Presis sonedeling for enkelte sammenslåtte kommuner uten postnummer. "
            "(3) Om sektorunntak (finans, konserntjenester mv.) kan gjelde. Se Regelverksgrunnlag.",
        ),
        (
            "ANBEFALT VIDERE OPPFØLGING",
            "1) Få bekreftet av lønn/juridisk om prosjektkommune er riktig AGA-basis for utleid "
            "arbeidskraft. 2) Åpne og verifisere Skatteetatens kildesider direkte fra et miljø med "
            "nettverkstilgang. 3) Kvalitetssikre prosjektene i Kommuneoppslag/Prosjektregister som "
            "ikke har status KOMMUNE_VERIFISERT.",
        ),
    ]
    report.skriv_lederoppsummering(wb, "Lederoppsummering", lederlinjer)

    banner = FORELOPIG_BEREGNING_BANNER
    beloep_kolonner_termin = {
        "arbeidstimer": "heltall", "lonn_uten_sosial_kost": "beloep", "sosial_kost": "beloep",
        "lonn_inkl_sosial_kost": "beloep", "foreloepig_aga_grunnlag": "beloep",
        "foreloepig_beregnet_aga": "beloep", "differanse_mot_sammenligningssats": "beloep",
    }
    report.skriv_tabellark(wb, "AGA_per_termin", agg_termin, beloep_kolonner_termin, undertittel=banner)
    report.skriv_tabellark(wb, "AGA_per_kommune", agg_kommune, beloep_kolonner_termin, undertittel=banner)
    report.skriv_tabellark(wb, "AGA_per_prosjekt", agg_prosjekt, beloep_kolonner_termin, undertittel=banner)
    report.skriv_tabellark(wb, "AGA_per_ansatt", agg_ansatt, beloep_kolonner_termin, undertittel=banner)

    detalj_kolonner = [
        "arbeidsdato", "aar", "maaned", "termin", "ansattnr", "ansatt", "bedrift", "prosjektnr",
        "prosjekt", "jobbnr", "jobb", "postnummer_normalisert", "kommune", "kommunenummer", "sone",
        "ordinaer_sats", "aga_kontrollstatus", "aga_kontrollkommentar", "artikkeltype", "artikkel",
        "lonnskomponent", "er_tilleggslinje", "er_korreksjon", "arbeidstid_fallback_brukt", "arbeidstimer", "lonn_uten_sosial_kost",
        "sosial_kost", "lonn_inkl_sosial_kost", "foreloepig_aga_belop", "aga_kilde_id", "vakt_id",
    ]
    detalj_df = df[[c for c in detalj_kolonner if c in df.columns]].rename(columns={
        "arbeidsdato": "Arbeidsdato", "aar": "År", "maaned": "Måned", "termin": "Termin",
        "ansattnr": "Ansattnr", "ansatt": "Ansatt", "bedrift": "Bedrift", "prosjektnr": "Prosjektnr",
        "prosjekt": "Prosjekt", "jobbnr": "Jobbnr", "jobb": "Jobb",
        "postnummer_normalisert": "Postnummer", "kommune": "Kommune", "kommunenummer": "Kommunenummer",
        "sone": "AGA-sone", "ordinaer_sats": "AGA-sats (ordinær)", "aga_kontrollstatus": "AGA kontrollstatus",
        "aga_kontrollkommentar": "AGA kontrollkommentar", "artikkeltype": "Artikkeltype",
        "artikkel": "Artikkel", "lonnskomponent": "Lønnskomponent", "er_tilleggslinje": "Er_tilleggslinje",
        "er_korreksjon": "Er_korreksjon",
        "arbeidstid_fallback_brukt": "Arbeidstid hentet fra Helg/Helligdag (ikke egen Arbeidstimer-rad)",
        "arbeidstimer": "Arbeidstimer",
        "lonn_uten_sosial_kost": "Lønn_uten_sosial_kost", "sosial_kost": "Sosial_kost",
        "lonn_inkl_sosial_kost": "Lønn_inkl_sosial_kost",
        "foreloepig_aga_belop": "Foreløpig AGA-beløp (MÅ KONTROLLERES AV LØNN)",
        "aga_kilde_id": "AGA kilde-ID (se Kilderegister)", "vakt_id": "Teknisk vakt-ID",
    })
    report.skriv_tabellark(
        wb, "Detaljgrunnlag", detalj_df,
        kolonnetyper={
            "Arbeidsdato": "dato", "Arbeidstimer": "heltall", "AGA-sats (ordinær)": "prosent",
            "Lønn_uten_sosial_kost": "beloep", "Sosial_kost": "beloep", "Lønn_inkl_sosial_kost": "beloep",
            "Foreløpig AGA-beløp (MÅ KONTROLLERES AV LØNN)": "beloep",
        },
        status_kolonner=["AGA kontrollstatus"],
        undertittel=banner,
    )

    prosjekt_rader = []
    kommuneoppslag_rader = []
    for p in prosjektregister:
        prosjekt_rader.append({
            "Nøkkeltype": p.noekkeltype, "Nøkkelverdi": p.noekkelverdi, "Prosjektnr": p.prosjektnr,
            "Jobbnr": p.jobbnr, "Prosjektnavn": "; ".join(p.prosjektnavn), "Jobbnavn": "; ".join(p.jobbnavn),
            "Bedrift": "; ".join(p.bedriftnavn), "Org.nr": "; ".join(p.org_nr),
            "Postnumre": "; ".join(p.postnumre), "Avdelinger": "; ".join(p.avdelinger),
            "Antall rader": p.antall_rader, "Flere postnumre": p.flere_postnumre,
            "Flere kommuner mulig": p.flere_kommuner_mulig, "Samme nr flere navn": p.samme_nr_flere_navn,
            "Samme navn flere nr": p.samme_navn_flere_nr,
            "Postnummer mangler/ugyldig": p.postnummer_mangler_eller_ugyldig,
        })
        kommuneoppslag_rader.append({
            "Prosjektnr": p.prosjektnr, "Jobbnr": p.jobbnr, "Bedrift": "; ".join(p.bedriftnavn),
            "Identifisert kommune": p.kommune_resultat.kommune, "Kommunenummer": p.kommune_resultat.kommunenummer,
            "Benyttet postnummer": p.kommune_resultat.benyttet_postnummer,
            "Identifikasjonsmetode": p.kommune_resultat.identifikasjonsmetode,
            "Identifikasjonskilde": p.kommune_resultat.identifikasjonskilde,
            "Kontrollstatus": p.kommune_resultat.kontrollstatus,
            "Kontrollkommentar": p.kommune_resultat.kontrollkommentar,
        })
    report.skriv_tabellark(wb, "Prosjektregister", pd.DataFrame(prosjekt_rader))
    report.skriv_tabellark(
        wb, "Kommuneoppslag", pd.DataFrame(kommuneoppslag_rader), status_kolonner=["Kontrollstatus"]
    )

    kilde_df = pd.DataFrame([
        {
            "Kilde-ID": r.kilde_id, "Kildetittel": r.kildetittel, "Kilde-URL": r.kildeadresse,
            "Offentlig utgiver": r.offentlig_utgiver, "Dokumenttype": r.dokumenttype,
            "Relevant år": r.relevant_aar, "Publisert dato": r.publisert_dato,
            "Sist oppdatert dato": r.sist_oppdatert_dato, "Kontrollert tidspunkt": r.kontrollert_tidspunkt,
            "Underbygger": r.underbygger, "Sitat/henvisning": r.sitat_henvisning,
            "Kontrollstatus": r.kontrollstatus, "Merknad": r.merknad, "SHA-256 (lokal kopi)": r.sha256,
        }
        for r in kilderader
    ])
    report.skriv_tabellark(wb, "Kilderegister", kilde_df, status_kolonner=["Kontrollstatus"])

    regelverk_linjer = [
        ("SPØRSMÅL SOM MÅ AVKLARES MED LØNN/JURIDISK", ""),
        (
            "1. Hvordan fastsettes geografisk differensiering av AGA?",
            "Ikke bekreftet mot primærkilde i denne kjøringen (nettverksblokkering mot skatteetaten.no/"
            "lovdata.no/regjeringen.no - se Kilderegister-raden REGELVERK-AGA-2026).",
        ),
        (
            "2. Hvilken betydning har registrert underenhet (Enhetsregisteret)?",
            "Ikke bekreftet. Vanlig praksis for de fleste virksomheter er at AGA-sone følger "
            "underenhetens registrerte kommune, men dette er IKKE lagt til grunn i denne rapporten "
            "uten bekreftelse, siden datagrunnlaget mangler informasjon om Agito Norge AS sine "
            "registrerte underenheter per kommune.",
        ),
        (
            "3. Hvilken betydning har faktisk arbeidssted?",
            "Denne rapporten bruker arbeidsstedets kommune (via postnummer/bedriftsnavn) som "
            "GEOGRAFISK ANALYSEGRUNNLAG - ikke fordi dette er bekreftet som riktig AGA-basis, men "
            "fordi det er den eneste geografiske informasjonen tilgjengelig i datagrunnlaget.",
        ),
        (
            "4. Regler ved utleie av arbeidskraft (bemanningsbransjen)?",
            "Ikke bekreftet mot primærkilde. Dette er trolig det mest kritiske åpne spørsmålet for "
            "Agito Norge AS, som er et bemanningsforetak - må avklares med lønn/juridisk før noen "
            "AGA-sone i denne rapporten legges til grunn.",
        ),
        (
            "5. Finnes særregler, unntak eller krav til fordeling?",
            "Ikke bekreftet. Sektorunntak (f.eks. finans/forsikring, konserninterne "
            "hovedkontortjenester) er kjent å eksistere i regelverket generelt, men er ikke "
            "kontrollert mot 2026-teksten for helse-/omsorgstjenester spesifikt.",
        ),
        (
            "6. Hvilket grunnlag gjelder for terminvis rapportering?",
            "Ikke bekreftet. Terminene i denne rapporten (jf. config.json) følger den vanlige "
            "seks-terminordningen for arbeidsgiveravgift, men er ikke kryssjekket mot en åpnet kilde "
            "i denne kjøringen.",
        ),
        (
            "7. Hvilke regler og satser gjelder spesifikt for 2026?",
            "Satstabellen (sone I-V, fribeløpsregel for sone Ia) er mottatt fra oppdragsgiver som et "
            "utdrag av Skatteetatens egen 2026-side, se Kilderegister SATS-2026 og "
            "source_archive/skatteetaten_satser_2026_brukeroppgitt.txt.",
        ),
        (
            "KONKLUSJON",
            "Siden det offisielle regelverket ikke er bekreftet å støtte bruk av prosjektkommunen "
            "som AGA-grunnlag for et bemanningsforetak, presenterer IKKE denne rapporten "
            "sone-/satsfunnene som endelig korrekt AGA. Betegnelsen "
            "'GEOGRAFISK ANALYSEGRUNNLAG - MÅ AVKLARES MED LØNN' er brukt konsekvent i stedet.",
        ),
    ]
    report.skriv_lederoppsummering(wb, "Regelverksgrunnlag", regelverk_linjer)

    report.skriv_tabellark(wb, "Avvik", pd.DataFrame(avvik))
    report.skriv_tabellark(wb, "Datakvalitet", pd.DataFrame(datakvalitet))

    kolonnemapping_df = pd.DataFrame(kolonnemapping.dokumentasjonstabell())
    report.skriv_tabellark(wb, "Kolonnemapping", kolonnemapping_df)

    metode_linjer = [
        ("KOLONNEMAPPING", "Se ark Kolonnemapping. Original -> standardfelt, normalisert på store/små bokstaver, punktum, bindestrek og mellomrom."),
        ("DATORENSING", "Arbeidsdato = kolonnen 'Dato' (fallback: 'Første dag'). Aldri fakturadato/registreringsdato. Excel-serienummer og tekst-datoer tolkes defensivt, se aga_lib/dates.py."),
        ("PROSJEKTIDENTIFIKASJON", "Prosjektnr primær nøkkel, Jobbnr sekundær nøkkel. Prosjektnavn brukes aldri alene som nøkkel når et nummer finnes."),
        ("KOMMUNEIDENTIFIKASJON", "Prioritet: 1) postnummer (kuratert referanse), 2-4) kommunenavn i Bedrift/Prosjekt/Jobb, 5) manuelt kuratert unntak, 6) manuell kontroll. Se aga_lib/municipality.py."),
        ("AGA-KLASSIFISERING", "Geografisk analysegrunnlag, IKKE endelig AGA - se Regelverksgrunnlag. Delte kommuner løses via postnummer der mulig."),
        ("KILDEHIERARKI", "Skatteetaten > Lovdata > Regjeringen.no > SSB/Kartverket. Se Kilderegister."),
        ("TILLEGGSLINJER", "Artikler klassifisert som arbeidstid/tillegg/korreksjon i config.json. Kun arbeidstid-artikler telles som Arbeidstimer - tillegg legges ikke oppå."),
        ("VAKT-ID", "Konstruert av Ansattnr+Arbeidsdato+Jobbnr+Prosjektnr+Fra kl.+Til kl. - IKKE en original RecMan-ID."),
        ("ØKONOMISKE BEREGNINGER", "Lønn_uten_sosial_kost = Total lønn. Sosial_kost fra egen kolonne. Foreløpig AGA = Lønn_uten_sosial_kost * ordinær sats for sonen, kun der sone er kjent."),
        ("USIKKERHETER OG BEGRENSNINGER", "Se Regelverksgrunnlag og Avvik. Nettverkstilgang til offisielle kilder var ikke tilgjengelig i denne kjøringen."),
        ("MANUELLE KONTROLLPUNKTER", "Kommuneoppslag/Prosjektregister-rader uten status KOMMUNE_VERIFISERT, og alle rader i Avvik."),
        ("GJENKJØRING NESTE LØNNSPERIODE", "Oppdater config.json (analyseaar), legg ny kildefil i .\\input, oppdater source_archive med ny kommunekatalog/satstabell, kjør run_aga_analyse.ps1 på nytt."),
    ]
    report.skriv_lederoppsummering(wb, "Metode", metode_linjer)

    kjoremetadata_linjer = [
        ("Kjøredato og klokkeslett", metadata.kjoredato_tidspunkt),
        ("Python-versjon", metadata.python_versjon),
        ("Kildefil", metadata.kildefil_navn),
        ("SHA-256 (kildefil)", metadata.kildefil_sha256),
        ("Antall innleste rader", str(metadata.antall_innleste_rader)),
        ("Antall analyserte rader (analyseår)", str(metadata.antall_analyserte_rader)),
        ("Antall ekskluderte rader (andre år)", str(metadata.antall_ekskluderte_rader)),
        ("Antall unike ansatte", str(metadata.antall_unike_ansatte)),
        ("Antall unike prosjekter", str(metadata.antall_unike_prosjekter)),
        ("Antall identifiserte kommuner", str(metadata.antall_identifiserte_kommuner)),
        ("Programversjon", metadata.programversjon),
    ] + [(f"Pakkeversjon: {p}", v) for p, v in metadata.pakkeversjoner.items()
    ] + [(f"Nettverksprobe: {p.vert}", f"{'Nåbar' if p.naaadd else 'IKKE nåbar'} ({p.detaljer})") for p in nettverksprobe]
    report.skriv_lederoppsummering(wb, "Kjøremetadata", kjoremetadata_linjer)


def main() -> int:
    prosjektmappe = Path.cwd()
    try:
        output_sti = kjoer_analyse(prosjektmappe)
    except KildefilIkkeFunnet as e:
        print(f"\nFEIL: {e}\n", file=sys.stderr)
        return 2
    except KritiskFeil as e:
        print(f"\nFEIL: {e}\n", file=sys.stderr)
        return 3
    except Exception:
        print("\nUVENTET FEIL under AGA-analysen:\n", file=sys.stderr)
        traceback.print_exc()
        return 1

    print("\nAGA-analysen er fullført.")
    print(f"Resultatfil: {output_sti.resolve()}\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
