"""STEG 4-9: seks frittstående leveransefiler i tillegg til AGA_Rapport_<år>.xlsx.

Gjenbruker data som allerede er beregnet av kjoer_analyse() i aga_analyse.py
(prosjektregister, kommune-/AGA-klassifisering, kilderegister, avvik) - ingen
ny innlesing eller nye nettverksoppslag gjøres her.

Filene som produseres i .\\output:

1. Prosjektregister.xlsx   - Prosjektnr | Prosjekt | Kommune | AGA-sone
2. AGA-detaljrapport.xlsx  - Ansatt | Prosjekt | Kommune | AGA-sone | Timer | Total lønn
3. AGA-oppsummering.xlsx   - AGA-sone | Kommune | Timer | Total lønn
4. Kilderegister.xlsx      - Kilde-ID | Kommune | AGA-sone | Kilde | URL | Utgiver | Kontrollert dato | Status
5. Avviksrapport.xlsx      - kategoriserte avvik (seksjon 8)
6. Lederoppsummering.md    - kort, lederorientert oppsummering på norsk

STEG 4 sin regel om at AGA-sone ALDRI skal gjettes håndteres ved at
Kilderegister/AGA-sone-kolonnene er tomme og merket
"MANGLER OFFISIELT KILDEGRUNNLAG" når kilde_tilgjengelig=False (jf. det
kritiske kravet i oppdraget), på samme måte som AGA_oppslagsbehov.xlsx i
hovedrapporten (se aga_lib/oppslagsbehov.py).
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path

import pandas as pd

from aga_lib import report

STATUS_KILDE_IKKE_FUNNET = "KILDE IKKE FUNNET"
STATUS_MANGLER_KILDE_STEG4 = "MANGLER OFFISIELT KILDEGRUNNLAG"


def _unik_sti(sti: Path) -> Path:
    """Returnerer sti uendret dersom filen ikke finnes fra før, ellers en
    tidsstemplet variant - slik at en tidligere leveranse aldri overskrives
    stille (jf. seksjon 22: «Ikke overskriv tidligere rapport uten eksplisitt
    tillatelse»)."""
    if not sti.exists():
        return sti
    tidsstempel = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    return sti.with_name(f"{sti.stem}_{tidsstempel}{sti.suffix}")


def _prosjektnoekkel_for_rad(rad) -> tuple[str, str]:
    prosjektnr = rad.get("prosjektnr")
    if prosjektnr is not None and str(prosjektnr).strip() not in ("", "nan", "None"):
        return ("prosjektnr", str(prosjektnr).strip())
    jobbnr = rad.get("jobbnr")
    if jobbnr is not None and str(jobbnr).strip() not in ("", "nan", "None"):
        return ("jobbnr", str(jobbnr).strip())
    return ("ukjent", f"UKJENT-{rad.get('bedrift', 'ukjent bedrift')}")


def _skriv_enkelt_ark(sti: Path, arknavn: str, df: pd.DataFrame, status_kolonner=None) -> None:
    wb = report.lag_arbeidsbok()
    report.skriv_tabellark(wb, arknavn, df, status_kolonner=status_kolonner)
    report.lagre(wb, sti)


def bygg_og_skriv_leveranser(
    output_mappe: Path,
    df: pd.DataFrame,
    prosjektregister: list,
    aga_per_prosjektnoekkel: dict,
    kilde_id_geografi: str,
    kilde_tilgjengelig: bool,
    kk_url: str,
    kk_tittel: str,
    kk_utgiver: str,
    kontrolltidspunkt: str,
    analyseaar: int,
    logger,
) -> list[Path]:
    filer: list[Path] = []

    # ---------------------------------------------------------------
    # 1) Prosjektregister.xlsx
    # ---------------------------------------------------------------
    rader1 = []
    for p in prosjektregister:
        klass = aga_per_prosjektnoekkel[(p.noekkeltype, p.noekkelverdi)]
        rader1.append({
            "Prosjektnr": p.prosjektnr or p.noekkelverdi,
            "Prosjekt": "; ".join(p.prosjektnavn),
            "Kommune": klass.kommune or "",
            "AGA-sone": klass.sone or "",
        })
    sti1 = _unik_sti(output_mappe / "Prosjektregister.xlsx")
    _skriv_enkelt_ark(sti1, "Prosjektregister", pd.DataFrame(rader1))
    filer.append(sti1)

    # ---------------------------------------------------------------
    # 2) AGA-detaljrapport.xlsx (én rad per registrering/arbeidspost)
    # ---------------------------------------------------------------
    noekler = df.apply(_prosjektnoekkel_for_rad, axis=1)
    detalj = pd.DataFrame({
        "Ansatt": df.get("ansatt"),
        "Prosjekt": df.get("prosjekt"),
        "Kommune": [aga_per_prosjektnoekkel[n].kommune or "" for n in noekler],
        "AGA-sone": [aga_per_prosjektnoekkel[n].sone or "" for n in noekler],
        "Timer": df.get("arbeidstimer"),
        "Total lønn": df.get("total_lonn"),
    })
    sti2 = _unik_sti(output_mappe / "AGA-detaljrapport.xlsx")
    _skriv_enkelt_ark(sti2, "AGA-detaljrapport", detalj)
    filer.append(sti2)

    # ---------------------------------------------------------------
    # 3) AGA-oppsummering.xlsx (per AGA-sone og kommune)
    # ---------------------------------------------------------------
    oppsummering_grunnlag = detalj.copy()
    oppsummering_grunnlag["AGA-sone"] = oppsummering_grunnlag["AGA-sone"].replace("", "(ikke fastslått)")
    oppsummering_grunnlag["Kommune"] = oppsummering_grunnlag["Kommune"].replace("", "(ikke fastslått)")
    agg3 = (
        oppsummering_grunnlag.groupby(["AGA-sone", "Kommune"], dropna=False)
        .agg(Timer=("Timer", "sum"), **{"Total lønn": ("Total lønn", "sum")})
        .reset_index()
    )
    sti3 = _unik_sti(output_mappe / "AGA-oppsummering.xlsx")
    _skriv_enkelt_ark(sti3, "AGA-oppsummering", agg3)
    filer.append(sti3)

    # ---------------------------------------------------------------
    # 4) Kilderegister.xlsx (én rad per kommune -> sone -> kilde)
    # ---------------------------------------------------------------
    if not kilde_tilgjengelig:
        logger.warning(
            "Kilderegister.xlsx opprettes TOMT (kun kolonneoverskrifter) - offisiell kilde er ikke "
            "tilgjengelig i denne kjøringen (jf. det kritiske kravet i STEG 4/9)."
        )
        kilde4_df = pd.DataFrame(
            columns=["Kilde-ID", "Kommune", "AGA-sone", "Kilde", "URL", "Utgiver", "Kontrollert dato", "Status"]
        )
    else:
        sett = {}
        for p in prosjektregister:
            klass = aga_per_prosjektnoekkel[(p.noekkeltype, p.noekkelverdi)]
            if klass.kommunenummer is None:
                continue
            sett[klass.kommunenummer] = klass
        rader4 = []
        for kommunenummer, klass in sett.items():
            if klass.sone:
                status = "MÅ KVALITETSSIKRES" if "motstridende" in klass.kontrollkommentar.lower() else "KILDEVERIFISERT"
            else:
                status = STATUS_KILDE_IKKE_FUNNET
            rader4.append({
                "Kilde-ID": kilde_id_geografi,
                "Kommune": klass.kommune or "",
                "AGA-sone": klass.sone or "",
                "Kilde": kk_tittel,
                "URL": kk_url,
                "Utgiver": kk_utgiver,
                "Kontrollert dato": kontrolltidspunkt,
                "Status": status,
            })
        kilde4_df = pd.DataFrame(rader4)
    sti4 = _unik_sti(output_mappe / "Kilderegister.xlsx")
    _skriv_enkelt_ark(sti4, "Kilderegister", kilde4_df, status_kolonner=["Status"] if len(kilde4_df) else None)
    filer.append(sti4)

    # ---------------------------------------------------------------
    # 5) Avviksrapport.xlsx (seksjon 8-kategoriene, eksplisitt)
    # ---------------------------------------------------------------
    def _prosjekter_der(vilkaar) -> str:
        treff = [
            f"{p.prosjektnr or p.noekkelverdi} ({'; '.join(p.prosjektnavn) or '(uten navn)'})"
            for p in prosjektregister if vilkaar(p)
        ]
        return "; ".join(treff)

    avvik5 = [
        {
            "Kategori": "Prosjekter uten kommune",
            "Antall": sum(1 for p in prosjektregister if aga_per_prosjektnoekkel[(p.noekkeltype, p.noekkelverdi)].kommune is None),
            "Berørte prosjekter": _prosjekter_der(lambda p: aga_per_prosjektnoekkel[(p.noekkeltype, p.noekkelverdi)].kommune is None),
        },
        {
            "Kategori": "Kommuner uten AGA-sone",
            "Antall": sum(1 for p in prosjektregister if not aga_per_prosjektnoekkel[(p.noekkeltype, p.noekkelverdi)].sone and aga_per_prosjektnoekkel[(p.noekkeltype, p.noekkelverdi)].kommune),
            "Berørte prosjekter": _prosjekter_der(lambda p: not aga_per_prosjektnoekkel[(p.noekkeltype, p.noekkelverdi)].sone and aga_per_prosjektnoekkel[(p.noekkeltype, p.noekkelverdi)].kommune),
        },
        {
            "Kategori": "Manglende postnummer",
            "Antall": sum(1 for p in prosjektregister if p.postnummer_mangler_eller_ugyldig),
            "Berørte prosjekter": _prosjekter_der(lambda p: p.postnummer_mangler_eller_ugyldig),
        },
        {
            "Kategori": "Motstridende datagrunnlag",
            "Antall": sum(1 for p in prosjektregister if p.kommune_resultat.kontrollstatus == "MOTSTRIDENDE_OPPLYSNINGER"),
            "Berørte prosjekter": _prosjekter_der(lambda p: p.kommune_resultat.kontrollstatus == "MOTSTRIDENDE_OPPLYSNINGER"),
        },
        {
            "Kategori": "Flere kommuner tilknyttet samme prosjekt",
            "Antall": sum(1 for p in prosjektregister if p.flere_postnumre or p.flere_kommuner_mulig),
            "Berørte prosjekter": _prosjekter_der(lambda p: p.flere_postnumre or p.flere_kommuner_mulig),
        },
        {
            "Kategori": "Flere AGA-soner tilknyttet samme prosjekt",
            "Antall": sum(
                1 for p in prosjektregister
                if "delt mellom flere soner" in aga_per_prosjektnoekkel[(p.noekkeltype, p.noekkelverdi)].kontrollkommentar
                or "flere AGA-soner" in aga_per_prosjektnoekkel[(p.noekkeltype, p.noekkelverdi)].kontrollkommentar
            ),
            "Berørte prosjekter": _prosjekter_der(
                lambda p: "delt mellom flere soner" in aga_per_prosjektnoekkel[(p.noekkeltype, p.noekkelverdi)].kontrollkommentar
                or "flere AGA-soner" in aga_per_prosjektnoekkel[(p.noekkeltype, p.noekkelverdi)].kontrollkommentar
            ),
        },
    ]
    sti5 = _unik_sti(output_mappe / "Avviksrapport.xlsx")
    _skriv_enkelt_ark(sti5, "Avviksrapport", pd.DataFrame(avvik5))
    filer.append(sti5)

    # ---------------------------------------------------------------
    # 6) Lederoppsummering.md
    # ---------------------------------------------------------------
    samlet_lonn = pd.to_numeric(df.get("total_lonn"), errors="coerce").fillna(0).sum()
    antall_prosjekter = len(prosjektregister)
    kommuner = sorted({aga_per_prosjektnoekkel[(p.noekkeltype, p.noekkelverdi)].kommune for p in prosjektregister if aga_per_prosjektnoekkel[(p.noekkeltype, p.noekkelverdi)].kommune})
    soner = sorted({aga_per_prosjektnoekkel[(p.noekkeltype, p.noekkelverdi)].sone for p in prosjektregister if aga_per_prosjektnoekkel[(p.noekkeltype, p.noekkelverdi)].sone})

    lonn_per_kommune_serie = (
        pd.DataFrame({"Kommune": [aga_per_prosjektnoekkel[n].kommune or "(ikke fastslått)" for n in noekler],
                       "Total lønn": pd.to_numeric(df.get("total_lonn"), errors="coerce").fillna(0)})
        .groupby("Kommune")["Total lønn"].sum().sort_values(ascending=False)
    )
    topp_kommuner = "\n".join(
        f"- {k}: {v:,.2f} kr".replace(",", " ").replace(".", ",", 1) for k, v in lonn_per_kommune_serie.head(5).items()
    )

    prosjekter_krever_kontroll = [
        f"- {p.prosjektnr or p.noekkelverdi} ({'; '.join(p.prosjektnavn) or '(uten navn)'}): {p.kommune_resultat.kontrollstatus}"
        for p in prosjektregister if p.kommune_resultat.kontrollstatus != "KOMMUNE_VERIFISERT"
    ]

    md = f"""# Lederoppsummering - AGA-analyse {analyseaar}

## 1. Totalt analysert lønnsgrunnlag
{samlet_lonn:,.2f} kr (Total lønn, uten sosial kost).

## 2. Totalt antall prosjekter
{antall_prosjekter}

## 3. Totalt antall kommuner
{len(kommuner)} ({", ".join(kommuner) if kommuner else "ingen identifisert"})

## 4. Antall identifiserte AGA-soner
{len(soner)} ({", ".join(soner) if soner else "ingen - se punkt 7"})

## 5. Kommuner med størst lønnsgrunnlag
{topp_kommuner or "Ingen data"}

## 6. Prosjekter som krever kontroll
{chr(10).join(prosjekter_krever_kontroll) if prosjekter_krever_kontroll else "Ingen - alle prosjekter har KOMMUNE_VERIFISERT."}

## 7. Vesentlige risikoer
- Ingen AGA-sone i denne analysen er endelig juridisk bekreftet: det er ikke avklart mot primærkilde om
  arbeidsstedets kommune (fremfor f.eks. registrert underenhet) er riktig AGA-basis for et
  bemanningsforetak som leier ut arbeidskraft. Se Regelverksgrunnlag i AGA_Rapport-arbeidsboken.
- Skatteetaten.no, lovdata.no og regjeringen.no var ikke nåbare fra dette kjøremiljøet ved kjøretidspunktet
  ({kontrolltidspunkt}) - kildene i Kilderegister.xlsx er derfor ikke selv åpnet og lest i denne kjøringen.
- {"Kommunekatalog/satstabell manglet helt lokalt - AGA-sone er IKKE fylt ut noe sted i denne leveransen." if not kilde_tilgjengelig else "Se Avviksrapport.xlsx for prosjekter med manglende/motstridende kommune- eller sonegrunnlag."}
"""
    sti6 = _unik_sti(output_mappe / "Lederoppsummering.md")
    sti6.write_text(md, encoding="utf-8")
    filer.append(sti6)

    return filer
