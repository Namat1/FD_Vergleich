import io
from dataclasses import dataclass
from typing import Dict, List, Set, Tuple

import pandas as pd
import streamlit as st


DEFAULT_SAP_TEXT = """213568
112681
214289
213458
218601
218804
214321
218801
12823
214296
214043
12923
214192
218607
214590
210455
214001
213406
214238
214109
210353
211152
217253
210750
210716
214588
214487
218394
210399
214015
210492
218418
211288
211399
213095
218390
211292
218373
218344
213016
210234
210276
218466
218411
218420
218426
218425
218468
218421
214285
214299
214297
214290
218200
218711
218461
210655
210765
218355
210701
213840
218208
211025
214094
210509
213580
218707
214376
211380
218867
213553
115339
215634
216425
216442
216467
216496
216630
216133
216432
216815
216466
216615
219545
219430
216590
215632
216144
216153
219208
216207
216464
216529
216570
216572
216586
216588
216628
216637
216744
219439
216656
215551
219544
216799
216774
216122
216177
216185
216221
216248
216253
216670
216672
219513
216010
216178
216655
216697
216853
216653
216791
216227
216290
216814
216828
219427
219570
216793
216617
215014
215180
216070
219586
216155
216569
216405
216623
219532
219501
210650
216371"""

DAY_NAMES = {
    1: "Montag",
    2: "Dienstag",
    3: "Mittwoch",
    4: "Donnerstag",
    5: "Freitag",
    6: "Samstag",
}

DAY_COLUMNS_FILE2 = {
    1: 6,   # Spalte G
    2: 7,   # Spalte H
    3: 8,   # Spalte I
    4: 9,   # Spalte J
    5: 10,  # Spalte K
    6: 11,  # Spalte L
}


@dataclass
class CompareResult:
    zusammenfassung: pd.DataFrame
    tagesabweichungen: pd.DataFrame
    nicht_gefunden: pd.DataFrame
    datei1_extrakt: pd.DataFrame
    datei2_extrakt: pd.DataFrame
    matrix: pd.DataFrame


def normalisiere_sap_nummer(wert) -> str:
    if pd.isna(wert):
        return ""

    text = str(wert).strip()
    if text == "":
        return ""

    try:
        zahl = float(text.replace(",", "."))
        if zahl.is_integer():
            return str(int(zahl))
    except Exception:
        pass

    return text


def parse_sap_liste(text: str) -> List[str]:
    if not text:
        return []

    teile = text.replace(",", " ").replace(";", " ").split()
    ergebnis: List[str] = []
    gesehen: Set[str] = set()

    for teil in teile:
        sap = normalisiere_sap_nummer(teil)
        if sap and sap not in gesehen:
            gesehen.add(sap)
            ergebnis.append(sap)

    return ergebnis


def ist_zahl_wert(wert) -> bool:
    if pd.isna(wert):
        return False

    text = str(wert).strip()
    if text == "":
        return False

    try:
        float(text.replace(",", "."))
        return True
    except Exception:
        return False


def lief_tag_aus_wert(wert) -> int:
    if not ist_zahl_wert(wert):
        return 0

    try:
        tag = int(float(str(wert).strip().replace(",", ".")))
        if 1 <= tag <= 6:
            return tag
    except Exception:
        pass

    return 0


def lese_datei_1(datei, startzeile: int, erlaubte_saps: Set[str]) -> Tuple[Dict[str, Set[int]], pd.DataFrame]:
    arbeitsmappe = pd.ExcelFile(datei)
    erstes_blatt = arbeitsmappe.sheet_names[0]
    daten = pd.read_excel(arbeitsmappe, sheet_name=erstes_blatt, header=None)

    lieferungen: Dict[str, Set[int]] = {}
    extrakt_zeilen: List[dict] = []
    gesehen: Set[Tuple[str, int]] = set()

    for index in range(max(startzeile - 1, 0), len(daten)):
        sap = normalisiere_sap_nummer(daten.iat[index, 0] if daten.shape[1] > 0 else None)
        if not sap or (erlaubte_saps and sap not in erlaubte_saps):
            continue

        tag = lief_tag_aus_wert(daten.iat[index, 6] if daten.shape[1] > 6 else None)
        if tag == 0:
            continue

        schluessel = (sap, tag)
        if schluessel in gesehen:
            continue

        gesehen.add(schluessel)
        lieferungen.setdefault(sap, set()).add(tag)
        extrakt_zeilen.append(
            {
                "SAP-Nummer": sap,
                "Liefertag Nummer": tag,
                "Liefertag": DAY_NAMES[tag],
                "Blatt": erstes_blatt,
                "Zeile": index + 1,
                "Fundstelle": f"{erstes_blatt} Zeile {index + 1}",
            }
        )

    extrakt = pd.DataFrame(extrakt_zeilen)
    if not extrakt.empty:
        extrakt = extrakt.sort_values(["SAP-Nummer", "Liefertag Nummer"]).reset_index(drop=True)

    return lieferungen, extrakt


def lese_datei_2(datei, startzeile: int, erlaubte_saps: Set[str], anzahl_blaetter: int) -> Tuple[Dict[str, Set[int]], pd.DataFrame]:
    arbeitsmappe = pd.ExcelFile(datei)
    blattnamen = arbeitsmappe.sheet_names[:anzahl_blaetter]

    lieferungen: Dict[str, Set[int]] = {}
    extrakt_zeilen: List[dict] = []
    gesehen: Set[Tuple[str, int]] = set()

    for blattname in blattnamen:
        daten = pd.read_excel(arbeitsmappe, sheet_name=blattname, header=None)

        for index in range(max(startzeile - 1, 0), len(daten)):
            sap = normalisiere_sap_nummer(daten.iat[index, 1] if daten.shape[1] > 1 else None)
            if not sap or (erlaubte_saps and sap not in erlaubte_saps):
                continue

            for tag, spalten_index in DAY_COLUMNS_FILE2.items():
                wert = daten.iat[index, spalten_index] if daten.shape[1] > spalten_index else None
                if not ist_zahl_wert(wert):
                    continue

                schluessel = (sap, tag)
                if schluessel not in gesehen:
                    gesehen.add(schluessel)
                    lieferungen.setdefault(sap, set()).add(tag)

                extrakt_zeilen.append(
                    {
                        "SAP-Nummer": sap,
                        "Liefertag Nummer": tag,
                        "Liefertag": DAY_NAMES[tag],
                        "Blatt": blattname,
                        "Zeile": index + 1,
                        "Zellenwert": wert,
                        "Fundstelle": f"{blattname} Zeile {index + 1}",
                    }
                )

    extrakt = pd.DataFrame(extrakt_zeilen)
    if not extrakt.empty:
        extrakt = extrakt.sort_values(["SAP-Nummer", "Liefertag Nummer", "Blatt", "Zeile"]).reset_index(drop=True)

    return lieferungen, extrakt


def lief_tags_text(tags: Set[int]) -> str:
    if not tags:
        return ""
    return ", ".join([f"{tag} {DAY_NAMES[tag]}" for tag in sorted(tags)])


def fundstellen_map(extrakt: pd.DataFrame) -> Dict[Tuple[str, int], str]:
    if extrakt.empty:
        return {}

    sammlung: Dict[Tuple[str, int], List[str]] = {}
    for _, zeile in extrakt.iterrows():
        schluessel = (str(zeile["SAP-Nummer"]), int(zeile["Liefertag Nummer"]))
        sammlung.setdefault(schluessel, [])
        fundstelle = str(zeile["Fundstelle"])
        if fundstelle not in sammlung[schluessel]:
            sammlung[schluessel].append(fundstelle)

    return {schluessel: " | ".join(fundstellen) for schluessel, fundstellen in sammlung.items()}


def baue_matrix(alle_saps: List[str], datei1_map: Dict[str, Set[int]], datei2_map: Dict[str, Set[int]]) -> pd.DataFrame:
    zeilen: List[dict] = []
    for sap in alle_saps:
        zeile = {"SAP-Nummer": sap}
        for tag in range(1, 7):
            zeile[f"{DAY_NAMES[tag]} Datei 1"] = "Ja" if tag in datei1_map.get(sap, set()) else "Nein"
            zeile[f"{DAY_NAMES[tag]} Datei 2"] = "Ja" if tag in datei2_map.get(sap, set()) else "Nein"
        zeilen.append(zeile)

    matrix = pd.DataFrame(zeilen)
    return matrix


def vergleiche_dateien(
    datei1,
    datei2,
    sap_liste: List[str],
    startzeile_datei1: int,
    startzeile_datei2: int,
    anzahl_blaetter_datei2: int,
) -> CompareResult:
    erlaubte_saps = set(sap_liste)

    datei1_map, datei1_extrakt = lese_datei_1(datei1, startzeile_datei1, erlaubte_saps)
    datei2_map, datei2_extrakt = lese_datei_2(datei2, startzeile_datei2, erlaubte_saps, anzahl_blaetter_datei2)

    fundstellen_datei1 = fundstellen_map(datei1_extrakt)
    fundstellen_datei2 = fundstellen_map(datei2_extrakt)

    zusammenfassung_zeilen: List[dict] = []
    tagesabweichung_zeilen: List[dict] = []
    nicht_gefunden_zeilen: List[dict] = []

    for sap in sap_liste:
        tags1 = datei1_map.get(sap, set())
        tags2 = datei2_map.get(sap, set())

        hat1 = len(tags1) > 0
        hat2 = len(tags2) > 0

        if not hat1 and not hat2:
            status = "In keiner Datei gefunden"
            nicht_gefunden_zeilen.append({"SAP-Nummer": sap, "Hinweis": status})
        elif tags1 != tags2:
            if not hat1:
                status = "Nur in Datei 2 vorhanden"
            elif not hat2:
                status = "Nur in Datei 1 vorhanden"
            else:
                status = "Liefertage unterschiedlich"
        else:
            status = "Gleich"

        zusammenfassung_zeilen.append(
            {
                "SAP-Nummer": sap,
                "Liefertage Datei 1": lief_tags_text(tags1),
                "Liefertage Datei 2": lief_tags_text(tags2),
                "Status": status,
                "Anzahl Liefertage Datei 1": len(tags1),
                "Anzahl Liefertage Datei 2": len(tags2),
            }
        )

        for tag in range(1, 7):
            in_datei1 = tag in tags1
            in_datei2 = tag in tags2
            if in_datei1 == in_datei2:
                continue

            if in_datei1 and not in_datei2:
                detail_status = "Nur in Datei 1 vorhanden"
            else:
                detail_status = "Nur in Datei 2 vorhanden"

            tagesabweichung_zeilen.append(
                {
                    "SAP-Nummer": sap,
                    "Liefertag Nummer": tag,
                    "Liefertag": DAY_NAMES[tag],
                    "Status": detail_status,
                    "Fundstelle Datei 1": fundstellen_datei1.get((sap, tag), ""),
                    "Fundstelle Datei 2": fundstellen_datei2.get((sap, tag), ""),
                }
            )

    zusammenfassung = pd.DataFrame(zusammenfassung_zeilen)
    if not zusammenfassung.empty:
        status_reihenfolge = {
            "Liefertage unterschiedlich": 1,
            "Nur in Datei 1 vorhanden": 2,
            "Nur in Datei 2 vorhanden": 3,
            "In keiner Datei gefunden": 4,
            "Gleich": 5,
        }
        zusammenfassung["_sort"] = zusammenfassung["Status"].map(status_reihenfolge)
        zusammenfassung = zusammenfassung.sort_values(["_sort", "SAP-Nummer"]).drop(columns=["_sort"]).reset_index(drop=True)

    tagesabweichungen = pd.DataFrame(tagesabweichung_zeilen)
    if not tagesabweichungen.empty:
        tagesabweichungen = tagesabweichungen.sort_values(["SAP-Nummer", "Liefertag Nummer"]).reset_index(drop=True)

    nicht_gefunden = pd.DataFrame(nicht_gefunden_zeilen)
    if not nicht_gefunden.empty:
        nicht_gefunden = nicht_gefunden.sort_values(["SAP-Nummer"]).reset_index(drop=True)

    matrix = baue_matrix(sap_liste, datei1_map, datei2_map)

    return CompareResult(
        zusammenfassung=zusammenfassung,
        tagesabweichungen=tagesabweichungen,
        nicht_gefunden=nicht_gefunden,
        datei1_extrakt=datei1_extrakt,
        datei2_extrakt=datei2_extrakt,
        matrix=matrix,
    )


def exportiere_bericht(result: CompareResult) -> bytes:
    ausgabe = io.BytesIO()
    with pd.ExcelWriter(ausgabe, engine="openpyxl") as writer:
        result.zusammenfassung.to_excel(writer, index=False, sheet_name="Zusammenfassung")
        result.tagesabweichungen.to_excel(writer, index=False, sheet_name="Tagesabweichungen")
        result.nicht_gefunden.to_excel(writer, index=False, sheet_name="Nicht gefunden")
        result.matrix.to_excel(writer, index=False, sheet_name="Matrix")
        result.datei1_extrakt.to_excel(writer, index=False, sheet_name="Extrakt Datei 1")
        result.datei2_extrakt.to_excel(writer, index=False, sheet_name="Extrakt Datei 2")

        arbeitsmappe = writer.book
        for blatt in arbeitsmappe.worksheets:
            for spalte in blatt.columns:
                max_laenge = 0
                spaltenbuchstabe = spalte[0].column_letter
                for zelle in spalte:
                    try:
                        inhalt = "" if zelle.value is None else str(zelle.value)
                    except Exception:
                        inhalt = ""
                    max_laenge = max(max_laenge, len(inhalt))
                blatt.column_dimensions[spaltenbuchstabe].width = min(max(max_laenge + 2, 12), 42)
            blatt.freeze_panes = "A2"

    ausgabe.seek(0)
    return ausgabe.getvalue()


def zeige_kennzahlen(result: CompareResult, anzahl_saps: int):
    anzahl_gleich = int((result.zusammenfassung["Status"] == "Gleich").sum()) if not result.zusammenfassung.empty else 0
    anzahl_abweichend = int((result.zusammenfassung["Status"] != "Gleich").sum()) if not result.zusammenfassung.empty else 0
    anzahl_nicht_gefunden = len(result.nicht_gefunden)
    anzahl_tagesabweichungen = len(result.tagesabweichungen)

    spalte1, spalte2, spalte3, spalte4 = st.columns(4)
    spalte1.metric("Ausgewählte Kunden", anzahl_saps)
    spalte2.metric("Kunden ohne Abweichung", anzahl_gleich)
    spalte3.metric("Kunden mit Abweichung", anzahl_abweichend)
    spalte4.metric("Tagesabweichungen", anzahl_tagesabweichungen)

    st.caption(f"Nicht gefundene Kunden: {anzahl_nicht_gefunden}")


def hauptseite():
    st.set_page_config(page_title="Liefertage Vergleich", layout="wide")
    st.title("Liefertage Vergleich als Streamlit Auswertung")
    st.write(
        "Lade zwei Excel-Dateien hoch. Die Anwendung liest aus Datei 1 die SAP-Nummer aus Spalte A "
        "und den Liefertag aus Spalte G. Aus Datei 2 werden die ersten Blätter gelesen, dort ist die "
        "SAP-Nummer in Spalte B und Montag bis Samstag stehen in Spalte G bis L."
    )

    with st.sidebar:
        st.header("Einstellungen")
        startzeile_datei1 = st.number_input("Startzeile Datei 1", min_value=1, value=2, step=1)
        startzeile_datei2 = st.number_input("Startzeile Datei 2", min_value=1, value=2, step=1)
        anzahl_blaetter = st.number_input("Anzahl Blätter in Datei 2", min_value=1, max_value=20, value=4, step=1)
        nur_abweichungen = st.checkbox("In Tabellen zuerst nur Abweichungen zeigen", value=True)

    spalte_links, spalte_rechts = st.columns(2)
    with spalte_links:
        datei1 = st.file_uploader("Datei 1 hochladen", type=["xlsx", "xlsm", "xls"])
    with spalte_rechts:
        datei2 = st.file_uploader("Datei 2 hochladen", type=["xlsx", "xlsm", "xls"])

    sap_text = st.text_area(
        "Zu vergleichende SAP-Nummern",
        value=DEFAULT_SAP_TEXT,
        height=260,
        help="Eine SAP-Nummer pro Zeile oder mit Leerzeichen getrennt einfügen.",
    )

    sap_liste = parse_sap_liste(sap_text)
    st.info(f"Es werden aktuell {len(sap_liste)} SAP-Nummern verglichen.")

    if not datei1 or not datei2:
        st.warning("Bitte beide Excel-Dateien hochladen.")
        return

    if not sap_liste:
        st.error("Bitte mindestens eine SAP-Nummer eintragen.")
        return

    with st.spinner("Dateien werden gelesen und verglichen..."):
        result = vergleiche_dateien(
            datei1=datei1,
            datei2=datei2,
            sap_liste=sap_liste,
            startzeile_datei1=int(startzeile_datei1),
            startzeile_datei2=int(startzeile_datei2),
            anzahl_blaetter_datei2=int(anzahl_blaetter),
        )

    zeige_kennzahlen(result, len(sap_liste))

    tabs = st.tabs([
        "Zusammenfassung",
        "Tagesabweichungen",
        "Nicht gefunden",
        "Matrix",
        "Extrakt Datei 1",
        "Extrakt Datei 2",
    ])

    with tabs[0]:
        daten = result.zusammenfassung.copy()
        if nur_abweichungen:
            daten = daten[daten["Status"] != "Gleich"].reset_index(drop=True)
        st.dataframe(daten, use_container_width=True, hide_index=True)

        if not result.zusammenfassung.empty:
            status_werte = result.zusammenfassung["Status"].value_counts().rename_axis("Status").reset_index(name="Anzahl")
            st.subheader("Verteilung nach Status")
            st.bar_chart(status_werte.set_index("Status"))

    with tabs[1]:
        st.dataframe(result.tagesabweichungen, use_container_width=True, hide_index=True)
        if not result.tagesabweichungen.empty:
            tage_werte = result.tagesabweichungen["Liefertag"].value_counts().reindex(list(DAY_NAMES.values()), fill_value=0)
            st.subheader("Abweichungen pro Liefertag")
            st.bar_chart(tage_werte)

    with tabs[2]:
        st.dataframe(result.nicht_gefunden, use_container_width=True, hide_index=True)

    with tabs[3]:
        st.dataframe(result.matrix, use_container_width=True, hide_index=True)

    with tabs[4]:
        st.dataframe(result.datei1_extrakt, use_container_width=True, hide_index=True)

    with tabs[5]:
        st.dataframe(result.datei2_extrakt, use_container_width=True, hide_index=True)

    excel_bytes = exportiere_bericht(result)
    st.download_button(
        label="Bericht als Excel herunterladen",
        data=excel_bytes,
        file_name="liefertage_vergleich_auswertung.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    hauptseite()
