import io
from typing import Dict, List, Set, Tuple

import pandas as pd
import streamlit as st

CUSTOMER_GROUPS: Dict[str, List[str]] = {
    "Malchow": [
        "115339", "215634", "216425", "216442", "216467", "216496", "216630", "216133",
        "216432", "216815", "216466", "216615", "219545", "219430", "216590", "215632",
        "216144", "216153", "219208", "216207", "216464", "216529", "216570", "216572",
        "216586", "216588", "216628", "216637", "216744", "219439", "216656", "215551",
        "219544", "216799", "216774", "216122", "216177", "216185", "216221", "216248",
        "216253", "216670", "216672", "219513", "216010", "216178", "216655", "216697",
        "216853", "216653", "216791", "216227", "216290", "216814", "216828", "219427",
        "219570", "216793", "216617", "215014", "215180", "216070", "219586", "216155",
        "216569", "216405", "216623", "219532", "219501", "210650", "216371",
    ],
    "Neumünster": [
        "213406", "214238", "214109", "210353", "211152", "217253", "210750", "210716",
        "214588", "214487", "218394", "210399", "214015", "210492", "218418", "211288",
        "211399", "213095", "218390", "211292", "218373", "218344", "213016", "210234",
        "210276", "218466", "218411", "218420", "218426", "218425", "218468", "218421",
        "214285", "214299", "214297", "214290", "218200", "218711", "218461", "210655",
        "210765", "218355", "210701", "213840", "218208", "211025",
    ],
    "Zarrentin": [
        "213568", "112681", "214289", "213458", "218601", "218804", "214321", "218801",
        "214094", "210509", "213580", "218707", "214376", "211380", "218867", "213553",
        "12823", "214296", "214043", "12923", "214192", "218607", "214590", "210455",
        "214001",
    ],
}

DAY_NAMES = {
    1: "Montag",
    2: "Dienstag",
    3: "Mittwoch",
    4: "Donnerstag",
    5: "Freitag",
    6: "Samstag",
}

DAY_COLUMNS_TOUR = {
    1: 6,
    2: 7,
    3: 8,
    4: 9,
    5: 10,
    6: 11,
}

LOCATION_ORDER = {name: index for index, name in enumerate(CUSTOMER_GROUPS.keys(), start=1)}
CUSTOMER_TO_LOCATION: Dict[str, str] = {}
CUSTOMER_TO_ORDER: Dict[str, int] = {}

for location_name, sap_list in CUSTOMER_GROUPS.items():
    for customer_index, sap_number in enumerate(sap_list, start=1):
        CUSTOMER_TO_LOCATION[sap_number] = location_name
        CUSTOMER_TO_ORDER[sap_number] = customer_index

SELECTED_SAPS: Set[str] = set(CUSTOMER_TO_LOCATION.keys())


def normalize_sap(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text == "":
        return ""
    try:
        number = float(text.replace(",", "."))
        if number.is_integer():
            return str(int(number))
    except Exception:
        pass
    return text


def is_numeric_cell(value) -> bool:
    if pd.isna(value):
        return False
    text = str(value).strip()
    if text == "":
        return False
    try:
        float(text.replace(",", "."))
        return True
    except Exception:
        return False


def read_sap_file(uploaded_file, selected_saps: Set[str]) -> Tuple[Dict[str, Set[int]], str]:
    excel = pd.ExcelFile(uploaded_file)
    sheet_name = excel.sheet_names[0]
    df = pd.read_excel(excel, sheet_name=sheet_name, header=0)

    days_by_sap: Dict[str, Set[int]] = {}

    for idx in range(len(df)):
        sap = normalize_sap(df.iloc[idx, 0] if df.shape[1] > 0 else None)
        if not sap or sap not in selected_saps:
            continue

        raw_day = df.iloc[idx, 6] if df.shape[1] > 6 else None
        if not is_numeric_cell(raw_day):
            continue

        day = int(float(str(raw_day).strip().replace(",", ".")))
        if 1 <= day <= 6:
            days_by_sap.setdefault(sap, set()).add(day)

    return days_by_sap, sheet_name


def compare_tourenplanung_against_sap(
    uploaded_file,
    selected_saps: Set[str],
    days_in_sap_file: Dict[str, Set[int]],
) -> Tuple[pd.DataFrame, List[str]]:
    excel = pd.ExcelFile(uploaded_file)
    sheet_names = excel.sheet_names[:4]

    result_rows: List[dict] = []

    for sheet_name in sheet_names:
        df = pd.read_excel(excel, sheet_name=sheet_name, header=0)

        for idx in range(len(df)):
            sap = normalize_sap(df.iloc[idx, 1] if df.shape[1] > 1 else None)
            if not sap or sap not in selected_saps:
                continue

            vorhandene_tage = days_in_sap_file.get(sap, set())
            standort = CUSTOMER_TO_LOCATION.get(sap, "Ohne Zuordnung")

            for day, col_idx in DAY_COLUMNS_TOUR.items():
                if df.shape[1] <= col_idx:
                    continue

                raw_value = df.iloc[idx, col_idx]
                if not is_numeric_cell(raw_value):
                    continue

                if day in vorhandene_tage:
                    continue

                result_rows.append(
                    {
                        "Standort": standort,
                        "SAP Nummer": sap,
                        "Fehlender Liefertag Nummer": day,
                        "Fehlender Liefertag": DAY_NAMES[day],
                        "Blatt Tourenplanung": sheet_name,
                        "Wert in Tourenplanung": raw_value,
                        "Liefertage in SAP": ", ".join(
                            f"{d} {DAY_NAMES[d]}" for d in sorted(vorhandene_tage)
                        ),
                        "Hinweis": "Tag in Tourenplanung vorhanden, aber in SAP nicht als Liefertag hinterlegt",
                        "_StandortSort": LOCATION_ORDER.get(standort, 999),
                        "_KundenSort": CUSTOMER_TO_ORDER.get(sap, 999999),
                    }
                )

    export_columns = [
        "Standort",
        "SAP Nummer",
        "Fehlender Liefertag Nummer",
        "Fehlender Liefertag",
        "Blatt Tourenplanung",
        "Wert in Tourenplanung",
        "Liefertage in SAP",
        "Hinweis",
    ]

    if not result_rows:
        return pd.DataFrame(columns=export_columns), sheet_names

    result_df = pd.DataFrame(result_rows)
    result_df = result_df.drop_duplicates(
        subset=["SAP Nummer", "Fehlender Liefertag Nummer", "Blatt Tourenplanung"]
    )
    result_df = result_df.sort_values(
        ["_StandortSort", "_KundenSort", "SAP Nummer", "Fehlender Liefertag Nummer", "Blatt Tourenplanung"]
    ).reset_index(drop=True)
    result_df = result_df[export_columns]

    return result_df, sheet_names


def build_excel(result_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=False, sheet_name="Fehlt in SAP")
    return output.getvalue()


def build_group_overview() -> str:
    parts: List[str] = []
    for location_name, sap_list in CUSTOMER_GROUPS.items():
        parts.append(f"{location_name} ({len(sap_list)} Kunden)")
        parts.append("\n".join(sap_list))
        parts.append("")
    return "\n".join(parts).strip()


st.set_page_config(page_title="Tourenplanung gegen SAP", layout="wide")

st.title("Tourenplanung gegen SAP")
st.write(
    "Es wird nur eines geprüft: "
    "Steht in der Tourenplanung auf Montag bis Samstag ein Tag, "
    "dann muss dieser Tag auch in der SAP-Datei als Liefertag hinterlegt sein."
)

st.info(
    "Richtung des Vergleichs:\n"
    "- SAP = Datei mit SAP Nummer in A und Liefertag in G\n"
    "- Tourenplanung = Datei mit Spalte B sowie Montag bis Samstag in G bis L\n"
    "- Ausgabe = nur Tage, die in der Tourenplanung stehen, aber in SAP fehlen\n"
    "- Sortierung = zuerst Malchow, dann Neumünster, dann Zarrentin"
)

col1, col2, col3 = st.columns(3)
col1.metric("Malchow", len(CUSTOMER_GROUPS["Malchow"]))
col2.metric("Neumünster", len(CUSTOMER_GROUPS["Neumünster"]))
col3.metric("Zarrentin", len(CUSTOMER_GROUPS["Zarrentin"]))

with st.expander("Hinterlegte Kundensortierung", expanded=False):
    st.text_area(
        "Die Kunden werden mit dieser Reihenfolge ausgewertet und in der Excel genauso sortiert.",
        value=build_group_overview(),
        height=420,
        disabled=True,
    )

sap_datei = st.file_uploader(
    "SAP hochladen – erstes Blatt, Spalte A = SAP Nummer, Spalte G = Liefertag 1 bis 6",
    type=["xlsx", "xlsm", "xls"],
    key="sap_datei",
)

tourenplanung_datei = st.file_uploader(
    "Tourenplanung hochladen – erste 4 Blätter, Spalte B = SAP Nummer, Spalte G bis L = Montag bis Samstag",
    type=["xlsx", "xlsm", "xls"],
    key="tourenplanung_datei",
)

if st.button("Excel erzeugen", type="primary"):
    if not sap_datei or not tourenplanung_datei:
        st.error("Bitte beide Excel-Dateien hochladen.")
        st.stop()

    try:
        days_sap_file, sap_sheet = read_sap_file(sap_datei, SELECTED_SAPS)
        result_df, tourenplanung_sheets = compare_tourenplanung_against_sap(
            tourenplanung_datei,
            SELECTED_SAPS,
            days_sap_file,
        )

        excel_bytes = build_excel(result_df)

        st.success(
            f"Fertig. Gefunden wurden {len(result_df)} Zeilen, die in der Tourenplanung stehen, "
            f"aber in SAP als Liefertag fehlen."
        )

        st.caption(
            f"SAP: erstes Blatt = {sap_sheet} | "
            f"Tourenplanung: geprüfte Blätter = {', '.join(tourenplanung_sheets)}"
        )

        if not result_df.empty:
            standort_zaehlung = result_df.groupby("Standort").size().reset_index(name="Anzahl fehlender Tage")
            st.dataframe(standort_zaehlung, use_container_width=True, hide_index=True)

        st.download_button(
            label="Excel herunterladen",
            data=excel_bytes,
            file_name="tourenplanung_tage_fehlen_in_sap_sortiert.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:
        st.error(f"Fehler beim Verarbeiten der Dateien: {exc}")
