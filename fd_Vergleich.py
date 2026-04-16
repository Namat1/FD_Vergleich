import io
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

# pandas benutzt nullbasierte Spaltenindizes
DAY_COLUMNS_FILE2 = {
    1: 6,   # G
    2: 7,   # H
    3: 8,   # I
    4: 9,   # J
    5: 10,  # K
    6: 11,  # L
}


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


def parse_sap_list(text: str) -> List[str]:
    tokens = text.replace(",", " ").replace(";", " ").split()
    result: List[str] = []
    seen: Set[str] = set()
    for token in tokens:
        sap = normalize_sap(token)
        if sap and sap not in seen:
            seen.add(sap)
            result.append(sap)
    return result


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


def file1_days(uploaded_file, selected_saps: Set[str]) -> Tuple[Dict[str, Set[int]], pd.DataFrame, str]:
    excel = pd.ExcelFile(uploaded_file)
    sheet_name = excel.sheet_names[0]
    df = pd.read_excel(excel, sheet_name=sheet_name, header=0)

    days_by_sap: Dict[str, Set[int]] = {}
    rows: List[dict] = []

    for idx in range(len(df)):
        sap = normalize_sap(df.iloc[idx, 0] if df.shape[1] > 0 else None)  # Spalte A
        if selected_saps and sap not in selected_saps:
            continue
        if not sap:
            continue

        raw_day = df.iloc[idx, 6] if df.shape[1] > 6 else None  # Spalte G
        if not is_numeric_cell(raw_day):
            continue

        day = int(float(str(raw_day).strip().replace(",", ".")))
        if day < 1 or day > 6:
            continue

        days_by_sap.setdefault(sap, set()).add(day)
        rows.append(
            {
                "SAP Nummer": sap,
                "Liefertag Nummer": day,
                "Liefertag": DAY_NAMES[day],
                "Blatt": sheet_name,
                "Zeile": idx + 2,
            }
        )

    extract = pd.DataFrame(rows).drop_duplicates()
    if not extract.empty:
        extract = extract.sort_values(["SAP Nummer", "Liefertag Nummer"]).reset_index(drop=True)
    return days_by_sap, extract, sheet_name


def file2_new_days(uploaded_file, selected_saps: Set[str], days_by_sap_file1: Dict[str, Set[int]]) -> Tuple[pd.DataFrame, pd.DataFrame, List[str]]:
    excel = pd.ExcelFile(uploaded_file)
    sheet_names = excel.sheet_names[:4]

    detail_rows: List[dict] = []
    summary_map: Dict[str, Set[int]] = {}

    for sheet_name in sheet_names:
        df = pd.read_excel(excel, sheet_name=sheet_name, header=0)

        for idx in range(len(df)):
            sap = normalize_sap(df.iloc[idx, 1] if df.shape[1] > 1 else None)  # Spalte B
            if selected_saps and sap not in selected_saps:
                continue
            if not sap:
                continue

            existing_days = days_by_sap_file1.get(sap, set())

            for day, col_idx in DAY_COLUMNS_FILE2.items():
                if df.shape[1] <= col_idx:
                    continue

                raw_value = df.iloc[idx, col_idx]
                if not is_numeric_cell(raw_value):
                    continue

                if day in existing_days:
                    continue

                detail_rows.append(
                    {
                        "SAP Nummer": sap,
                        "Liefertag Nummer": day,
                        "Liefertag": DAY_NAMES[day],
                        "Neu in Datei 2": "Ja",
                        "Wert in Datei 2": raw_value,
                        "Blatt Datei 2": sheet_name,
                        "Zeile Datei 2": idx + 2,
                        "Tage in Datei 1": ", ".join(
                            f"{d} {DAY_NAMES[d]}" for d in sorted(existing_days)
                        ),
                    }
                )
                summary_map.setdefault(sap, set()).add(day)

    details = pd.DataFrame(detail_rows)
    if not details.empty:
        details = details.drop_duplicates(subset=["SAP Nummer", "Liefertag Nummer", "Blatt Datei 2", "Zeile Datei 2"])
        details = details.sort_values(["SAP Nummer", "Liefertag Nummer", "Blatt Datei 2", "Zeile Datei 2"]).reset_index(drop=True)

    summary_rows: List[dict] = []
    for sap, days in sorted(summary_map.items(), key=lambda item: item[0]):
        summary_rows.append(
            {
                "SAP Nummer": sap,
                "Neue Tage in Datei 2": ", ".join(f"{day} {DAY_NAMES[day]}" for day in sorted(days)),
                "Anzahl neue Tage": len(days),
            }
        )

    summary = pd.DataFrame(summary_rows)
    return details, summary, sheet_names


def build_excel(details: pd.DataFrame, summary: pd.DataFrame, extract_file1: pd.DataFrame, info: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, index=False, sheet_name="Uebersicht")
        details.to_excel(writer, index=False, sheet_name="Neu in Datei 2")
        extract_file1.to_excel(writer, index=False, sheet_name="Datei 1 Extrakt")
        info.to_excel(writer, index=False, sheet_name="Info")
    return output.getvalue()


def make_info_df(
    file1_name: str,
    file2_name: str,
    file1_sheet: str,
    file2_sheets: List[str],
    selected_count: int,
    summary: pd.DataFrame,
    details: pd.DataFrame,
) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"Feld": "Datei 1", "Wert": file1_name},
            {"Feld": "Datei 1 Blatt", "Wert": file1_sheet},
            {"Feld": "Datei 2", "Wert": file2_name},
            {"Feld": "Geprüfte Blätter Datei 2", "Wert": ", ".join(file2_sheets)},
            {"Feld": "Anzahl ausgewählte SAP Nummern", "Wert": selected_count},
            {"Feld": "SAP Nummern mit neuen Tagen in Datei 2", "Wert": len(summary)},
            {"Feld": "Anzahl Trefferzeilen", "Wert": len(details)},
            {"Feld": "Logik", "Wert": "Es werden nur Tage ausgegeben, die in Datei 2 vorhanden sind und in Datei 1 nicht vorhanden sind."},
        ]
    )


st.set_page_config(page_title="Liefertage nur neu in Datei 2", layout="wide")

st.title("Liefertage: nur neue Tage aus Datei 2")
st.write(
    "Diese App erzeugt nur eine Excel-Datei. "
    "Geprüft wird ausschließlich, ob in Datei 2 Liefertage vorhanden sind, die in Datei 1 nicht vorhanden sind."
)

with st.expander("SAP Nummern", expanded=False):
    sap_text = st.text_area(
        "Nur diese SAP Nummern vergleichen",
        value=DEFAULT_SAP_TEXT,
        height=320,
    )

file1 = st.file_uploader(
    "Datei 1 hochladen – erstes Blatt, Spalte A = SAP Nummer, Spalte G = Liefertag 1 bis 6",
    type=["xlsx", "xlsm", "xls"],
    key="file1",
)

file2 = st.file_uploader(
    "Datei 2 hochladen – erste vier Blätter, Spalte B = SAP Nummer, Spalte G bis L = Montag bis Samstag",
    type=["xlsx", "xlsm", "xls"],
    key="file2",
)

if st.button("Excel erzeugen", type="primary"):
    selected_list = parse_sap_list(sap_text)
    selected_set = set(selected_list)

    if not file1 or not file2:
        st.error("Bitte beide Excel-Dateien hochladen.")
        st.stop()

    if not selected_list:
        st.error("Bitte mindestens eine SAP Nummer eingeben.")
        st.stop()

    try:
        days_file1, extract_file1, file1_sheet = file1_days(file1, selected_set)
        details, summary, file2_sheets = file2_new_days(file2, selected_set, days_file1)

        info = make_info_df(
            file1_name=file1.name,
            file2_name=file2.name,
            file1_sheet=file1_sheet,
            file2_sheets=file2_sheets,
            selected_count=len(selected_list),
            summary=summary,
            details=details,
        )

        excel_bytes = build_excel(details, summary, extract_file1, info)

        st.success(
            f"Fertig. Gefunden wurden {len(details)} Zeilen mit Tagen, die nur in Datei 2 vorhanden sind."
        )
        st.download_button(
            label="Excel herunterladen",
            data=excel_bytes,
            file_name="liefertage_nur_neu_in_datei2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:
        st.error(f"Fehler beim Verarbeiten der Dateien: {exc}")
