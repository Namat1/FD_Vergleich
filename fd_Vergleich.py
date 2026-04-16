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

DAY_COLUMNS_PLANUNG = {
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


def read_liefertag_datei(uploaded_file, selected_saps: Set[str]) -> Tuple[Dict[str, Set[int]], str]:
    excel = pd.ExcelFile(uploaded_file)
    sheet_name = excel.sheet_names[0]
    df = pd.read_excel(excel, sheet_name=sheet_name, header=0)

    days_by_sap: Dict[str, Set[int]] = {}

    for idx in range(len(df)):
        sap = normalize_sap(df.iloc[idx, 0] if df.shape[1] > 0 else None)  # Spalte A
        if not sap or (selected_saps and sap not in selected_saps):
            continue

        raw_day = df.iloc[idx, 6] if df.shape[1] > 6 else None  # Spalte G
        if not is_numeric_cell(raw_day):
            continue

        day = int(float(str(raw_day).strip().replace(",", ".")))
        if 1 <= day <= 6:
            days_by_sap.setdefault(sap, set()).add(day)

    return days_by_sap, sheet_name


def compare_planung_against_liefertage(
    uploaded_file,
    selected_saps: Set[str],
    days_in_liefertag_datei: Dict[str, Set[int]],
) -> Tuple[pd.DataFrame, List[str]]:
    excel = pd.ExcelFile(uploaded_file)
    sheet_names = excel.sheet_names[:4]

    result_rows: List[dict] = []

    for sheet_name in sheet_names:
        df = pd.read_excel(excel, sheet_name=sheet_name, header=0)

        for idx in range(len(df)):
            sap = normalize_sap(df.iloc[idx, 1] if df.shape[1] > 1 else None)  # Spalte B
            if not sap or (selected_saps and sap not in selected_saps):
                continue

            vorhandene_tage = days_in_liefertag_datei.get(sap, set())

            for day, col_idx in DAY_COLUMNS_PLANUNG.items():
                if df.shape[1] <= col_idx:
                    continue

                raw_value = df.iloc[idx, col_idx]
                if not is_numeric_cell(raw_value):
                    continue

                if day in vorhandene_tage:
                    continue

                result_rows.append(
                    {
                        "SAP Nummer": sap,
                        "Fehlender Liefertag Nummer": day,
                        "Fehlender Liefertag": DAY_NAMES[day],
                        "Blatt Planung": sheet_name,
                        "Zeile Planung": idx + 2,
                        "Wert in Planung": raw_value,
                        "Liefertage in anderer Datei": ", ".join(
                            f"{d} {DAY_NAMES[d]}" for d in sorted(vorhandene_tage)
                        ),
                        "Hinweis": "Tag in Planung vorhanden, aber in anderer Datei nicht als Liefertag hinterlegt",
                    }
                )

    result_df = pd.DataFrame(result_rows)
    if not result_df.empty:
        result_df = result_df.drop_duplicates(
            subset=["SAP Nummer", "Fehlender Liefertag Nummer", "Blatt Planung", "Zeile Planung"]
        )
        result_df = result_df.sort_values(
            ["SAP Nummer", "Fehlender Liefertag Nummer", "Blatt Planung", "Zeile Planung"]
        ).reset_index(drop=True)

    return result_df, sheet_names


def build_excel(result_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=False, sheet_name="Fehlt in anderer Datei")
    return output.getvalue()


st.set_page_config(page_title="Planung gegen Liefertage", layout="wide")

st.title("Planung gegen Liefertag-Datei")
st.write(
    "Es wird nur eines geprüft: "
    "Steht in der Planung auf Montag bis Samstag ein Tag, "
    "dann muss dieser Tag auch in der anderen Datei als Liefertag hinterlegt sein."
)

st.info(
    "Richtung des Vergleichs:\n"
    "- Planung = Datei mit Spalte B sowie Montag bis Samstag in G bis L\n"
    "- Andere Datei = Datei mit SAP Nummer in A und Liefertag in G\n"
    "Ausgegeben wird nur, was in der Planung steht, aber in der anderen Datei fehlt."
)

with st.expander("SAP Nummern", expanded=False):
    sap_text = st.text_area(
        "Nur diese SAP Nummern vergleichen",
        value=DEFAULT_SAP_TEXT,
        height=320,
    )

liefertag_datei = st.file_uploader(
    "Andere Datei hochladen – erstes Blatt, Spalte A = SAP Nummer, Spalte G = Liefertag 1 bis 6",
    type=["xlsx", "xlsm", "xls"],
    key="liefertag_datei",
)

planung_datei = st.file_uploader(
    "Planung hochladen – erste 4 Blätter, Spalte B = SAP Nummer, Spalte G bis L = Montag bis Samstag",
    type=["xlsx", "xlsm", "xls"],
    key="planung_datei",
)

if st.button("Excel erzeugen", type="primary"):
    selected_list = parse_sap_list(sap_text)
    selected_set = set(selected_list)

    if not liefertag_datei or not planung_datei:
        st.error("Bitte beide Excel-Dateien hochladen.")
        st.stop()

    if not selected_list:
        st.error("Bitte mindestens eine SAP Nummer eingeben.")
        st.stop()

    try:
        days_other_file, other_sheet = read_liefertag_datei(liefertag_datei, selected_set)
        result_df, planung_sheets = compare_planung_against_liefertage(
            planung_datei,
            selected_set,
            days_other_file,
        )

        excel_bytes = build_excel(result_df)

        st.success(
            f"Fertig. Gefunden wurden {len(result_df)} Zeilen, die in der Planung stehen, "
            f"aber in der anderen Datei als Liefertag fehlen."
        )

        st.caption(
            f"Andere Datei: erstes Blatt = {other_sheet} | "
            f"Planung: geprüfte Blätter = {', '.join(planung_sheets)}"
        )

        st.download_button(
            label="Excel herunterladen",
            data=excel_bytes,
            file_name="planung_tage_fehlend_in_anderer_datei.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:
        st.error(f"Fehler beim Verarbeiten der Dateien: {exc}")
