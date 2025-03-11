import streamlit as st

st.set_page_config(
    page_title="SAS - AUC", 
    page_icon="🏃",
    layout="wide")

import xml.etree.ElementTree as ET
import pandas as pd
from io import BytesIO
import os

def parse_row(row, ns):
    """
    Reconstructs a full row from an XML <Row> by checking for the ss:Index attribute.
    Missing cells are filled in with None so that the output list represents the complete row.
    """
    row_values = []
    current_index = 1  # ss:Index is 1-based.
    for cell in row.findall('ss:Cell', ns):
        index_attr = cell.get('{urn:schemas-microsoft-com:office:spreadsheet}Index')
        if index_attr is not None:
            index_val = int(index_attr)
            while current_index < index_val:
                row_values.append(None)
                current_index += 1
        data_elem = cell.find('ss:Data', ns)
        text_val = data_elem.text if data_elem is not None else None
        row_values.append(text_val)
        current_index += 1
    return row_values

def parse_filename(filename):
    """
    Parses the file name to extract:
      - Surname: first token.
      - Name: all tokens between the first token and "Treadmill" (joined with spaces).
      - Session: the token at index (Treadmill_index + 9).
      - Musica: the token at index (Treadmill_index + 10); if it starts with "NM", output "no musica",
                if it starts with "M", output "musica".

    Expected file name format:
      Surname_Name(...additional)_Treadmill_8km_h_dd_mm_yyyy_hh_mm_ss_T1(orT2)_M(orNM).xml
    """
    base = os.path.basename(filename)
    if base.lower().endswith('.xml'):
        base = base[:-4]
    tokens = base.split('_')
    try:
        treadmill_idx = tokens.index("Treadmill")
    except ValueError:
        raise ValueError(f"File name '{filename}' does not contain 'Treadmill'.")
    surname = tokens[0]
    name = " ".join(tokens[1:treadmill_idx])
    if len(tokens) < treadmill_idx + 11:
        raise ValueError(f"File name '{filename}' not long enough to extract session/musica.")
    session_token = tokens[treadmill_idx + 9]
    musica_token_raw = tokens[treadmill_idx + 10]
    if '.' in musica_token_raw:
        musica_token = musica_token_raw.split('.')[0]
    else:
        musica_token = musica_token_raw
    if musica_token.upper().startswith("NM"):
        musica = "no musica"
    elif musica_token.upper().startswith("M"):
        musica = "musica"
    else:
        musica = musica_token
    return surname, name, session_token, musica

def process_single_file(file_bytes, original_filename):
    """
    Processes a single XML file from in-memory bytes:
      - Parses the XML and reconstructs rows.
      - Extracts subject details from the original file name.
      - Finds the STOP row & Tempo[s].
      - Averages columns (to the right of Tempo[s]) in last 15 minutes.
      - Returns dict with subject details + numeric averages.
    """
    ns = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}

    # Parse the file from in-memory bytes
    tree = ET.parse(file_bytes)
    root = tree.getroot()

    worksheet = root.find('.//ss:Worksheet[@ss:Name="Dati OJ"]', ns)
    if worksheet is None:
        raise ValueError("Worksheet 'Dati OJ' not found.")
    table = worksheet.find('.//ss:Table', ns)
    if table is None:
        raise ValueError("No <Table> in the Worksheet.")

    rows = table.findall('ss:Row', ns)
    parsed_rows = [parse_row(r, ns) for r in rows]
    if len(parsed_rows) < 2:
        raise ValueError("Not enough rows.")

    header = parsed_rows[0]
    try:
        hash_col = header.index('#')
    except ValueError:
        hash_col = 23  # fallback
    try:
        tempo_col = header.index('Tempo[s]')
    except ValueError:
        tempo_col = 25  # fallback

    # Extract subject details from the original file name
    subject_surname, subject_name, session_token, musica = parse_filename(original_filename)

    # Find STOP row
    stop_row_index = None
    T_stop = None
    for i, row in enumerate(parsed_rows):
        if len(row) > hash_col and row[hash_col]:
            if "Impulso esterno STOP" in row[hash_col]:
                stop_row_index = i
                if len(row) > tempo_col and row[tempo_col]:
                    try:
                        T_stop = float(row[tempo_col].replace(',', '.'))
                    except ValueError:
                        raise ValueError(f"Invalid Tempo[s] in STOP row: {row[tempo_col]}")
                break
    if stop_row_index is None or T_stop is None:
        raise ValueError("No valid STOP row/Tempo[s].")

    T_start = T_stop - 900.0  # 15 minutes
    num_cols = len(header)

    sums = {col: 0.0 for col in range(tempo_col+1, num_cols)}
    counts = {col: 0 for col in range(tempo_col+1, num_cols)}

    for row in parsed_rows[1:stop_row_index]:
        if len(row) > tempo_col and row[tempo_col]:
            try:
                t_val = float(row[tempo_col].replace(',', '.'))
            except ValueError:
                continue
            if T_start <= t_val <= T_stop:
                for col in range(tempo_col+1, num_cols):
                    if col < len(row) and row[col] and row[col].strip():
                        try:
                            val = float(row[col].replace(',', '.'))
                        except ValueError:
                            continue
                        sums[col] += val
                        counts[col] += 1

    measurement_headers = header[tempo_col+1:]
    measurement_avgs = []
    for col in range(tempo_col+1, num_cols):
        if counts[col] > 0:
            measurement_avgs.append(sums[col] / counts[col])  # store as float
        else:
            measurement_avgs.append(None)

    return {
        "Cognome": subject_surname,
        "Nome": subject_name,
        "Sessione": session_token,
        "Musica": musica,
        "Measurements": measurement_avgs,
        "MeasurementHeaders": measurement_headers
    }

def main():
    st.title("Optojump Data Analyser")
    st.write("Seleziona i file di optojump, attenta che siano descritti come Cognome_Nome(o_più_nomi)_Treadmill_8km_h_28_02_2025_13_18_12_Sessione_Musica/Nomusica")

    # Let the user upload multiple XML files
    uploaded_files = st.file_uploader(
        "Select one or more XML files",
        type=["xml"],
        accept_multiple_files=True
    )

    # Let the user specify the output Excel file name
    default_name = "combined_averages.xlsx"
    output_filename = st.text_input("Output Excel File Name", value=default_name)

    if uploaded_files and st.button("Process Files"):
        results = []
        measurement_headers = None

        for up_file in uploaded_files:
            if up_file is not None:
                try:
                    # 'up_file' is a UploadedFile, we can pass its buffer directly to ET.parse
                    res = process_single_file(up_file, up_file.name)
                    results.append(res)
                    if measurement_headers is None:
                        measurement_headers = res["MeasurementHeaders"]
                except Exception as e:
                    st.error(f"Error processing {up_file.name}: {e}")

        if not results:
            st.warning("No valid results to process.")
            return

        # Build the final DataFrame
        col_names = ["Cognome", "Nome", "Sessione", "Musica"]
        if measurement_headers:
            col_names += measurement_headers

        data_rows = []
        for r in results:
            row = [r["Cognome"], r["Nome"], r["Sessione"], r["Musica"]]
            row += r["Measurements"]
            data_rows.append(row)

        df = pd.DataFrame(data_rows, columns=col_names)

        # Write to an in-memory BytesIO buffer
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Averages")

        st.success(f"Processed {len(df)} files successfully!")

        # Provide a download button
        st.download_button(
            label="Download Excel File",
            data=buffer.getvalue(),
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
