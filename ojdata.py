import streamlit as st

st.set_page_config(
    page_title="Optojump Analyser", 
    page_icon=":material/sprint:",
    layout="wide")

import streamlit as st
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
      - Cognome: first token.
      - Nome: all tokens between the first token and "Treadmill" (joined with spaces).
      - Sessione: the token at index (Treadmill_index + 9).
      - Musica: the token at index (Treadmill_index + 10); if it starts with "NM", output "no musica",
                if it starts with "M", output "musica".
    Expected file name format:
      Cognome_Nome(...additional)_Treadmill_8km_h_dd_mm_yyyy_hh_mm_ss_T1(orT2)_M(orNM).xml
    """
    base = os.path.basename(filename)
    if base.lower().endswith('.xml'):
        base = base[:-4]
    tokens = base.split('_')
    try:
        treadmill_idx = tokens.index("Treadmill")
    except ValueError:
        raise ValueError(f"Il file '{filename}' non contiene 'Treadmill'.")
    surname = tokens[0]
    name = " ".join(tokens[1:treadmill_idx])
    if len(tokens) < treadmill_idx + 11:
        raise ValueError(f"Il file '{filename}' non ha abbastanza token per estrarre sessione e musica.")
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
    Processa un singolo file XML:
      - Parsea l'XML e ricostruisce le righe.
      - Estrae le informazioni del soggetto dal nome del file.
      - Trova la riga in cui compare "Impulso esterno STOP" e legge il valore di Tempo[s].
      - Definisce un intervallo di 15 minuti (900 secondi) precedenti T_stop.
      - Per tutte le colonne a destra di Tempo[s], calcola la media sui dati nell'intervallo, 
        eccetto per la prima colonna (Distanza[m]), per la quale viene preso l'ultimo valore valido.
      - Restituisce un dizionario con le informazioni del soggetto, T_start, T_stop e i valori calcolati.
    """
    ns = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}
    tree = ET.parse(file_bytes)
    root = tree.getroot()

    worksheet = root.find('.//ss:Worksheet[@ss:Name="Dati OJ"]', ns)
    if worksheet is None:
        raise ValueError("Worksheet 'Dati OJ' non trovato.")
    table = worksheet.find('.//ss:Table', ns)
    if table is None:
        raise ValueError("Nessun <Table> trovato nel file.")
    rows = table.findall('ss:Row', ns)
    parsed_rows = [parse_row(r, ns) for r in rows]
    if len(parsed_rows) < 2:
        raise ValueError("Numero insufficiente di righe nel file.")

    header = parsed_rows[0]
    try:
        hash_col = header.index('#')
    except ValueError:
        hash_col = 23
    try:
        tempo_col = header.index('Tempo[s]')
    except ValueError:
        tempo_col = 25

    # Estrae informazioni soggetto dal nome del file
    subject_surname, subject_name, session_token, musica = parse_filename(original_filename)

    # Trova la riga STOP e il valore di T_stop
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
                        raise ValueError(f"Valore non valido in Tempo[s] nella riga STOP: {row[tempo_col]}")
                break
    if stop_row_index is None or T_stop is None:
        raise ValueError("Nessuna riga STOP valida o Tempo[s] non numerico.")

    T_start = T_stop - 900.0  # intervallo di 15 minuti

    num_cols = len(header)
    sums = {col: 0.0 for col in range(tempo_col+1, num_cols)}
    counts = {col: 0 for col in range(tempo_col+1, num_cols)}
    last_values = {col: None for col in range(tempo_col+1, num_cols)}  # per memorizzare l'ultimo valore valido

    # Processa le righe dall'indice 1 fino a quella precedente la STOP
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
                        # Per la prima colonna dopo Tempo[s] (Distanza[m]), salviamo l'ultimo valore
                        if col == tempo_col + 1:
                            last_values[col] = val
                        else:
                            sums[col] += val
                            counts[col] += 1

    averages = {}
    for col in range(tempo_col+1, num_cols):
        if col == tempo_col + 1:
            # Utilizza l'ultimo valore valido invece della media
            averages[col] = last_values[col]
        else:
            if counts[col] > 0:
                averages[col] = sums[col] / counts[col]
            else:
                averages[col] = None

    measurement_headers = header[tempo_col+1:]
    measurement_avgs = []
    for col in range(tempo_col+1, num_cols):
        if averages[col] is not None:
            measurement_avgs.append(averages[col])
        else:
            measurement_avgs.append(None)

    return {
        "Cognome": subject_surname,
        "Nome": subject_name,
        "Sessione": session_token,
        "Musica": musica,
        "T_start": T_start,
        "T_stop": T_stop,
        "Measurements": measurement_avgs,
        "MeasurementHeaders": measurement_headers
    }

def process_multiple_files(file_list):
    results = []
    measurement_headers = None
    for up_file in file_list:
        if up_file is not None:
            try:
                res = process_single_file(up_file, up_file.name)
                results.append(res)
                if measurement_headers is None:
                    measurement_headers = res["MeasurementHeaders"]
            except Exception as e:
                st.error(f"Errore nel file {up_file.name}: {e}")
    return results, measurement_headers

def main():
    st.title("Estrazione dati Optojump (ultimi 15 minuti)")
    
    st.markdown("""
    **Descrizione del Programma**

    Questo programma consente di caricare uno o più file XML (in formato Excel 2003 XML) contenenti dati di test su tapis roulant.  
    **Funzionalità:**
    - **Estrazione delle Informazioni del Soggetto:**  
      Le informazioni (Cognome, Nome, Sessione e condizione "Musica" o "no musica") vengono estratte dal nome del file.  
      Il nome del file deve seguire il formato:  
      `Cognome_Nome(_secondonome_...)_Treadmill_8km_h_dd_mm_yyyy_hh_mm_ss_T1(orT2)_M(orNM).xml`  
      Assicurati che il file sia nominato correttamente; in caso contrario, il programma genererà un errore. Se c'è scritto M2, M1, NM1, NM2 non è un problema, ma perché farlo? 🤨
    
    - **Elaborazione dei Dati:**  
      Il programma analizza il file XML per individuare la riga in cui compare "Impulso esterno STOP" e legge il valore di `Tempo[s]`.  
      Viene definito un intervallo di 15 minuti (900 secondi) precedenti il valore di STOP e vengono calcolate le medie dei valori per ogni colonna a destra di `Tempo[s]`.
    
    - **Output:**  
      I risultati vengono salvati in un file Excel con le seguenti colonne:  
      **Cognome, Nome, Sessione, Musica**, seguiti dalle colonne con i valori medi calcolati.  
      Inoltre, viene mostrato un riepilogo sullo schermo con il nome del soggetto, la sessione, la condizione e i tempi di inizio (T_start) e fine (T_stop) dell'intervallo analizzato.
    
    **Come Usare il Programma:**
    1. Carica uno o più file XML utilizzando il pulsante "Carica uno o più file XML".
    2. Inserisci il nome desiderato per il file Excel di output.
    3. Clicca sul pulsante "Process Files".
    4. Scarica il file Excel generato tramite il pulsante "Download File Excel".
    """)

    uploaded_files = st.file_uploader("Carica uno o più file XML", type=["xml"], accept_multiple_files=True)
    output_filename = st.text_input("Nome del file Excel di output", value="combined_averages.xlsx")
    
    if uploaded_files and st.button("Process Files"):
        results, measurement_headers = process_multiple_files(uploaded_files)
        if not results:
            st.warning("Nessun file processato correttamente.")
            return
        
        # Build and display summary table
        summary_data = []
        for res in results:
            summary_data.append({
                "Cognome": res["Cognome"],
                "Nome": res["Nome"],
                "Sessione": res["Sessione"],
                "Musica": res["Musica"],
                "T_start": res["T_start"],
                "T_stop": res["T_stop"]
            })
        summary_df = pd.DataFrame(summary_data)
        st.subheader("Riepilogo File Processati")
        st.dataframe(summary_df)
        
        # Build final DataFrame for Excel output
        col_names = ["Cognome", "Nome", "Sessione", "Musica"] + ["T_start", "T_stop"]
        if measurement_headers:
            col_names += measurement_headers
        data_rows = []
        for r in results:
            row = [r["Cognome"], r["Nome"], r["Sessione"], r["Musica"], r["T_start"], r["T_stop"]] + r["Measurements"]
            data_rows.append(row)
        df = pd.DataFrame(data_rows, columns=col_names)
        
        # Write to Excel file in-memory
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Averages")
        
        st.success("Elaborazione completata con successo!")
        st.download_button(
            label="Download File Excel",
            data=buffer.getvalue(),
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
