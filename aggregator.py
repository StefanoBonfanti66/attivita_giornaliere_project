
import pandas as pd
import glob
import os

def aggregate_data():
    # Otteniamo il percorso assoluto della directory in cui si trova lo script.
    # Questo garantisce che i percorsi funzionino correttamente sia in locale che sulla VPS.
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Creiamo il pattern di ricerca per i file Excel all'interno di questa directory.
    search_pattern = os.path.join(base_dir, "OpzioniEsportazione*.xlsx")
    excel_files = glob.glob(search_pattern)

    all_dfs = []

    for file in excel_files:
        try:
            df_raw = pd.read_excel(file, header=None)

            operator_name = "Sconosciuto"
            if df_raw.shape[0] > 1:
                raw_name = df_raw.iloc[1, 0]
                if pd.notna(raw_name):
                    operator_name = str(raw_name).split('-')[0].strip()

            header_row_index = -1
            for i, row in df_raw.iterrows():
                if row.count() > 2:
                    header_row_index = i
                    break
            
            if header_row_index == -1:
                print(f"Nessuna riga di intestazione trovata in {file}")
                continue

            df_data = df_raw.iloc[header_row_index:].copy()
            df_data.columns = df_data.iloc[0]
            df_data = df_data[1:]
            df_data.reset_index(drop=True, inplace=True)
            df_data['Operatore'] = operator_name
            
            all_dfs.append(df_data)

        except Exception as e:
            print(f"Errore durante l'elaborazione del file {file}: {e}")

    # Definiamo il percorso completo per il file CSV di output.
    output_csv_path = os.path.join(base_dir, "aggregated_data.csv")

    if all_dfs:
        final_df = pd.concat(all_dfs, ignore_index=True)
        final_df.columns = final_df.columns.str.strip()
        final_df.dropna(how='all', inplace=True)
        final_df.to_csv(output_csv_path, index=False)
        print(f"Dati aggregati e salvati in {output_csv_path}")
    else:
        print("Nessun file Excel trovato o nessun dato da aggregare.")
        pd.DataFrame().to_csv(output_csv_path, index=False)


if __name__ == "__main__":
    aggregate_data()
