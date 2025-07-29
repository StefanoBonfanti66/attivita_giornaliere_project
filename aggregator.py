import pandas as pd
import glob

def aggregate_data():
    excel_files = glob.glob("OpzioniEsportazione*.xlsx")
    all_dfs = []

    for file in excel_files:
        try:
            # Legge l'intero foglio per trovare l'operatore e la tabella dei dati
            df_raw = pd.read_excel(file, header=None)

            # 1. Estrae il nome dell'operatore. Ora legge da A2 (riga 2, colonna 1).
            operator_name = "Sconosciuto"
            if df_raw.shape[0] > 1:
                # Legge il valore dalla prima colonna (A2)
                raw_name = df_raw.iloc[1, 0]
                if pd.notna(raw_name):
                    # Estrae il nome prima del trattino
                    operator_name = str(raw_name).split('-')[0].strip()

            # 2. Trova la riga di intestazione. È la prima riga con più di 2 valori non vuoti.
            header_row_index = -1
            for i, row in df_raw.iterrows():
                if row.count() > 2:
                    header_row_index = i
                    break
            
            if header_row_index == -1:
                print(f"Non è stata trovata nessuna riga di intestazione nel file {file}")
                continue

            # 3. Crea il DataFrame con le intestazioni corrette
            df_data = df_raw.iloc[header_row_index:].copy()
            df_data.columns = df_data.iloc[0]
            df_data = df_data[1:]
            df_data.reset_index(drop=True, inplace=True)

            # 4. Aggiunge la colonna 'Operatore'
            df_data['Operatore'] = operator_name
            
            all_dfs.append(df_data)

        except Exception as e:
            print(f"Errore durante l'elaborazione del file {file}: {e}")

    if all_dfs:
        final_df = pd.concat(all_dfs, ignore_index=True)
        # Pulisce i nomi delle colonne
        final_df.columns = final_df.columns.str.strip()
        # Rimuove le righe completamente vuote
        final_df.dropna(how='all', inplace=True)
        final_df.to_csv("aggregated_data.csv", index=False)
    else:
        print("Nessun dato è stato aggregato.")
        # Crea un file vuoto per evitare errori
        pd.DataFrame().to_csv("aggregated_data.csv", index=False)


if __name__ == "__main__":
    aggregate_data()