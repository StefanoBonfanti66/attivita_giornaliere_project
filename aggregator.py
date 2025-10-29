import pandas as pd
import glob
import os
import datetime
import subprocess
import time
import argparse
import traceback
import sys

def run_git_command(command, cwd, check_exit_code=True, input=None):
    """Helper function to run git commands."""
    try:
        result = subprocess.run(command, cwd=cwd, check=check_exit_code, capture_output=True, text=True, input=input)
        print(f"Git command output: {result.stdout.strip()}")
        return result.stdout, result.returncode
    except subprocess.CalledProcessError as e:
        print(f"Error running Git command: {e}")
        print(f"Stdout: {e.stdout.strip()}")
        print(f"Stderr: {e.stderr.strip()}")
        return None, e.returncode

def aggregate_data():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    search_pattern = os.path.join(base_dir, "OpzioniEsportazione*.xlsx")
    excel_files = glob.glob(search_pattern)
    # Tieni traccia del file Excel più recente per l'email
    latest_excel = None
    if excel_files:
        latest_excel = max(excel_files, key=os.path.getmtime)

    all_dfs = []

    for file in excel_files:
        try:
            # Read all sheets from the Excel file
            xls = pd.ExcelFile(file)
            
            for sheet_name in xls.sheet_names:
                # Skip sheets that are not the processed data (e.g., original "Foglio1" or "Sheet1")
                # Assuming processed sheets have an underscore, like "Operatore_Categoria"
                if "_" not in sheet_name:
                    continue

                df_sheet = pd.read_excel(xls, sheet_name=sheet_name, skiprows=2) # Skip logo and title rows
                
                # Extract Operatore and Categoria from sheet_name
                parts = sheet_name.split('_')
                if len(parts) >= 2:
                    operator_name = parts[0]
                    category = parts[1]
                    df_sheet['Operatore'] = operator_name
                    df_sheet['Categoria'] = category # Add Categoria column
                else:
                    df_sheet['Operatore'] = "Sconosciuto"
                    df_sheet['Categoria'] = "Sconosciuto"

                all_dfs.append(df_sheet)

        except Exception as e:
            print(f"Errore durante l'elaborazione del file {file}: {e}")

    output_csv_path = os.path.join(base_dir, "aggregated_data.csv")
    output_csv_filename = "aggregated_data.csv"

    if all_dfs:
        final_df = pd.concat(all_dfs, ignore_index=True)
        final_df.columns = final_df.columns.str.strip()
        final_df.dropna(how='all', inplace=True)
        final_df.to_csv(output_csv_path, index=False)
        print(f"Dati aggregati e salvati in {output_csv_path}")
        print(f"DEBUG: Size of saved CSV: {os.path.getsize(output_csv_path)} bytes")

        time.sleep(1)

        try:
            if not os.path.exists(output_csv_path) or os.path.getsize(output_csv_path) == 0:
                print(f"Warning: {output_csv_filename} does not exist or is empty. Skipping Git operations.")
                return None

            # Force remove from cache if it was previously tracked, then add again
            print(f"Attempting to remove {output_csv_filename} from Git cache (if present).")
            run_git_command(["git", "rm", "--cached", output_csv_filename], base_dir, check_exit_code=False)

            # Add the file to staging area
            print(f"DEBUG: Attempting to add {output_csv_filename} to Git staging area.")
            run_git_command(["git", "add", output_csv_filename], base_dir)

            # Check what Git has staged for this file
            stdout_ls_files_staged, returncode_ls_files_staged = run_git_command(["git", "ls-files", "--stage", output_csv_filename], base_dir)
            print(f"DEBUG: git ls-files --stage output: '{stdout_ls_files_staged.strip()}'")

            # Check if aggregated_data.csv is modified (M) or added (A) in git status --porcelain
            stdout_status, returncode_status = run_git_command(["git", "status", "--porcelain", output_csv_filename], base_dir)

            # Strip any leading/trailing whitespace, including newlines
            stdout_status = stdout_status.strip()

            # Always attempt to commit and push if we reached this point
            print(f"Attempting to commit and push changes for {output_csv_filename}...")

            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            commit_message = f"Automated data aggregation update - {timestamp}"
            run_git_command(["git", "commit", "--allow-empty", "-F", "-"], base_dir, input=commit_message)

            run_git_command(["git", "push", "origin", "main"], base_dir)
            print("Git push completed.")

            # Return both paths so callers can choose which allegare
            return output_csv_path, latest_excel

        except Exception as e:
            print(f"Errore durante l'automazione Git: {e}")

    else:
        print("Nessun file Excel trovato o nessun dato da aggregare.")
        pd.DataFrame().to_csv(output_csv_path, index=False)
        return None, None


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Aggrega i file Excel in aggregated_data.csv e opzionalmente prepara una bozza Outlook.")
    parser.add_argument("--email", dest="email", action="store_true", help="Se passato, apre una bozza Outlook con il file aggregato come allegato (Windows + Outlook + pywin32).")
    parser.add_argument("--email-to", dest="email_to", nargs="*", help="Lista di destinatari per la bozza Outlook")
    parser.add_argument("--email-subject", dest="email_subject", default="Report giornaliero", help="Oggetto per la bozza Outlook")
    parser.add_argument("--email-body", dest="email_body", default="In allegato il report.", help="Corpo del messaggio per la bozza Outlook")
    args = parser.parse_args()

    csv_path, latest_excel = aggregate_data()

    if args.email:
        # Preferisci il file Excel originale se esiste
        email_attachment = latest_excel if latest_excel else csv_path
        if email_attachment and os.path.exists(email_attachment) and os.path.getsize(email_attachment) > 0:
            try:
                # Import lazily because win32com is Windows-only
                from outlook_email import create_outlook_draft

                print("[debug] Chiamata a create_outlook_draft()...")
                sys.stdout.flush()
                result = None
                try:
                    result = create_outlook_draft(email_attachment, subject=args.email_subject, body=args.email_body, to=args.email_to, display=True)
                finally:
                    print(f"[debug] create_outlook_draft() ha restituito: {result!r}")
                    sys.stdout.flush()
                print(f"Bozza Outlook creata con successo allegando: {os.path.basename(email_attachment)}")
            except Exception as e:
                print("Impossibile creare la bozza Outlook:")
                traceback.print_exc()
                sys.stdout.flush()
        else:
            print("Nessun file valido da allegare (né Excel né CSV); non è stata creata la bozza Outlook.")