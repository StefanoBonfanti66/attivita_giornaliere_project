
import pandas as pd
import glob
import os
import datetime
import subprocess # Import subprocess
import time # Import time module

def run_git_command(command, cwd, check_exit_code=True):
    """Helper function to run git commands."""
    try:
        result = subprocess.run(command, cwd=cwd, check=check_exit_code, capture_output=True, text=True)
        print(f"Git command output: {result.stdout.strip()}")
        return result.stdout, result.returncode # Return stdout and returncode
    except subprocess.CalledProcessError as e:
        print(f"Error running Git command: {e}")
        print(f"Stdout: {e.stdout.strip()}")
        print(f"Stderr: {e.stderr.strip()}")
        return None, e.returncode # Return None for stdout and the error returncode

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
    output_csv_filename = "aggregated_data.csv" # For git add

    if all_dfs:
        final_df = pd.concat(all_dfs, ignore_index=True)
        final_df.columns = final_df.columns.str.strip()
        final_df.dropna(how='all', inplace=True)
        final_df.to_csv(output_csv_path, index=False)
        print(f"Dati aggregati e salvati in {output_csv_path}")

        # Give the OS a moment to write the file
        time.sleep(1)

        # --- Git Automation ---
        try:
            # Ensure the file exists and is not empty before attempting git add
            if not os.path.exists(output_csv_path) or os.path.getsize(output_csv_path) == 0:
                print(f"Warning: {output_csv_filename} does not exist or is empty. Skipping Git operations.")
                return

            # Force remove from cache if it was previously tracked, then add again
            print(f"Attempting to remove {output_csv_filename} from Git cache (if present)...")
            run_git_command(["git", "rm", "--cached", output_csv_filename], base_dir, check_exit_code=False)

            # Add the file to staging area
            print(f"Adding {output_csv_filename} to Git staging area...")
            run_git_command(["git", "add", output_csv_filename], base_dir)

            # --- DIAGNOSTIC STEP: Check git status immediately after add ---
            print("Running git status after add...")
            stdout_status, returncode_status = run_git_command(["git", "status"], base_dir, check_exit_code=False)
            print(f"git status after add stdout:\n{stdout_status}")
            print(f"git status after add returncode: {returncode_status}")
            # --- END DIAGNOSTIC STEP ---

            # --- NEW DIAGNOSTIC STEP: Check git check-ignore ---
            print(f"Running git check-ignore -v {output_csv_filename}...")
            stdout_check_ignore, returncode_check_ignore = run_git_command(["git", "check-ignore", "-v", output_csv_filename], base_dir, check_exit_code=False)
            print(f"git check-ignore -v stdout:\n{stdout_check_ignore}")
            print(f"git check-ignore -v returncode: {returncode_check_ignore}")
            # --- END NEW DIAGNOSTIC STEP ---

            # --- NEW DIAGNOSTIC STEP: Check git ls-files --stage ---
            print(f"Running git ls-files --stage {output_csv_filename}...")
            stdout_ls_files, returncode_ls_files = run_git_command(["git", "ls-files", "--stage", output_csv_filename], base_dir, check_exit_code=False)
            print(f"git ls-files --stage stdout:\n{stdout_ls_files}")
            print(f"git ls-files --stage returncode: {returncode_ls_files}")
            # --- END NEW DIAGNOSTIC STEP ---

            # Check if there are any staged changes to commit
            # git diff --staged --quiet returns 0 if no changes, 1 if changes
            print("Running git diff --staged --quiet...")
            stdout, returncode = run_git_command(["git", "diff", "--staged", "--quiet"], base_dir, check_exit_code=False)
            print(f"git diff --staged --quiet stdout: '{stdout}', returncode: {returncode}")

            if returncode == 0: # No changes staged
                print(f"No changes detected in {output_csv_filename} after staging. Skipping Git commit/push.")
            else: # Changes are staged (returncode is 1 for changes, or other for error)
                if returncode == 1:
                    print(f"Changes detected in {output_csv_filename}. Committing and pushing...")

                    # Commit the changes
                    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    commit_message = f"Automated data aggregation update - {timestamp}"
                    run_git_command(["git", "commit", "-m", commit_message], base_dir)

                    # Push to remote (assuming 'origin' and 'main' branch)
                    run_git_command(["git", "push", "origin", "main"], base_dir)
                    print("Git push completed.")
                else:
                    print(f"Unexpected return code from git diff --staged --quiet: {returncode}. Cannot proceed with commit/push.")

        except Exception as e:
            print(f"Errore durante l'automazione Git: {e}")

    else:
        print("Nessun file Excel trovato o nessun dato da aggregare.")
        pd.DataFrame().to_csv(output_csv_path, index=False)


if __name__ == "__main__":
    aggregate_data()
