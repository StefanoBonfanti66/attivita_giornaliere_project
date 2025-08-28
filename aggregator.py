import pandas as pd
import glob
import os
import datetime
import subprocess
import time

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

        time.sleep(1)

        try:
            if not os.path.exists(output_csv_path) or os.path.getsize(output_csv_path) == 0:
                print(f"Warning: {output_csv_filename} does not exist or is empty. Skipping Git operations.")
                return

            # Force remove from cache if it was previously tracked, then add again
            print(f"Attempting to remove {output_csv_filename} from Git cache (if present).")
            run_git_command(["git", "rm", "--cached", output_csv_filename], base_dir, check_exit_code=False)

            # Add the file to staging area
            print(f"Adding {output_csv_filename} to Git staging area.")
            run_git_command(["git", "add", output_csv_filename], base_dir)

            # Check if aggregated_data.csv is modified (M) or added (A) in git status --porcelain
            stdout_status, returncode_status = run_git_command(["git", "status", "--porcelain", output_csv_filename], base_dir)

            # Strip any leading/trailing whitespace, including newlines
            stdout_status = stdout_status.strip()

            print(f"DEBUG: stdout_status (stripped) = '{stdout_status}'")
            print(f"DEBUG: Checking for 'M {output_csv_filename}' = 'M {output_csv_filename}'")
            print(f"DEBUG: Checking for 'A {output_csv_filename}' = 'A {output_csv_filename}'")

            # Check if the output contains 'M ' (modified) or 'A ' (added) for the file
            # The space after M/A is important to match the status output format
            if f"M {output_csv_filename}" in stdout_status or f"A {output_csv_filename}" in stdout_status:
                print(f"Changes detected in {output_csv_filename}. Committing and pushing...")

                timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                commit_message = f"Automated data aggregation update - {timestamp}"
                run_git_command(["git", "commit", "-F", "-"], base_dir, input=commit_message)

                run_git_command(["git", "push", "origin", "main"], base_dir)
                print("Git push completed.")
            else:
                print(f"No changes detected in {output_csv_filename} after staging. Skipping Git commit/push.")

        except Exception as e:
            print(f"Errore durante l'automazione Git: {e}")

    else:
        print("Nessun file Excel trovato o nessun dato da aggregare.")
        pd.DataFrame().to_csv(output_csv_path, index=False)


if __name__ == "__main__":
    aggregate_data()