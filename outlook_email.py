"""Utility per preparare una bozza di email in Outlook e allegare un file.

Funziona solo su Windows con Outlook installato.

Esempio di utilizzo:
    python outlook_email.py --file aggregated_data.csv --subject "Report giornaliero" --body "In allegato il report." --to stefano@example.com

Se `pywin32` non è installato lo script stampa un messaggio esplicativo.
"""
from __future__ import annotations
import os
import sys
import argparse
import time
from typing import List
import tempfile
import mimetypes
from email.message import EmailMessage


_RETRY_DELAY = 2.0
_MAX_RETRIES = 3


def create_outlook_draft(file_path: str, subject: str = "", body: str = "", to: List[str] | None = None, display: bool = True):
    """Crea una bozza di email in Outlook e allega `file_path`.

    Args:
        file_path: percorso del file da allegare (se esiste).
        subject: oggetto dell'email.
        body: corpo del messaggio.
        to: lista di destinatari (indirizzi email).
        display: se True mostra la finestra di composizione, altrimenti invia direttamente.

    Ritorna l'oggetto MailItem di Outlook se disponibile, altrimenti None.
    """
    try:
        import pythoncom
        import win32com.client
        from win32com.client import gencache
    except Exception as e:  # pragma: no cover - run only on Windows with pywin32
        print("Impossibile importare win32com.client/pythoncom. Questa funzionalità richiede pywin32 e Outlook su Windows.")
        print("Installa con: pip install pywin32")
        raise

    # Ensure COM is initialized for the current thread
    try:
        pythoncom.CoInitialize()
    except Exception:
        # If already initialized, ignore
        pass

    outlook = None
    last_exc = None
    # Prefer to check whether Outlook is already running to avoid DispatchEx blocking
    def _is_outlook_running() -> bool:
        try:
            import subprocess
            res = subprocess.run(["tasklist", "/FI", "IMAGENAME eq OUTLOOK.EXE"], capture_output=True, text=True)
            out = res.stdout or ""
            return "OUTLOOK.EXE" in out.upper()
        except Exception:
            return False

    # If Outlook is not running, try to start it and wait a bit for the COM server to be ready
    if not _is_outlook_running():
        try:
            import subprocess
            print("Outlook non rilevato in esecuzione: provo ad avviarlo...")
            subprocess.Popen(["outlook.exe"], shell=False)
            # wait for process to appear
            waited = 0.0
            while waited < (_MAX_RETRIES * _RETRY_DELAY):
                if _is_outlook_running():
                    break
                time.sleep(1.0)
                waited += 1.0
            if not _is_outlook_running():
                print("Avvio Outlook non riuscito o processo non disponibile dopo l'attesa.")
        except Exception as launch_exc:
            print(f"Impossibile avviare Outlook automaticamente: {launch_exc}")

    # Try to obtain the COM object using EnsureDispatch first, then Dispatch as a fallback
    try:
        outlook = gencache.EnsureDispatch("Outlook.Application")
    except Exception as e:
        last_exc = e
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as e2:
            last_exc = e2

    if outlook is None:
        # Provide detailed diagnostic information and attempt a CLI fallback
        err_msg = f"Impossibile ottenere l'istanza COM di Outlook. Ultima eccezione: {repr(last_exc)}"
        print(err_msg)

        # Diagnostic: print python bitness
        try:
            import struct
            print(f"Python process bitness: {struct.calcsize('P') * 8}-bit")
        except Exception:
            pass

        # Fallback: try to launch Outlook with /a to attach the file (works on many Outlook installs)
        if file_path and os.path.exists(os.path.abspath(file_path)):
            try:
                import subprocess
                abs_path = os.path.abspath(file_path)
                print("Tentativo fallback: avvio Outlook con comando '/a' per allegare il file...")
                # Note: outlook.exe should be in PATH if Office is installed; this relies on the system to resolve it.
                subprocess.Popen(["outlook.exe", "/a", abs_path], shell=False)
                print("Comando di avvio Outlook con allegato eseguito. Controlla se si apre una nuova bozza.")
                return None
            except Exception as launch_exc:
                print(f"Fallback tramite linea di comando fallito: {launch_exc}")

        # Ulteriore fallback: creare un file .eml (RFC822) con l'allegato e aprirlo con il client predefinito.
        try:
            abs_path = os.path.abspath(file_path) if file_path else None
            if abs_path and os.path.exists(abs_path):
                print("Tentativo fallback alternativo: creo un file .eml e lo apro con il client predefinito...")
                msg = EmailMessage()
                msg["Subject"] = subject or ""
                if to:
                    msg["To"] = ";".join(to)
                msg["From"] = ""
                msg.set_content(body or "")

                ctype, encoding = mimetypes.guess_type(abs_path)
                if ctype is None:
                    ctype = "application/octet-stream"
                maintype, subtype = ctype.split("/", 1)

                with open(abs_path, "rb") as f:
                    file_data = f.read()
                msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=os.path.basename(abs_path))

                with tempfile.NamedTemporaryFile(delete=False, suffix=".eml") as tmp:
                    tmp_path = tmp.name
                    tmp.write(msg.as_bytes())
                print(f"File .eml creato in: {tmp_path}; provo ad aprirlo con il client predefinito...")
                try:
                    os.startfile(tmp_path)
                    print("Aperto file .eml con il client predefinito. Controlla la bozza.")
                    return None
                except Exception as open_exc:
                    print(f"Impossibile aprire il file .eml: {open_exc}")
        except Exception as eml_exc:
            print(f"Errore durante creazione/apertura .eml fallback: {eml_exc}")

        raise RuntimeError(err_msg)

    mail = outlook.CreateItem(0)  # 0: olMailItem
    mail.Subject = subject or ""
    mail.Body = body or ""
    if to:
        # Outlook accetta destinatari separati da ';'
        mail.To = ";".join(to)

    if file_path:
        abs_path = os.path.abspath(file_path)
        if os.path.exists(abs_path):
            mail.Attachments.Add(Source=abs_path)
        else:
            print(f"Attenzione: file non trovato: {abs_path}")

    try:
        if display:
            mail.Display(True)
        else:
            mail.Send()
    except Exception as e:
        print(f"Errore durante Display/Send della mail: {e}")
        raise

    try:
        # Uninitialize COM for the thread if we initialized it here
        pythoncom.CoUninitialize()
    except Exception:
        pass

    return mail


def _parse_args(argv: List[str] | None = None):
    parser = argparse.ArgumentParser(description="Prepara una bozza di email in Outlook e allega un file.")
    parser.add_argument("--file", "-f", dest="file", required=True, help="Percorso del file da allegare")
    parser.add_argument("--subject", "-s", dest="subject", default="", help="Oggetto dell'email")
    parser.add_argument("--body", "-b", dest="body", default="", help="Corpo dell'email")
    parser.add_argument("--to", dest="to", nargs="*", help="Destinatari (separati da spazio)")
    parser.add_argument("--send", dest="send", action="store_true", help="Se passato invia l'email invece di aprire la bozza")
    return parser.parse_args(argv)


def main(argv: List[str] | None = None):
    args = _parse_args(argv)
    try:
        create_outlook_draft(args.file, subject=args.subject, body=args.body, to=args.to, display=not args.send)
        print("Operazione completata: bozza creata (o inviata se --send).")
    except Exception as e:
        print("Errore durante la creazione della bozza Outlook:", e)
        sys.exit(1)


if __name__ == "__main__":
    main()
