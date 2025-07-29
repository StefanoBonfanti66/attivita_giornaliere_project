
import schedule
import time
import json
import importlib
import os

def load_automations():
    """Carica dinamicamente i moduli di automazione dalla cartella 'automations'."""
    automations = {}
    for filename in os.listdir('automations'):
        if filename.endswith('.py'):
            module_name = f"automations.{filename[:-3]}"
            module = importlib.import_module(module_name)
            automations[filename[:-3]] = module
    return automations

def main():
    """Funzione principale che avvia lo scheduler."""
    print("Avvio dello scheduler di automazione...")
    
    automations = load_automations()
    
    with open('config.json', 'r') as f:
        config = json.load(f)
        
    for job_config in config.get('jobs', []):
        automation_name = job_config.get('automation')
        schedule_time = job_config.get('time')
        
        if automation_name in automations:
            automation_module = automations[automation_name]
            # Assumendo che ogni modulo di automazione abbia una funzione 'run'
            if hasattr(automation_module, 'run'):
                schedule.every().day.at(schedule_time).do(automation_module.run)
                print(f"Pianificata l'automazione '{automation_name}' alle {schedule_time}")
            else:
                print(f"Attenzione: L'automazione '{automation_name}' non ha una funzione 'run'.")
        else:
            print(f"Attenzione: Automazione '{automation_name}' non trovata.")

    print("Scheduler avviato. In attesa dei task...")
    while True:
        schedule.run_pending()
        time.sleep(1)

if __name__ == "__main__":
    main()
