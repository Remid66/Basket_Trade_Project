import time 
import subprocess
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import json
from Bloom import get_isin_info,get_ticker_info_from_tickers

BLOOM_USER = 'user_a'
CONFIG = 'config.json'


 

processing_file = set()


def wait_for_file_ready(path:Path,timeout = 10):
    
    start = time.time()
    last_size = -1
    while time.time() - start < timeout:
        size = path.stat().st_size
        if size == last_size:
            return
        last_size = size
        time.sleep(0.2)

def safe_json_load(path,retries = 5, delay = 0.3):
    for _ in range (retries):
        try:
            with open(path,'r',encoding="utf-8") as f:
                return json.load(f)
        
        except json.JSONDecodeError:
            time.sleep(delay)
    raise

def process_request(data: dict) ->dict:
    input = data.get('input')
    user_id = data.get('user_id')
    id = data.get('id')
    fonction = data.get('fonction')
    if fonction == 'isin_info':
        result = get_isin_info(input)
    
    if fonction == 'ticker_info':
        idx_list = input['idx_list']
        ticker = input['tickers']
        result = get_ticker_info_from_tickers(ticker,idx_list)
    
    
    return {
        "status" : "OK",
        "user_id": user_id,
        "fonction": fonction,
        "id" : id,
        "result": result
    }


class NewFileHandler(FileSystemEventHandler):

    def on_created(self,event):
        path = Path(event.src_path)
        if path.suffix != '.json' or path in processing_file:
            return
        processing_file.add(path)
        time.sleep(0.5)

        try:
            wait_for_file_ready(path)
            request = safe_json_load(path)      
            response = process_request(request)
        
            response_path = REPONSE_DIR / path.name
            tmp = response_path.with_suffix(".tmp")
            
            
            with open(response_path,'w',encoding="utf-8") as f:
                json.dump(response,f,indent=4)
        except Exception as e :
            print(f"Erreur sur {path.name}: {e}")
        finally:
            processing_file.remove(path)

    



if __name__ == '__main__':
    with open(CONFIG,'r',encoding='utf-8') as f:
        dic = json.load(f) 
    echange_information_bloom__name_global = dic['echange_information_bloom__name_global']
    work_path = dic['absolut_path']

    REQUEST_DIR = Path(f"{work_path}\{echange_information_bloom__name_global}_{BLOOM_USER}\\requests")
    REPONSE_DIR = Path(f"{work_path}\{echange_information_bloom__name_global}_{BLOOM_USER}\\responses")


    observer = Observer()
    observer.schedule(NewFileHandler(), REQUEST_DIR, recursive=False)
    observer.start()

    print(f"Surveillance du dossier ")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    
    observer.join()
