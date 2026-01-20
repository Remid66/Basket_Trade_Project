import json
import uuid
from pathlib import Path
import json 
import time

# =================================================================================================
# =                                   SECTION: ENVOIE REQUETE 
# =================================================================================================

CONFIG = 'config.json'


def wait_for_file(path:Path,timeout = 120):
    start = time.time()
    while not path.exists():
        if time.time() - start > timeout:
            raise TimeoutError(f"{path.name} non reçu")
        time.sleep(0.5)


def write_isin_info_to_bloom(isin_list_dic,request_path):
    json_path = Path(request_path)
    json_path.parent.mkdir(parents=True,exist_ok=True)
    
    
    with open(json_path,"w",encoding="utf-8") as f:
        json.dump(isin_list_dic,f,indent = 4, ensure_ascii=False)
    


def wait_for_file_ready(path:Path,timeout = 10):
    
    start = time.time()
    last_size = -1
    while time.time() - start < timeout:
        size = path.stat().st_size
        if size == last_size:
            return
        last_size = size
        time.sleep(0.2)

def read_json(path : Path) ->dict:
    with open(path,'r',encoding = 'utf-8') as f:
        return json.load(f)
    
    


def discussion (fonction,input,user_id):
    """
    fonction : fonction à utiliser avec Bloom
    input : données d'entrées 

    Fonction permettant l'envoie et la récupération des données dans Bloom.
    """
    with open(CONFIG, 'r', encoding='utf-8') as f: 
        dic = json.load(f)
    work_path = dic['absolut_path']
    echange_information_bloom__name_global = dic['echange_information_bloom__name_global']
    
    REQUEST_DIR = Path(f"{work_path}\{echange_information_bloom__name_global}_{user_id}\\requests")
    REPONSE_DIR = Path(f"{work_path}\{echange_information_bloom__name_global}_{user_id}\\responses")

    id = uuid.uuid4().hex
    
    request_id = f"{fonction}{id}"
    request_path = REQUEST_DIR/f"{user_id}_{request_id}.json"

    reponse_path = REPONSE_DIR/f"{user_id}_{request_id}.json"
    input_dic = {'user_id': user_id, 'request_id' : request_id, 'input' : input, 'id' : id,'fonction' : fonction}

    
    write_isin_info_to_bloom(input_dic,request_path)


    wait_for_file(reponse_path)
    wait_for_file_ready(reponse_path)

    reponse = read_json(reponse_path)
    

    #request_path.unlink(missing_ok=True)
    #reponse_path.unlink(missing_ok=True)
    
    return reponse["result"]
