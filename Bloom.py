import blpapi 
from blpapi import SessionOptions, Session


from datetime import datetime,timedelta
import statistics


def top_ticker_by_volume(isins):

    allowed_exch_codes = {'FP','GR','IM','US','UP','LN','GY','GR','NA','HK', 'IE','JT','JP','JJ','DC','NO','ID','LX','CH','SE'}  

    session_options = SessionOptions()
    session_options.serverHost = "localHost"
    session_options.serverPort = 8194
    session = Session(session_options)
    if not session.start():
        raise RuntimeError('Cannot start bloom')
    if not session.openService("//blp/refdata"):
        raise RuntimeError("Cannot open refdata")
    refdata = session.getService("//blp/refdata")

    results = {}

    for isin in isins :
        if not session.openService("//blp/instruments"):
            raise RuntimeError("Cannot open instrument")
        instruments_service = session.getService("//blp/instruments")
        instr_req = instruments_service.createRequest("instrumentListRequest")
        instr_req.set('query',isin)
        instr_req.set('maxResults',100)
        session.sendRequest(instr_req)

        tickers = []

        while True :
            ev = session.nextEvent()
            for msg in ev:
                if msg.messageType() !=blpapi.Name("InstrumentListResponse"):
                    continue 
                results_list = msg.getElement("results")
                for i in range(results_list.numValues()):
                    r = results_list.getValue(i)
                    ticker = r.getElementAsString("security")
                    
                    
                    
                    ticker = ticker.split('<')[0]
                    

                    if ticker.split(' ')[1] in allowed_exch_codes:
                        tickers.append({"ticker": f"{ticker} Equity"})
                            
                    
                           
                  
            if ev.eventType() == blpapi.Event.RESPONSE:
                break
        
        if not tickers:
            results[isin] = None
            continue

        
        end_date = datetime.today()
        start_date = end_date - timedelta(2)
        liste_volume = []
        
        if tickers :

            for t in tickers:
                avg_volumes = {}
                ref_request = refdata.createRequest("HistoricalDataRequest")
                ref_request.getElement("fields").appendValue("PX_VOLUME")
                
                ref_request.getElement("securities").appendValue(t['ticker'])
                ref_request.set("startDate", start_date.strftime("%Y%m%d"))
                ref_request.set("endDate", end_date.strftime("%Y%m%d"))
                ref_request.set("periodicitySelection","DAILY")
                
                session.sendRequest(ref_request)

                while True:
                    ev2 = session.nextEvent()
                    for msg2 in ev2:
                        if msg2.messageType() != blpapi.Name("HistoricalDataResponse"):
                            continue
                        sec_data = msg2.getElement("securityData")
                        if sec_data.numValues() == 0 :
                            avg_volumes[t['ticker']] = 0 
                            continue
                        fd = sec_data.getElement("fieldData")
                        volumes = []
                        for j in range(fd.numValues()):
                            day = fd.getValue(j)
                            if day.hasElement('PX_VOLUME'):
                                volumes.append(day.getElementAsFloat("PX_VOLUME"))
                        if volumes == []:
                            continue
                        else:
                            avg_volumes['volume'] = statistics.mean(volumes)
                            avg_volumes['ticker'] = t['ticker']
                            liste_volume.append(avg_volumes)
                    
                    if ev2.eventType() == blpapi.Event.RESPONSE:
                        break
            
            results[isin] = liste_volume

        else:
            results[isin] = None
            
    

    
    return results




def ticker_bloom(isins):
    """
    - Entree : une liste d'isins 
    - Sortie : un dictionnaire avec pour clé les isins et en valeurs une liste de dictionnaire avec les ticker correspondant
    """
    
    
    allowed_exch_codes = {'FP','GR','IM','US','UP','LN','GY','GR','NA','HK', 'IE','JT','JP','JJ','DC','NO','ID','LX','AT','GF','CN','CH','SE'}
      
    
    
    session_options = SessionOptions()
    session_options.serverHost = "localHost"
    session_options.serverPort = 8194
    session = Session(session_options)

    if not session.start():
        raise RuntimeError('Cannot start bloom')
    if not session.openService("//blp/refdata"):
        raise RuntimeError("Cannot open refdata")
    session.getService("//blp/refdata")

    results = {}

    for isin in isins :
        if not session.openService("//blp/instruments"):
            raise RuntimeError("Cannot open instrument")
        instruments_service = session.getService("//blp/instruments")
        instr_req = instruments_service.createRequest("instrumentListRequest")
        instr_req.set('query',isin)
        instr_req.set('maxResults',100)
        session.sendRequest(instr_req)

        tickers = []

        while True :
            ev = session.nextEvent()
            for msg in ev:
                if msg.messageType() !=blpapi.Name("InstrumentListResponse"):
                    continue 
                results_list = msg.getElement("results")
                for i in range(results_list.numValues()): 
                    r = results_list.getValue(i)
                    ticker = r.getElementAsString("security")
                    
                    
                   
                    ticker = ticker.split('<')[0] 
                    if ticker != "NONE":
                        if ticker.split(' ')[1] in allowed_exch_codes:
                            tickers.append( f"{ticker} Equity")
        
            if ev.eventType() == blpapi.Event.RESPONSE:
                    break
        
        if not tickers:
            results[isin] = None
            continue

        results[isin] = tickers
        
        
    
    return results

def get_ticker_info_from_dic(info_dic):
    options = SessionOptions()
    options.serverHost="localhost"
    options.serverPort = 8194

    session = Session(options)
    session.start()
    session.openService("//blp/refdata")
    refdata = session.getService("//blp/refdata")

    request = refdata.createRequest("ReferenceDataRequest")

    final = {}

    for isin in info_dic:
        request = refdata.createRequest("ReferenceDataRequest")
        tickers_list = info_dic[isin]
        
        if tickers_list :
            for ticker in tickers_list:
                
                request.getElement("securities").appendValue(ticker)

            for f in ["CRNCY","ID_MIC_PRIM_EXCH","PX_VOLUME","VOLUME_AVG_30D","ID_ISIN", "NAME","SECURITY_TYP"]:
                
                request.getElement("fields").appendValue(f)

            session.sendRequest(request)
            results = {}
            

            while True :
                ev = session.nextEvent()
                for msg in ev:
                    if msg.messageType() != blpapi.Name("ReferenceDataResponse"):
                        continue
                    else:

                        data = msg.getElement("securityData")
                        results = {}
                        for i in range(data.numValues()):
                            sec = data.getValue(i)
                            ticker = sec.getElementAsString("security")

                    
                                
                            if sec.hasElement("securityError"):
                                
                                results[ticker]=[]
                                continue


                
                            fieldData = sec.getElement("fieldData")

                            
                            currency = fieldData.getElementAsString("CRNCY") if fieldData.hasElement("CRNCY") else None 
                            exchange = fieldData.getElementAsString("ID_MIC_PRIM_EXCH") if fieldData.hasElement("ID_MIC_PRIM_EXCH") else None 
                            volume = fieldData.getElementAsString("PX_VOLUME") if fieldData.hasElement("PX_VOLUME") else None
                            volume_moyenne_30 = fieldData.getElementAsString("VOLUME_AVG_30D") if fieldData.hasElement("VOLUME_AVG_30D") else None
                            isin = fieldData.getElementAsString("ID_ISIN") if fieldData.hasElement("ID_ISIN") else None
                            name = fieldData.getElementAsString("NAME") if fieldData.hasElement("NAME") else None
                            type = fieldData.getElementAsString("SECURITY_TYP") if fieldData.hasElement("SECURITY_TYP") else None
                             
                            if volume_moyenne_30:
                                n = int(float(volume_moyenne_30))
                                volume_moyenne_30 = f"{n:,}"
                                results[ticker]= {'currency' : currency, 'exchange':exchange, "volume" : volume, "volume_moyen" : volume_moyenne_30, "isin" : isin, 'name' : name, 'type' : type}


                if ev.eventType() == blpapi.Event.RESPONSE:
                    break
            final[isin] = results
        else:
            final[isin] = None
    return final



def get_isin_info(isins):

    info_dic = ticker_bloom(isins)
   
   
    result = get_ticker_info_from_dic(info_dic)
   

    return result



def get_ticker_info_from_tickers(tickers,idx_liste):
    
    
    options = SessionOptions()
    options.serverHost="localhost"
    options.serverPort = 8194

    session = Session(options)
    session.start()
    session.openService("//blp/refdata")
    refdata = session.getService("//blp/refdata")

    request = refdata.createRequest("ReferenceDataRequest")

    results = {}
    isin_etf = {}

    request = refdata.createRequest("ReferenceDataRequest")

    if tickers:

        for ticker in tickers :
            
            request.getElement("securities").appendValue(ticker)

        for f in ["CRNCY","ID_MIC_PRIM_EXCH","PX_VOLUME","VOLUME_AVG_30D","ID_ISIN", 'NAME',"SECURITY_TYP"]:
            request.getElement("fields").appendValue(f)
        
        session.sendRequest(request)
        
        securité = 0 
        while True :
            securité +=1
            ev = session.nextEvent()
            
            for msg in ev:
                
                
                
                if msg.messageType() != blpapi.Name("ReferenceDataResponse"):
                    
                    continue
                else:

                    data = msg.getElement("securityData")
                    
                    for i in range(data.numValues()):
                        
                        
                        sec = data.getValue(i)
                        id_int = int(sec.getElementAsString('sequenceNumber'))
                        
                        idx = idx_liste[id_int]
                        
                        ticker = sec.getElementAsString("security")
                        

                        if sec.hasElement("securityError"):
                            
                            results[ticker]={'idx':idx,'isin' : 'FR0000000000','currency': 'Not Valid', 'exchange': 'Not Valid'}
                            continue

                        fieldData = sec.getElement("fieldData")

                        
                        currency = fieldData.getElementAsString("CRNCY") if fieldData.hasElement("CRNCY") else None 
                        exchange = fieldData.getElementAsString("ID_MIC_PRIM_EXCH") if fieldData.hasElement("ID_MIC_PRIM_EXCH") else None 
                        volume = fieldData.getElementAsString("PX_VOLUME") if fieldData.hasElement("PX_VOLUME") else None
                        volume_moyen = fieldData.getElementAsString("VOLUME_AVG_30D") if fieldData.hasElement("VOLUME_AVG_30D") else None
                        isin = fieldData.getElementAsString("ID_ISIN") if fieldData.hasElement("ID_ISIN") else None
                        name = fieldData.getElementAsString("NAME") if fieldData.hasElement("NAME") else None
                        type = fieldData.getElementAsString("SECURITY_TYP") if fieldData.hasElement("SECURITY_TYP") else None

                        if volume_moyen :
                            n = int(float(volume_moyen))
                            volume_moyen = f"{n:,}"
                            results[ticker]= {'currency' : currency, 'exchange':exchange, "volume" : volume, "volume_moyen" : volume_moyen, 'isin' : isin, 'idx' : idx , 'name' : name, 'type' : type} 
                        else: 
                            
                            results[ticker] = {'idx': idx, 'exchange' : None, 'currency' : None}
                  
                    if securité >50:
                        break                
                    if ev.eventType() == blpapi.Event.RESPONSE:
                        
                        return results
    else : 
        results = None      
    return results

isins = ['US00724F1012','US0079031078','US0010551028','US00846U1016','US02079K3059','US0231351067','US0258161092','US0268747849','US03076C1062','US0378331005','US0382221051','US0404131064','US0527691069','US0530151036','US0640581007','US0758871091','US11133T1034','US1273871087','US14040H1059','US1491231015','CH0044328745','US1720621010','US17275R1023','US1941621039','US22160K1051','US2441991054','US2600031080','US2786421030','US2855121099','US0367521038','US5324571083','US2944291051','US3032501047','US3377381088','US34959E1091','CH0114405324','US36266G1076','US3848021040','US42824C1099','US40434L1052','US4448591028','US4523081093','US4581401001','US4592001014','US4612021034','US46120E6023','US46625H1005']


        
tickers = ['ADP US Equity', 'AMTM US Equity', 'AMZN US Equity', 'ARM US Equity', 'ASML NA Equity', 'ASSAB SS Equity', 'AVGO US Equity', 'AXON US Equity', 'AZO US Equity', 'BABA US Equity', 'BKNG US Equity', 'BN FP Equity', 'BNYL LN Equity', 'BVI FP Equity', 'CDNS US Equity', 'CHD US Equity', 'COLOB DC Equity', 'COO US Equity', 'COST US Equity', 'CPAY US Equity', 'CPG LN Equity', 'CSU CN Equity', 'DOL CN Equity', 'DPLM LN Equity', 'DPZ US Equity', 'DSY FP Equity', 'EL FP Equity', 'EXPN LN Equity', 'FISV US Equity', 'GOOG US Equity', 'GPN US Equity', 'HLMA LN Equity', 'HMB SS Equity', 'ISRG US Equity', 'ITRK LN Equity', 'ITX SQ Equity', 'JKHY US Equity', 'JMT PL Equity', 'LIFCB SS Equity', 'MA US Equity', 'META US Equity', 'MKC US Equity', 'MSFT US Equity', 'NSISB DC Equity', 'NVDA US Equity', 'NVO US Equity', 'NXT LN Equity', 'ONON US Equity', 'OR FP Equity', 'ORLY US Equity', 'PSMT US Equity', 'PYPL US Equity', 'RMD US Equity', 'ROL US Equity', 'ROST US Equity', 'SAP GY Equity', 'SBAC US Equity', 'SHOP US Equity', 'SNPS US Equity', 'SPSC US Equity', 'SYK US Equity', 'TEAM US Equity', 'TJX US Equity', 'TSCO US Equity', 'UNFI US Equity', 'UNP US Equity', 'UTDI GY Equity', 'V US Equity', 'VIS SQ Equity', 'VRT US Equity', 'XYZ US Equity']
result = get_ticker_info_from_tickers(tickers,[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70])
#print(result)
def bloom_format_transformation_ticker(ticker):

    if len(ticker.split(" ")) == 2:
        ticker = ticker + " Equity"
        
    return ticker



