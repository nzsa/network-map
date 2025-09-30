import pandas as pd
import yfinance as yf
from threading import Thread
from queue import Queue
from datetime import datetime
import numpy as np

class ArgsForListFunctions():
    def __init__(self, specificArg, globalArgs):
        self.specificArg = specificArg
        self.globalArgs = globalArgs

def getAllStocks(exchange):
    """Gets a screen object for all tickers on a passed exchange, this can be used to get the Ticker objects.
    NZE=NZX, ASX=ASX"""
    maxMarketCap = 0
    tickerList = []

    while maxMarketCap==0 or tickerList[-1].empty == False:
        eqQuery = yf.EquityQuery('and', [yf.EquityQuery('is-in', ['exchange', exchange]), yf.EquityQuery('gt', ['intradaymarketcap', maxMarketCap])])
        rawYf = yf.screen(eqQuery, size=250, sortField='intradaymarketcap', sortAsc=True)['quotes']
        tickerList.append(pd.DataFrame(rawYf))
        try:
            maxMarketCap = max(tickerList[-1]['marketCap'])
            print(maxMarketCap)
            if len(tickerList[-1]) < 250:
                break
        except:
            break
    df = pd.concat(tickerList)
    return df

def createTicker(tickerArg):
    """Creates a ticker object from the ticker string stored in the specificArg of a ArgsForListFunction class"""
    return yf.Ticker(tickerArg.specificArg)

def threadWorker(q):
    """Completes jobs in the passed queue"""
    while True:
        dictionaryOut, key, func, args, argsGlobal = q.get()
        if key is None:
            q.task_done()
            break
        args = ArgsForListFunctions(args, argsGlobal)
        dictionaryOut[key] = func(args)
        q.task_done()

def fixDates(df, dateColumns):
    """Convert epoch 0+seconds to datetime"""
    for dateColumn in dateColumns:
        df[dateColumn] = df[dateColumn].fillna(0)
        try:
            df[dateColumn] = df[dateColumn].astype(int)
            df[dateColumn] = df[dateColumn].apply(datetime.fromtimestamp)
            df[dateColumn] = df[dateColumn].replace(datetime(1970, 1, 1, 0, 0, 0), None)
        except:
            df[dateColumn] = df[dateColumn].apply(lambda x: x.to_pydatetime())
            df[dateColumn] = df[dateColumn].replace(datetime(1970, 1, 1, 0, 0, 0), None)
    
    return df

def filterColumnsContaining(df, strings):
    desiredColumns = [[]]*len(strings)
    i=0
    for string in strings:
        desiredColumns[i] = [col for col in df.columns if (string in col)]
        i+=1
    desiredColumns = [item for sublist in desiredColumns for item in sublist]
    df = df[desiredColumns]
    return df

def unpackExecutives(df, executiveColumns):
    for column in executiveColumns:  
        for ticker in df.index:
            print(df.loc[ticker][column])
            try:
                for person in df.loc[ticker][column]:
                    role = person['title'] + ' 1'
                    name = person['name']
                    if role not in df.columns:
                        # Create the column with NaNs
                        df[role] = np.nan

                    inserted = 0
                    directorNum = 1
                    while inserted == 0:
                        # Check if the specific cell is NaN before assigning
                        if pd.isna(df.loc[ticker, role]):
                            df.loc[ticker, role] = name
                            inserted = 1
                        else:
                            directorNum += 1
                            role = role[:-1] + str(directorNum)
            except:
                continue
    df = df.drop(executiveColumns, axis=1) 

    return df

def unpackEmbeddedDicts(embeddedDict, tickObj):
    try:
        i = 1
        for person in tickObj.info[embeddedDict]:
            role = 'Person ' + str(i)
            name = person['name']
            tickObj.info[role] = name
            i += 1
        del tickObj.info[embeddedDict]
    except:
        pass

    return tickObj


def createTickInfo(allArgs):
    """Takes a dict -containing: tickObj and converts its into to a dataframe.
    embeddedDicts a list of embeddedDicts in a ticker object to expand them. ready for df conversion. 
    Expands with title 'Person #'.
    dropColumns with a list of columns to drop, or keepColumns with a list of columns to keep. these are mutually exclusive.
    """

    tickObj = allArgs.specificArg

    try:
        embeddedDicts = allArgs.globalArgs['embeddedDicts']
    except:
        embeddedDicts = []
    try:
        dropColumns = allArgs.globalArgs['dropColumns']
    except:
        dropColumns = None
    try:
        keepColumns = allArgs.globalArgs['keepColumns']
    except:
        keepColumns = None
    
    print(tickObj.ticker)

    for embeddedDict in embeddedDicts:
        tickObj = unpackEmbeddedDicts(tickObj, embeddedDict)

    tickInfo = pd.DataFrame([tickObj.info])
    if dropColumns is not None:
        tickInfo = tickInfo.drop(columns=dropColumns, errors='ignore')
    elif keepColumns is not None:
        safeKeeps = [col for col in keepColumns if col in tickInfo.columns]
        tickInfo = tickInfo[safeKeeps]

    return tickInfo

def runThreadedJobs(keys, func, argDictionary=None, argsGlobal={}, dictionaryOut={}):
    """Runs 16 threads in parallel, 
    takes keys a list of keys.
    takes fun a function
    takes argDictionary a dictionary containing arguments relating to each key,
    takes argsGlobal a dictionary of args to apply for all keys.
    takes dictionaryOut a dictionary to store the output in.
    if there is no argDict, the keys are used as the argument."""
    job_queue = Queue()
    threads = []
    num_workers = 16

    if argDictionary == None:
        for key in keys:
            job_queue.put([dictionaryOut, key, func, key, argsGlobal])
    else:
        for key in keys:
            job_queue.put([dictionaryOut, key, func, argDictionary[key], argsGlobal])
        
    for _ in range(num_workers):
        thread = Thread(target=threadWorker, args=(job_queue,), daemon=True)
        thread.start()
        threads.append(thread)

    job_queue.join()
    for _ in range(num_workers):
        job_queue.put([None, None, None, None, None])

    for thread in threads:
        thread.join()

    return dictionaryOut


