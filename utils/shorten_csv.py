import pandas as pd
import datetime
from utils.cloud import main as cloud

def shorten_csv(df,dayCounter=30):
    df = df.copy()
    today = datetime.date.today()
    start = today - datetime.timedelta(days=dayCounter)
    start = pd.to_datetime(start)
    df = df[df['date'] >= start].copy()
    return df

def processBroker(df,inputFolder,masterName,masterFolder='./data/master/',cloudFolder='BROKER/MASTER'):
    df = df.copy()
    df.to_csv(masterFolder + masterName + '.csv', index=False)
    cloud(df,cloudFolder,masterName + '.csv',df_index=False)

    daysList = [60,30]
    for i in daysList:
        df = shorten_csv(df,i)
        df.to_csv('./data/shortened/' + masterName + f'_{i}.csv', index=False)
        cloud(df,cloudFolder,masterName + f'_{i}.csv',df_index=False)