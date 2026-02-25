def runCheck():
    import pandas as pd
    df = pd.read_csv('./data/master/ocean_solutions_master.csv',parse_dates=['date'])
    maxDate = df['date'].max().normalize()
    today = pd.Timestamp.today().normalize()
    return maxDate < today

if runCheck() == True:
    from utils.settles import main as ocean
    df = ocean()
