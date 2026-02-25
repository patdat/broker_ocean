


import win32com.client
import pandas as pd
import numpy as np
import datetime
import re
import os
from tabula.io import read_pdf  # for ocean

from utils.cloud import main as cloud
from utils.shorten_csv import processBroker

import warnings

warnings.filterwarnings(
    "ignore", 
    category=FutureWarning, 
    message=".*errors='ignore' is deprecated.*"
)


# %% [markdown]
# # EMAIL


# %%
def OS_download(dayStart):
    codePath = os.getcwd()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        codePath = os.getcwd()
        start_date = datetime.datetime.now() - datetime.timedelta(days=dayStart)
        messages = inbox.Items.Restrict(
            "[ReceivedTime] >= '{0}'".format(start_date.strftime("%m/%d/%Y %H:%M %p"))
        )
        lst_subject = []

        for message in messages:
            for attachment in message.Attachments:
                try:
                    if attachment.FileName.startswith(
                        "OCEAN"
                    ) and attachment.FileName.endswith(".pdf"):
                        filename = message.Subject
                        match = re.search(r"\d+", filename)
                        if match:
                            date_str = match.group()
                            if len(date_str) == 6:
                                date_obj = datetime.datetime.strptime(
                                    date_str, "%m%d%y"
                                )
                            else:
                                date_obj = datetime.datetime.strptime(
                                    date_str, "%m%d%Y"
                                )
                            formatted_date = date_obj.strftime("%Y.%m.%d")
                            print("Formatted date:", formatted_date)
                        else:
                            print("No digits found.")
                        fullname = os.path.join(
                            codePath, "./data/settles", formatted_date + ".pdf"
                        )
                        attachment.SaveAsFile(fullname)
                        lst_subject.append(formatted_date)
                except Exception as e:
                    print(f"Error processing attachment: {str(e)}")

    except Exception as e:
        print(f"Error accessing Outlook: {str(e)}")


# OS_download(5) #TEST CASE


# %% [markdown]
# # PROCESSING

# %% [markdown]
# ## HELPER


# %%
class frequency:
    def other(self):
        ocean = ["spot", "balmo"]
        period = ["spot", "balmo"]
        periodType = ["Spot", "Balmo"]
        df = pd.DataFrame({"ocean ": ocean, "period": period, "periodType": periodType})
        return df

    def monthly(self):
        date_range = pd.date_range(start="2015-01-01", end="2030-12-01", freq="MS")
        formatted_dates = [date.strftime("%b-%y").lower() for date in date_range]
        date_range = [date.strftime("%#m/%#d/%Y") for date in date_range]
        df = pd.DataFrame(
            {"ocean ": formatted_dates, "period": date_range, "periodType": "Months"}
        )
        return df

    def quarterly(self):
        date_range = pd.date_range(start="2015-01-01", end="2030-12-31", freq="QE")
        ocean_list = [f"q{date.quarter}-{str(date.year)[-2:]}" for date in date_range]
        format_list = [f"{str(date.year)[-2:]}Q{date.quarter}" for date in date_range]
        df = pd.DataFrame(
            {"ocean ": ocean_list, "period": format_list, "periodType": "Quarters"}
        )
        return df

    def yearly(self):
        year_list = list(range(2015, 2030))
        c_list = ["C" + str(year)[-2:] for year in year_list]
        year_list = ["cal-" + str(year)[-2:] for year in year_list]
        df = pd.DataFrame(
            {"ocean ": year_list, "period": c_list, "periodType": "Years"}
        )
        return df

    def main(self):
        df0 = self.other()
        df1 = self.monthly()
        df2 = self.quarterly()
        df3 = self.yearly()
        df = pd.concat([df0, df1, df2, df3])
        df.columns = ["period", "newName", "periodType"]
        return df


oceanLookup = frequency().main()


# %%
def convertPeriod(fileName):
    dateConversion = pd.read_csv("lookup/ocean_period_conversion.csv")
    fileName = fileName.split(".")
    fileDate = datetime.date(int(fileName[0]), int(fileName[1]), int(fileName[2]))

    beginning_of_month = fileDate.replace(day=1)
    formatted_beginning_of_month = beginning_of_month.strftime("%d/%m/%Y")

    dateConversion.loc[0, "newName"] = formatted_beginning_of_month
    dateConversion.loc[1, "newName"] = formatted_beginning_of_month

    dateConversion["newName"] = pd.to_datetime(
        dateConversion["newName"], errors="coerce"
    ).dt.strftime("%Y-%m-%d")

    return dateConversion


# convertedDate = convertPeriod('2023.10.12.pdf') #TEST CASE

# %% [markdown]
# ## PROCESS FILES

# %%
inputFolder = "./data/settles/"
fileName = "2023.10.12.pdf"
fullPath = inputFolder + fileName


# %%
def get_tradedate(filename):
    """the purpose of this function is to get the date from the filename in a datetime object"""
    # get the filename without the extension
    filename = os.path.basename(filename)
    # first 10 characters of the filename
    date_str = filename[:10]
    # convert to datetime object
    date_obj = datetime.datetime.strptime(date_str, "%Y.%m.%d")
    date_obj = date_obj.strftime("%m/%d/%Y")
    return date_obj


trade_date = get_tradedate(fileName)


# %%
def get_data(filename):
    elements = read_pdf(filename, pages="all", lattice=True)
    elements = [ele for ele in elements if not ele.empty]
    for ele in elements:
        ele.dropna(axis=1, how="all", inplace=True)
    return elements


# elements = get_data(fileName)


# %%
def td3c_td20(raw, route):
    """MEG to China, WAF to UKC"""
    df = raw.copy()
    # drop last 2 columns. the columns are not necessary
    df = df.iloc[:, :-2]
    # rename columns
    df.columns = ["period", "ws", "price"]
    # drop last 2 rows
    df = df.iloc[:-2, :]
    # drop row if == 'Contract'.
    df = df[df["period"] != "Contract"]
    df.reset_index(drop=True, inplace=True)

    # this will ensure that balmo which is error at end of month is np.nan if it is '#DIV/0!'
    if df.loc[df["period"] == "Balmo", "ws"].values == "#DIV/0!":
        df.loc[df["period"] == "Balmo", "ws"] = np.nan
        df.loc[df["period"] == "Balmo", "price"] = np.nan

    # make a $/t spot value
    df.iloc[0, 0] = "Spot"
    df.iloc[0, 2] = np.nan
    df["ws"] = df["ws"].astype(float)
    df["price"] = df["price"].astype(float)
    df.iloc[0, 2] = round(df.iloc[0, 1] / (df.iloc[2, 1] / df.iloc[2, 2]), 4)
    # round 4 decimal places

    # df['period'] = pd.to_datetime(df['period'], errors='ignore')

    df["period"] = df["period"].astype(str)
    df["period"] = df["period"].str.lower()
    df["instrument"] = route

    return df


# %%
def td22(raw, route):
    """USGC to China"""
    df = raw.copy()
    # drop last 2 columns. the columns are not necessary
    df = df.iloc[:, :-2]
    # rename columns
    df.columns = ["period", "ws", "price"]
    # drop last 2 rows
    df = df.iloc[:-2, :]
    # drop row if == 'Contract'.

    # convert df['period] to string
    df["period"] = df["period"].astype(str)
    # make all lower case
    df["period"] = df["period"].str.lower()
    # remove leading spaces from df['period']
    df["period"] = df["period"].str.lstrip()
    # remove trailing spaces from df['period']
    df["period"] = df["period"].str.rstrip()
    # remove double spaces from df['period']
    df["period"] = df["period"].str.replace("  ", " ")

    df = df[df["period"] != "contract"]
    df.reset_index(drop=True, inplace=True)

    # this will ensure that balmo which is error at end of month is np.nan if it is '#DIV/0!'
    if df.loc[df["period"] == "balmo", "ws"].values == "#DIV/0!":
        df.loc[df["period"] == "balmo", "ws"] = np.nan
        df.loc[df["period"] == "balmo", "price"] = np.nan

    # #make a $/t spot value
    df.iloc[0, 0] = "spot"
    df.iloc[0, 2] = df.iloc[1, 1]

    df = df[df["period"] != "$/mt settle"]

    df["ws"] = df["ws"].astype(float)
    df["price"] = df["price"].astype(float)

    df["instrument"] = route

    return df


# %%
def td25(raw, route):
    """USGC to UKC"""
    df = raw.copy()
    # drop last 2 columns. the columns are not necessary
    df = df.iloc[:, :-2]
    # rename columns
    df.columns = ["period", "ws", "price"]
    # drop last 2 rows
    df = df.iloc[:-2, :]
    # drop row if == 'Contract'.

    # convert df['period] to string
    df["period"] = df["period"].astype(str)
    # make all lower case
    df["period"] = df["period"].str.lower()
    # remove leading spaces from df['period']
    df["period"] = df["period"].str.lstrip()
    # remove trailing spaces from df['period']
    df["period"] = df["period"].str.rstrip()
    # remove double spaces from df['period']
    df["period"] = df["period"].str.replace("  ", " ")

    df = df[df["period"] != "contract"]
    df.reset_index(drop=True, inplace=True)

    # this will ensure that balmo which is error at end of month is np.nan if it is '#DIV/0!'
    if df.loc[df["period"] == "balmo", "ws"].values == "#DIV/0!":
        df.loc[df["period"] == "balmo", "ws"] = np.nan
        df.loc[df["period"] == "balmo", "price"] = np.nan

    # #make a $/t spot value
    df.iloc[0, 0] = "spot"
    df.iloc[0, 2] = df.iloc[1, 1]

    df = df[df["period"] != "$/mt settle"]
    df = df[df['period'] != 'exc $/mt\rsettle']

    df["ws"] = df["ws"].astype(float)
    df["price"] = df["price"].astype(float)

    df["instrument"] = route

    return df


# %%
def quarterlry_processing(df):
    def quarterly_calculation(df, route):
        df = df.copy()
        df = df[df["instrument"] == route]
        df = df[df["periodType"] == "Quarters"]
        # Convert the 'period' column to a datetime type
        df["period"] = pd.to_datetime(df["period"])

        # Set the 'period' column as the index
        df.set_index("period", inplace=True)

        # Resample to monthly frequency and forward fill
        df_monthly = df.resample("MS").ffill()

        # Create a new index with 2 additional months
        new_index = pd.date_range(
            start=df_monthly.index.min(), periods=len(df_monthly) + 2, freq="MS"
        )

        # Reindex the DataFrame
        df_monthly = df_monthly.reindex(new_index).ffill()

        # Reset index to move 'period' back as a column
        df_monthly.reset_index(inplace=True)
        df_monthly = df_monthly.rename(columns={"index": "period"})

        # Update 'periodType' to 'Monthly'
        df_monthly["periodType"] = "Months"
        # rearrange columns to be in the order source, periodType, date, instrument, period, price, ws
        df_monthly = df_monthly[
            ["source", "periodType", "date", "instrument", "period", "price", "ws"]
        ]
        return df_monthly

    df1 = quarterly_calculation(df, "TD3C")
    df2 = quarterly_calculation(df, "TD20")
    df3 = quarterly_calculation(df, "TD22")
    df4 = quarterly_calculation(df, "TD25")
    dff = pd.concat([df1, df2, df3, df4], axis=0)
    df = pd.concat([df, dff], axis=0)
    # drop duplicates subset is periodType, date, instrument, period. keep first
    df.drop_duplicates(
        subset=["periodType", "date", "instrument", "period"],
        keep="first",
        inplace=True,
    )
    # filter out periodType == 'Quarters'
    df = df[df["periodType"] != "Quarters"]

    return df


# %%
def yearly_processing(df):
    def yearly_calculation(df, route):
        df = df.copy()
        df = df[df["instrument"] == route]
        df = df[df["periodType"] == "Years"]

        # Convert the 'period' column to a datetime type
        df["period"] = pd.to_datetime(df["period"])

        # Set the 'period' column as the index
        df.set_index("period", inplace=True)

        # Create an empty DataFrame for the extrapolated data
        df_extrapolated = pd.DataFrame()

        # For each row in the original DataFrame, generate 12 monthly entries
        for idx, row in df.iterrows():
            new_index = pd.date_range(start=idx, periods=12, freq="MS")
            temp_df = pd.DataFrame(index=new_index)
            for column in df.columns:
                temp_df[column] = row[column]
            df_extrapolated = pd.concat([df_extrapolated, temp_df])

        # Adjust the periodType and reset the index
        df_extrapolated["periodType"] = "Months"
        df_extrapolated.reset_index(inplace=True)
        df_extrapolated = df_extrapolated.rename(columns={"index": "period"})

        return df_extrapolated

    df1 = yearly_calculation(df, "TD3C")
    df2 = yearly_calculation(df, "TD20")
    df3 = yearly_calculation(df, "TD22")
    df4 = yearly_calculation(df, "TD25")
    dff = pd.concat([df1, df2, df3, df4], axis=0)
    df = pd.concat([df, dff], axis=0)
    # drop duplicates subset is periodType, date, instrument, period. keep first
    df.drop_duplicates(
        subset=["periodType", "date", "instrument", "period"],
        keep="first",
        inplace=True,
    )
    df = df[df["periodType"] != "Years"]
    return df


# %%
def compileAll(fileName):
    folderName = "./data/settles/"
    elements = get_data(folderName + fileName)
    trade_date = get_tradedate(folderName + fileName)
    df_td3c = td3c_td20(elements[0], "TD3C")
    df_td20 = td3c_td20(elements[1], "TD20")
    df_td22 = td22(elements[2], "TD22")
    df_td25 = td25(elements[3], "TD25")

    df = pd.concat([df_td3c, df_td20, df_td22, df_td25], axis=0)
    df = pd.merge(df, oceanLookup, how="left", on="period")
    df["date"] = trade_date
    df.drop("period", axis=1, inplace=True)
    df.rename(columns={"newName": "period"}, inplace=True)
    df["source"] = "OCEAN"
    df = df[["source", "periodType", "date", "instrument", "period", "price", "ws"]]

    convertedDate = convertPeriod(fileName)
    df["period"] = df["period"].map(convertedDate.set_index("oldName")["newName"])
    df["date"] = pd.to_datetime(df["date"], errors="coerce")

    df = quarterlry_processing(df)
    df = yearly_processing(df)
    df["period"] = pd.to_datetime(df["period"], errors="coerce")
    df.drop_duplicates(subset=['periodType', 'date', 'instrument', 'period'], keep='first', inplace=True)
    #reset index
    df.reset_index(drop=True, inplace=True)

    for index, row in df.iterrows():
        if row['periodType'] == 'Spot' or row['periodType'] == 'Balmo':
            df.at[index, 'period'] = row['date'].replace(day=1) 

    return df


# df = compileAll('2023.10.12.pdf') #TEST CASE


# %%
def create_master():
    folder = "./data/settles/"
    files = os.listdir(folder)
    df = pd.DataFrame()
    for file in files:
        dff = compileAll(file)
        df = pd.concat([df, dff], axis=0)
    return df


# df = create_master()
# df.to_csv('./data/master/ocean_solutions_master.csv',index=False)


# %%
def run_daily(counter):
    folder = "./data/settles/"
    files = os.listdir(folder)
    files = files[-counter:]
    df = pd.read_csv(
        "./data/master/ocean_solutions_master.csv", parse_dates=["date", "period"]
    )
    for file in files:
        dff = compileAll(file)
        df = pd.concat([df, dff], axis=0)
        
    #remove duplicates from the master file
    
    
    df['price'] = pd.to_numeric(df['price'], errors='coerce')
    #force df['ws'] to numeric
    df['ws'] = pd.to_numeric(df['ws'], errors='coerce')
    
    df['date'] = pd.to_datetime(df['date'], errors='coerce')
    df['period'] = pd.to_datetime(df['period'], errors='coerce')
    
    cols = ['date','period','price','ws']
    df.dropna(subset=cols, inplace=True)

    df.drop_duplicates(subset=['periodType', 'date', 'instrument', 'period'], keep='last', inplace=True)
    
        
    return df


# df = run_daily(5) #TEST CASE


def pvData(df):
    df = df.copy()
    pv1 = pd.pivot_table(df, values='price', index=['source', 'periodType', 'date', 'period'], columns=['instrument'])
    pv1.columns = pv1.columns.map(lambda x: f'{x}_pmt')
    pv1 = pv1.reset_index()
    pv2 = pd.pivot_table(df, values='ws', index=['source', 'periodType', 'date', 'period'], columns=['instrument'])
    pv2.columns = pv2.columns.map(lambda x: f'{x}_ws')
    pv2.rename(columns={'TD22_ws': 'td22_ls'}, inplace=True)
    pv2 = pv2.reset_index()
    pv = pd.merge(pv1, pv2, how='left', on=['source', 'periodType', 'date', 'period'])
    pv.to_csv('./data/master/ocean_solutions_pivot.csv', index=False)
    cloud(pv, 'BROKER/MASTER', 'ocean_solutions_pivot.csv',df_index=False)


# %%
def main():
    OS_download(5)
    df = run_daily(5)
    processBroker(
        df,
        "./data/settles/",
        "ocean_solutions_master",
        masterFolder="./data/master/",
        cloudFolder="BROKER/MASTER",
    )
    pvData(df)

    return df


if __name__ == "__main__":
    df = main()
