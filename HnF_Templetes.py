import shutil
import pandas as pd
from datetime import datetime,timedelta
import os
import numpy as np
from send_email import send_Hnf
from auto_mail_download import emailAttachment

# emailAttachment()

def process_HNF(infile):

    data = pd.read_excel(infile,skiprows=2,skipfooter=3)

    # try:
    #     data = data.rename(columns ={"Rms/Avl.":"Capacity","Total Occ":"Rooms Sold","Room Rev":"Revenue"})
    #     data["Availability"] = data["Capacity"] - data["Rooms Sold"]
    # except:
    #     pass
    #     try:
    #         data = data.rename(columns={"Rooms": "Capacity", "Rooms.1": "Rooms Sold", "OOO": "out_of_order"})
    #         data["Availability"] = data["Capacity"] - (data["Rooms Sold"] + data["out_of_order"])
    #     except:
    #         data = data.rename(columns={"Total Occ.": "Rooms Sold"})
    #         data["Capacity"] = 80
    #         data["Availability"] = data["Capacity"] - data["Rooms Sold"]

    try:
        data = data.rename(columns ={"Rms/Avl.":"Capacity","Total Occ":"Rooms Sold","Room Rev":"Revenue"})
        data["Availability"] = data["Capacity"] - data["Rooms Sold"]
    except:
        pass
        try:
            data = data.rename(columns={"Rooms": "Capacity", "Rooms.1": "Rooms Sold", "OOO": "out_of_order"})
            data["Availability"] = data["Capacity"] - (data["Rooms Sold"] + data["out_of_order"])
        except:
            data = data.rename(columns={"Rooms": "Capacity", "Rooms.1": "Rooms Sold"})
            data["Availability"] = data["Capacity"] - data["Rooms Sold"]

    # data["Availability"] = data["Capacity"] - data["Rooms Sold"]
    data = data[["Date","Capacity",	"Rooms Sold","Availability","Revenue"]]
    # data = data.sort_values(['Date'])
    # data = data.dropna(subset = ["Date"],how = all)
    data = data.dropna(how="all", thresh=4)
    data['Revenue']= (data['Revenue'].astype(str)).str.replace(',', '').astype(float)
    patternDel = "__"
    filter = data['Date'].str.contains(patternDel)
    data = data[~filter]

################# date range from today to next 365 days in dataframe ####################
    # data1 = pd.date_range(datetime.today(), periods=365)
    # data1= pd.DataFrame(data1)
    # data1 = data1.rename(columns={0: "Date"})      ##### comment on 7thjune2022

    try:
        data["Date"] = data["Date"].apply(lambda x: datetime.strptime(x, '%d-%b-%Y'))
    except:
        data["Date"] = data["Date"].apply(lambda x: datetime.strptime(x, '%d/%m/%Y'))

    data["Date"] = data["Date"].dt.date
    # data1["Date"] = data1["Date"].dt.date

############# merge the data with date range #############
    # data = data1.merge(data, how="left", on="Date")

##########################################################
    data["Capacity"] = data["Capacity"].fillna(method="ffill")
    data = data.fillna(0)
    data["Availability"] = np.where(data["Availability"] == 0,
                                    data["Capacity"] - data["Rooms Sold"],
                                    data["Availability"])
    data = data.astype({"Capacity":"int","Rooms Sold":"int","Availability":"int","Revenue":"int"})

    return data


if __name__=="__main__":
    today = datetime.today().date()
    filename = 'Hotel HNF'
    std_path = r"C:\Revseed_HNF/"
    in_path = std_path + "/Input/"
    outpath = std_path + "/Output/" + str(today)
    recyclebin = std_path + "/recyclebin/" + "{}_{}.xlsx"

    if os.path.isdir(in_path):
        files = os.listdir(in_path)
        if len(files) > 0:
            for i in files:
                if i.__contains__("HNF"):
                    infile = in_path + "/" + i
                    code = (infile.split("_")[-1]).split(".")[0]
                    final_df = process_HNF(infile)
                if os.path.isdir(outpath):
                    print(filename + " " + "Data is already Present" )
                else:
                    os.mkdir(outpath)
                final_df.to_excel(outpath + "\HNF_{}.xlsx".format(code),index=False)
                send_Hnf(code, today)
                print("HNF_{} is Dumped and send Successfully".format(code))
                # shutil.move(infile,recyclebin)
                t = datetime.today().strftime("%d_%Y %H_%M_%S %p")
                os.replace(infile, recyclebin.format(code,t))
            else:
                print("Data file is not Found")