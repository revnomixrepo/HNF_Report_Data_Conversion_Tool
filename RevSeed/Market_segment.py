import pandas as pd

# data = pd.read_excel("E:\Yadnesh\Revseed_HNF\MS_data/TheResort_MS.xls",header=None,skipfooter=4,skiprows=3)
# mapfile = pd.read_excel("E:\Yadnesh\Revseed_HNF\Mapping/map_col.xlsx")
# data_col = data.iloc[0]
# data.columns = data_col
# data = data[1:]
# data1 = pd.Series(data.columns)
# data1 = data1.fillna('unnamed:' + (data1.groupby(data1.isnull()).cumcount() + 1).astype(str))
# data1 = data1.to_list
data = pd.read_excel("E:\Yadnesh\Revseed_HNF\MS_data/TheResort_MS.xls", header=None, skipfooter=4, skiprows=4)
mapfile = pd.read_excel("E:\Yadnesh\Revseed_HNF\Mapping/map_col.xlsx")
# data_col = data.iloc[0]
# data.columns = data_col
# data = data[1:]
# data1 = pd.Series(data.columns)
# data1 = data1.fillna('unnamed:' + (data1.groupby(data1.isnull()).cumcount() + 1).astype(str))
data = data.rename(columns=mapfile.standard)
data1 = data.iloc[:, :-3]

data1.set_index(['Date'], inplace=True)
df_len = len(data1.columns)
df_ = data1.iloc[:, 0:3]

list_df = []
i = 0
while(i<df_len):
    data_1 = data1.iloc[:, i:i+3]
    data_1 = pd.DataFrame(data_1)
    i = i +3
    data_2 = data_1.reset_index()
    new_header = list(data_2.columns.values)[1]
    data_col = data_2.iloc[0]
    data_2.columns = data_col
    data_2 = data_2[1:]
    data_2["MS_cluster"] = new_header
    data_2.columns = data_2.columns.fillna('Date')
    list_df.append(data_2)
data_2 = pd.DataFrame(pd.concat(list_df, sort=False))
data_2.to_excel('E:/MS_data.xlsx', index=False)

# a = list(data1.columns.unique())
# list_df = []

# for i in a:
#     data1 = data1[data1.columns == i]
#     data1 = pd.DataFrame(data1)
# list_df.append(data1)
# data12 = pd.DataFrame(pd.concat(list_df, sort=False))
# data12.to_excel('E:/MS_data.xlsx',index=False)



# ren_dict = dict(zip(mapfile.columns, mapfile.standard))
# data = pd.DataFrame(data,columns=mapfile.before.to_list())
# data = pd.DataFrame(data, columns=mapfile.columns.to_list())
#
# data.rename(columns=ren_dict, inplace=True)

# data1 = data.T
#
# data_ = data1.iloc[0]
# data_ = data_.dropna()
# data_ = data_.to_frame()
# data_ = data_.rename(columns = {0:"Date"})
# data1.columns = data1.iloc[0]
# # data1 = data1.rename({'nan':'cluster','nan':'para'},axis = 1)


# data1 = data1.rename(columns = {data1.columns[0]:"clusters"})
# data1 = data1.rename(columns = {data1.columns[1]:"param"})

# data1 = data1[1:]
# data1 = data1.rename(columns)
# col = data.iloc[0]
# col = col.dropna()
# col = col.to_frame()
# ms_clust = col.rename(columns = {0:"cluters"})
# list_df = []
# for i in ms_clust:
#     df = data1[data1[""]]
