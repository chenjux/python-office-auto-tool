import pandas as pd

# -*- coding: utf-8 -*-
# df1 df2 df3 empty
# Get information replace to df2, copy df2 to df3, go back to the first step

# loop ends export df3

pd.set_option('display.max_rows', 1000)

# _____________________________________________________________________
df3 = pd.DataFrame()  # create an empty df
df1 = pd.read_csv('data.csv', encoding='gbk')  # need to put df1's data into df2

# List1: make df1 into list
# Make sure the beneath columns can pair with your data.csv file columns
list1 = [df1['cn_name'], df1['server_name'], df1['phone'], df1['contact'],
         df1['keyword'], df1['service_order'], df1['intro'], df1['address1'], df1['address2']]
# List2: create a list of the name that is used in df2
list2 = ['cn_name', 'server_name', 'phone', 'contact', 'keyword', 'service_order', 'intro', 'address1', 'address2']

# -----------------------------------------------------------------------------------
# step1

for v in range(len(list1[0])):
    #Extract information from cmd_template.csv to df2
    df2 = pd.read_csv('cmd_template.csv', encoding='gbk')
    # Keep insert the data into df3(by repeat 'cmd_template.csv' format)
    for i, j in enumerate(list1):
        # print(list2[i], j[v])
        df2.replace(to_replace=list2[i], value=j[v], inplace=True)
    df3 = df3.append(df2, ignore_index=True)

print(df3)
#export command to new_cmd.csv
df3.to_csv('new_cmd.csv', encoding='utf-8_sig')
