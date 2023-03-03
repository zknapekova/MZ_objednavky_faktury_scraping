import pandas as pd
import numpy as np
import functionss as func
import os
from urllib.request import build_opener, install_opener, urlretrieve
from datetime import datetime
# import gdown
import tabula
import camelot
import re
from functions_ZK import *
from config import *

current_date_time = datetime.now()
print('Start:', current_date_time)

# load excel and create dictionary
df = func.load_df(source_path)
df['nazov']=df['Nazov_full']

# clean data
df = func.clean_str_cols(df, cols=['Nazov_full'])
df['Nazov_full'] = df['Nazov_full'].replace(',|\.', '', regex=True)
dict = df.set_index('Nazov_full').T.to_dict('dict')

keysList = list(dict.keys())

# update
dict = update_dict(dict)

#############################################################################################################
# Web scraping
#############################################################################################################

# set up opener
opener = build_opener()
opener.addheaders = [('User-agent', 'Mozilla/5.0')]
install_opener(opener)

# 0 - Centrum pre liečbu drogových závislostí Banská\xa0Bystrica - link for scraping not available


# 1 - Centrum pre liečbu drogových závislostí Bratislava
urlretrieve(dict['centrum pre liecbu drogovych zavislosti bratislava']['objednavky_faktury_link'],
            data_path + current_date_time.strftime("%d-%m-%Y") + str(keysList[1]).replace(" ", "_") + '.xlsx')

# 2 - Centrum pre liečbu drogových závislostí Košice
urlretrieve(dict['centrum pre liecbu drogovych zavislosti kosice']['objednavky_faktury_link'],
            data_path + current_date_time.strftime("%d-%m-%Y") + str(keysList[2]).replace(" ", "_") + '.xlsx')

# 3 - Detská fakultná nemocnica Košice
# TODO
# gdown.download(dict['Detská fakultná nemocnica Košice']['objednavky_faktury_link'], data_path+current_date_time.strftime("%d-%m-%Y")+str(keysList[3]).replace(" ", "")+'.xlsx')

# 4 - Detská fakultná nemocnica s poliklinikou Banská Bystrica
file_name = data_path + current_date_time.strftime("%d-%m-%Y") + str(keysList[4]).replace(" ", "_") + \
            dict['detska fakultna nemocnica s poliklinikou banska bystrica']['objednavky_faktury_file_ext']
urlretrieve(dict['detska fakultna nemocnica s poliklinikou banska bystrica']['objednavky_faktury_link'], file_name)

list_of_dfs = camelot.read_pdf(file_name, pages='all')
df_conc = pd.DataFrame(columns=list_of_dfs[0].df.iloc[0])

for i in range(len(list_of_dfs)):
    tab = list_of_dfs[i].df
    tab.rename(columns=tab.iloc[0], inplace=True)
    tab.drop(tab.index[0], inplace=True)
    df_conc = pd.concat([df_conc, tab], ignore_index=True)



# 5 - Detská psychiatrická liecebna n.o. Hráň
urlretrieve(dict['detska psychiatricka liecebna n o hran']['objednavky_faktury_link'],
            data_path + current_date_time.strftime("%d-%m-%Y") + str(keysList[5]).replace(" ", "_") +
            dict['detska psychiatricka liecebna n o hran']['objednavky_faktury_file_ext'])

# 6 - Fakultna nemocnica Nitra
table = pd.read_html(dict['fakultna nemocnica nitra']['objednavky_faktury_link'])[0]


# 7 - fakultna nemocnica s poliklinikou f d roosevelta banska bystrica

output = FNsP_BB_objednavky(link=dict['fakultna nemocnica s poliklinikou f d roosevelta banska bystrica']['objednavky_faktury_link'],
                            search_by='nazov_dodavatela', value='Intermedical s.r.o.', name=str(keysList[7]).replace(" ", "_"))
if output[0] == 'fail':
    print('First attempt failed. Trying again.')
    output = FNsP_BB_objednavky()
if output[0] == 'ok':
    data = output[1]

#############################################################################################################
# Data handling
#############################################################################################################

cols = np.delete(df.columns.values, 0)
objednavky_all = pd.DataFrame(columns=cols)

df_conc.columns = df_conc.columns.str.replace('\r', ' ')
df_conc.columns = df_conc.columns.str.replace('vyhotoveni a', 'vyhotovenia')

# insert scraped data
for i in objednavky_all.columns.values:
    for j in range(len(df_conc.columns.values)):
        if dict['detska fakultna nemocnica s poliklinikou banska bystrica'][i] == df_conc.columns.values[j]:
            print(dict['detska fakultna nemocnica s poliklinikou banska bystrica'][i], df_conc.columns.values[j])
            objednavky_all[i] = df_conc[df_conc.columns[j]]

# insert data from dictionary
columns_to_insert = ['100percent', 'financovaneMZSR', 'spoluzakladatelNO', 'VUC', 'emaevo', 'nazov 2022',
                     'riaditeliaMAIL_2022', 'zaujem_co_liekov', 'poznamky', 'chceme', 'zverejnovanie_objednavok_faktur_rozne', 'nazov']

for col_name in objednavky_all.columns.values:
    if col_name in columns_to_insert:
        objednavky_all[col_name]=dict['detska fakultna nemocnica s poliklinikou banska bystrica'][col_name]

# insert last update date
objednavky_all['insert_date'] = datetime.now()
objednavky_all.to_excel('output.xlsx')


# cleaning data
