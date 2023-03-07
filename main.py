import pandas as pd
import numpy as np
import functionss as func
import os
from urllib.request import build_opener, install_opener, urlretrieve
from datetime import datetime
# import gdown
import camelot
import re
from functions_ZK import *
from config import *
from get_outlook_emails import OutlookTools
import win32com.client

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
cols = list_of_dfs[0].df.iloc[0]
df_conc = pd.DataFrame(columns=list_of_dfs[0].df.columns)

for i in range(len(list_of_dfs)):
    df_conc = pd.concat([df_conc, list_of_dfs[i].df.drop(list_of_dfs[i].df.index[0], axis=0)], ignore_index=True)
df_conc.columns = cols

# data cleaning
df_conc.columns = df_conc.columns.str.replace('\n', '')

### fix overflowed values ###

# fix Odhadovaná hodnota
df_conc['extracted_values_odhad_hodnota'] = df_conc['Popis'].str.extract(r'(\d+,\d+)')
df_conc['Odhadovaná hodnota'] = np.where(
    (df_conc['Odhadovaná hodnota'] == '') & (df_conc['Popis'].str.match(r'.*\d+,\d+')),
    df_conc['extracted_values_odhad_hodnota'], df_conc['Odhadovaná hodnota'])

# fix Cislo zmluvy
df_conc['extracted_values_cislo_zmluvy'] = df_conc['Dátum vyhotovenia'].str.extract(r'(^Z.+)')
df_conc['Dátum vyhotovenia'] = df_conc['Dátum vyhotovenia'].str.replace(r'^Z.+', '', regex=True)
df_conc['Číslo zmluvy'] = np.where(
    (df_conc['Číslo zmluvy'] == '') & (pd.isna(df_conc['extracted_values_cislo_zmluvy'])==False),
    df_conc['extracted_values_cislo_zmluvy'], df_conc['Číslo zmluvy'])

# fix Dodavatel - ico
df_conc['extracted_values_ico'] = df_conc['Dodávateľ - adresa'].str.extract(r'(\d{8})')
df_conc['Dodávateľ - adresa'] = df_conc['Dodávateľ - adresa'].str.replace(r'(\d{8})', '', regex=True)
df_conc['Dodávateľ - IČO'] = np.where(
    (df_conc['Dodávateľ - IČO'] == '') & (pd.isna(df_conc['extracted_values_ico']) == False),
    df_conc['extracted_values_ico'], df_conc['Dodávateľ - IČO'])

# fix Dodavatel-adresa
df_conc['extracted_values_dod_adresa'] = df_conc['Dodávateľ - názov'].str.extract(r'(Pribylinsk.+)')
df_conc['Dodávateľ - názov'] = df_conc['Dodávateľ - názov'].str.replace(r'(Pribylinsk.+)', '', regex=True)
df_conc['Dodávateľ - adresa'] = np.where(
    (df_conc['Dodávateľ - adresa'] == '') & (pd.isna(df_conc['extracted_values_dod_adresa']) == False),
    df_conc['extracted_values_dod_adresa'], df_conc['Dodávateľ - adresa'])

# fix Dodavatel - nazov
df_conc['extracted_values_dod_nazov'] = df_conc['Dátum vyhotovenia'].str.extract(r'([a-zA-Z].+)')
df_conc['Dátum vyhotovenia'] = df_conc['Dátum vyhotovenia'].str.replace(r'([a-zA-Z].+)', '', regex=True)
df_conc['Dodávateľ - názov'] = np.where((df_conc['Dodávateľ - názov']=='') & (pd.isna(df_conc['extracted_values_dod_nazov'])==False),
                                        df_conc['extracted_values_dod_nazov'], df_conc['Dodávateľ - názov'])

# fix Datum vyhotovenia
df_conc['Dátum vyhotovenia'] = df_conc['Dátum vyhotovenia'].str.replace(r'\n', '', regex=True)
df_conc['extracted_values_dat_vyhot'] = df_conc['Dodávateľ - názov'].str.extract(r'(\d+\.\d+.\d{4})')
df_conc['Dodávateľ - názov'] = df_conc['Dodávateľ - názov'].str.replace(r'\d+\.\d+.\d{4}', '', regex=True)
df_conc['Dodávateľ - názov'] = df_conc['Dodávateľ - názov'].str.strip()

df_conc['Dátum vyhotovenia'] = np.where(
    ((df_conc['Dátum vyhotovenia'] == '') ) & (pd.isna(df_conc['extracted_values_dat_vyhot'])==False),
    df_conc['extracted_values_dat_vyhot'], df_conc['Dátum vyhotovenia'])

df_conc = func.clean_str_cols(df_conc)


# 5 - Detská psychiatrická liecebna n.o. Hráň
urlretrieve(dict['detska psychiatricka liecebna n o hran']['objednavky_faktury_link'],
            data_path + current_date_time.strftime("%d-%m-%Y") + str(keysList[5]).replace(" ", "_") +
            dict['detska psychiatricka liecebna n o hran']['objednavky_faktury_file_ext'])

# 6 - Fakultna nemocnica Nitra
df_fnnr = pd.read_html(dict['fakultna nemocnica nitra']['objednavky_faktury_link'], encoding='utf-8')[0]
df_fnnr = func.clean_str_cols(df_fnnr)

# 7 - fakultna nemocnica s poliklinikou f d roosevelta banska bystrica

output = FNsP_BB_objednavky(link=dict['fakultna nemocnica s poliklinikou f d roosevelta banska bystrica']['objednavky_faktury_link'],
                            search_by='nazov_dodavatela', value='Intermedical s.r.o.', name=str(keysList[7]).replace(" ", "_"))
if output[0] == 'fail':
    print('First attempt failed. Trying again.')
    output = FNsP_BB_objednavky()
if output[0] == 'ok':
    data = output[1]

#############################################################################################################
# Data handling (from scraped data)
#############################################################################################################

cols = np.delete(df.columns.values, 0)
# columns from dictionary
columns_to_insert = ['100percent', 'financovaneMZSR', 'spoluzakladatelNO', 'VUC', 'emaevo', 'nazov 2022',
                     'riaditeliaMAIL_2022', 'zaujem_co_liekov', 'poznamky', 'chceme', 'zverejnovanie_objednavok_faktur_rozne', 'nazov']

objednavky_all = pd.DataFrame(columns=cols)


objednavky_all = pd.concat([objednavky_all, create_standardized_table('detska fakultna nemocnica s poliklinikou banska bystrica', df_conc, cols, columns_to_insert)], ignore_index=True)
objednavky_all = pd.concat([objednavky_all, create_standardized_table('fakultna nemocnica nitra', df_fnnr, cols, columns_to_insert)], ignore_index=True)


# insert last update date
objednavky_all['insert_date'] = datetime.now()
objednavky_all.to_excel('output.xlsx')


#############################################################################################################
# Data handling (from mails)
#############################################################################################################

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
otl = OutlookTools(outlook)

path = outlook.Folders['obstaravanie'].Folders['Doručená pošta'].Folders['Priame objednávky']

# FNsPZA
search_result = otl.find_message(path, "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%fnspza.sk' ")
hosp_path = data_path + "fnspza\\"
otl.save_attachement(hosp_path, search_result)

# load downloaded folders
all_tables = []
for file_name in os.listdir(hosp_path):
    print(file_name)
    if file_name.split(sep='.')[-1] in ('pdf', 'png', 'jpeg'):
        continue
    elif file_name.split(sep='.')[-1] == 'ods':
        df = pd.read_excel(os.path.join(hosp_path, file_name), engine='odf', sheet_name= None)
    else:
        df = func.load_df(name=file_name, path=hosp_path, sheet_name= None)
    all_tables.append([file_name, df])

# remove rows outside of table
for i in range(len(all_tables)):
    print(all_tables[i][0])
    doc = all_tables[i][-1]
    for key, value in doc.items():
        if any(s in '|'.join(map(str, doc[key].columns)) for s in ('Unnamed', 'nan')):
            doc[key] = doc[key].dropna(thresh=int(len(doc[key].columns)/3)).reset_index(drop=True)
            doc[key] = doc[key].dropna(axis=1, how='all')
            doc[key].columns = doc[key].columns.replace('\n', '')
            if not doc[key].empty:
                doc[key].columns = doc[key].iloc[0]
                doc[key] = doc[key].drop(doc[key].index[0])
        if not doc[key].empty:
            all_tables[i].append(doc[key])
            print(doc[key].columns)
















