import pandas as pd
import numpy as np
import functionss as func
import os
from urllib.request import build_opener, install_opener, urlretrieve
from datetime import datetime, date
import camelot
import re
from functions_ZK import *
from config import *
from schemas import *
import win32com.client
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.common.exceptions import ElementNotVisibleException, TimeoutException
from docx import Document

current_date_time = datetime.now()

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
otl = OutlookTools(outlook)
path = outlook.Folders['obstaravanie'].Folders['Doručená pošta'].Folders['Priame objednávky']

db = ObjednavkyDB(objednavky_db_connection)

#############################################################################################################
# Web scraping
#############################################################################################################

# set up opener
opener = build_opener()
opener.addheaders = [('User-agent', 'Mozilla/5.0')]
install_opener(opener)

# 0 - Centrum pre liečbu drogových závislostí Banská\xa0Bystrica - link for scraping not available


# 1 - Centrum pre liečbu drogových závislostí Bratislava
urlretrieve(dict_all['centrum pre liecbu drogovych zavislosti bratislava']['objednavky_faktury_link'],
            data_path + current_date_time.strftime("%d-%m-%Y") + str(keysList[1]).replace(" ", "_") + '.xlsx')

# 2 - Centrum pre liečbu drogových závislostí Košice
urlretrieve(dict_all['centrum pre liecbu drogovych zavislosti kosice']['objednavky_faktury_link'],
            data_path + current_date_time.strftime("%d-%m-%Y") + str(keysList[2]).replace(" ", "_") + '.xlsx')

# 3 - Detská fakultná nemocnica Košice
# TODO
# gdown.download(dict_all['Detská fakultná nemocnica Košice']['objednavky_faktury_link'], data_path+current_date_time.strftime("%d-%m-%Y")+str(keysList[3]).replace(" ", "")+'.xlsx')

# 4 - Detská fakultná nemocnica s poliklinikou Banská Bystrica
file_name = data_path + current_date_time.strftime("%d-%m-%Y") + str(keysList[4]).replace(" ", "_") + \
            dict_all['detska fakultna nemocnica s poliklinikou banska bystrica']['objednavky_faktury_file_ext']
urlretrieve(dict_all['detska fakultna nemocnica s poliklinikou banska bystrica']['objednavky_faktury_link'], file_name)

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
    (df_conc['Číslo zmluvy'] == '') & (pd.isna(df_conc['extracted_values_cislo_zmluvy']) == False),
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
df_conc['Dodávateľ - názov'] = np.where(
    (df_conc['Dodávateľ - názov'] == '') & (pd.isna(df_conc['extracted_values_dod_nazov']) == False),
    df_conc['extracted_values_dod_nazov'], df_conc['Dodávateľ - názov'])

# fix Datum vyhotovenia
df_conc['Dátum vyhotovenia'] = df_conc['Dátum vyhotovenia'].str.replace(r'\n', '', regex=True)
df_conc['extracted_values_dat_vyhot'] = df_conc['Dodávateľ - názov'].str.extract(r'(\d+\.\d+.\d{4})')
df_conc['Dodávateľ - názov'] = df_conc['Dodávateľ - názov'].str.replace(r'\d+\.\d+.\d{4}', '', regex=True)
df_conc['Dodávateľ - názov'] = df_conc['Dodávateľ - názov'].str.strip()

df_conc['Dátum vyhotovenia'] = np.where(
    ((df_conc['Dátum vyhotovenia'] == '')) & (pd.isna(df_conc['extracted_values_dat_vyhot']) == False),
    df_conc['extracted_values_dat_vyhot'], df_conc['Dátum vyhotovenia'])

df_conc = func.clean_str_cols(df_conc)

# 5 - Detská psychiatrická liecebna n.o. Hráň
urlretrieve(dict_all['detska psychiatricka liecebna n o hran']['objednavky_faktury_link'],
            data_path + current_date_time.strftime("%d-%m-%Y") + str(keysList[5]).replace(" ", "_") +
            dict_all['detska psychiatricka liecebna n o hran']['objednavky_faktury_file_ext'])

# 6 - Fakultna nemocnica Nitra
df_fnnr = pd.read_html(dict_all['fakultna nemocnica nitra']['objednavky_faktury_link'], encoding='utf-8')[0]
df_fnnr2 = pd.read_html('https://fnnitra.sk/objd/2022/', encoding='utf-8')[0]
df_fnnr = func.clean_str_cols(df_fnnr)

# 7 - fakultna nemocnica s poliklinikou f d roosevelta banska bystrica

output = FNsP_BB_objednavky(
    link=dict_all['fakultna nemocnica s poliklinikou f d roosevelta banska bystrica']['objednavky_faktury_link'],
    search_by='nazov_dodavatela', value='Intermedical s.r.o.', name=str(keysList[7]).replace(" ", "_"))
if output[0] == 'fail':
    print('First attempt failed. Trying again.')
    output = FNsP_BB_objednavky()
if output[0] == 'ok':
    data = output[1]

# 7 - fakultna nemocnica s poliklinikou zilina

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(chromedriver_path2, options=options)
driver.get(dict_all['fakultna nemocnica s poliklinikou zilina']['objednavky_link'])
n_records_dropdown = driver.find_element(By.ID, "limit1")

select = Select(n_records_dropdown)
select.select_by_value('100')
table_lst = WebDriverWait(driver, 10).until(
    EC.visibility_of_element_located((By.ID, "list_1_com_fabrik_1"))).get_attribute(
    "outerHTML")
result_df = pd.read_html(table_lst)[0]

while True:
    try:
        WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH,
                                                                   "//li[contains(@class, 'pagination-next')]//a[contains(@title, 'Nasled')]"))).click()
        next_table_lst = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, "list_1_com_fabrik_1"))).get_attribute(
            "outerHTML")
        next_table = pd.read_html(next_table_lst)[0]
        result_df = pd.concat([result_df, next_table], ignore_index=True)
    except TimeoutException as ex:
        print('Data were retrieved')
        break
    except:
        driver.quit()
        print('Retrieving data failed')
        break

for i in result_df.columns.values:
    if 'Unnamed' in i:
        result_df.drop(i, axis=1, inplace=True)

result_df = result_df.dropna(thresh=int(len(result_df.columns) / 3)).reset_index(drop=True)
result_df = result_df[result_df['Číslo objednávky'].str.match(r'(^Uk.+)') == False]
result_df2 = func.clean_str_cols(result_df)

result_df2 = func.load_df(os.path.join(data_path + "fnspza\\" + 'fnspza_all_web.pkl'), path=os.getcwd())
result_df2 = clean_str_col_names(result_df2)
result_df2.columns = ['objednavka_cislo1', 'cpv kod', 'objednavka_predmet', 'cena', 'pocet mj', 'kategoria', 'pridane']

result_df2['cena'] = result_df2['cena'].str.replace(r"[a-z|'|\s|-|\(\)]+", '', regex=True).str.replace(r",", '.',
                                                                                                       regex=True)
result_df2['cena'] = result_df2['cena'].astype(float)

result_df2['rok_objednavky'] = result_df2['objednavka_cislo1'].str.extract(r'(20\d{2})')
result_df2['objednavka_cislo'] = result_df2['objednavka_cislo1'].apply(lambda x: x.split('/')[-1])
result_df2['objednavka_cislo'] = result_df2['objednavka_cislo'].str.replace(r'^0+', '', regex=True)
result_df2['objednavka_cislo'] = result_df2['objednavka_cislo'].apply(
    lambda x: x.split('-')[-1] + '-' + x.split('-')[0] if ('-' in x) else x)

#############################################################################################################
# Data handling (from scraped data)
#############################################################################################################

cols = np.delete(df.columns.values, 0)
# columns from dictionary
columns_to_insert = ['100percent', 'financovaneMZSR', 'spoluzakladatelNO', 'VUC', 'emaevo', 'nazov 2022',
                     'riaditeliaMAIL_2022', 'zaujem_co_liekov', 'poznamky', 'chceme',
                     'zverejnovanie_objednavok_faktur_rozne', 'nazov']

objednavky_all = pd.DataFrame(columns=cols)

objednavky_all = pd.concat([objednavky_all,
                            create_standardized_table('detska fakultna nemocnica s poliklinikou banska bystrica',
                                                      df_conc, cols, columns_to_insert)], ignore_index=True)
objednavky_all = pd.concat(
    [objednavky_all, create_standardized_table('fakultna nemocnica nitra', df_fnnr, cols, columns_to_insert)],
    ignore_index=True)

# insert last update date
objednavky_all['insert_date'] = datetime.now()
objednavky_all.to_excel('output.xlsx')

#############################################################################################################
# Data handling (from mails)
#############################################################################################################

# FNsPZA - first load ###
fnspza = PriameObjednavkyMail('fnspza')

# ## download all attachements available from outlook
search_result = otl.find_message(path, "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%" + fnspza.hosp + '.sk' + "' ")

#otl.save_attachment(fnspza.hosp_path, search_result)
# load data
fnspza.load()
fnspza.clean_tables()

fnspza.all_tables_list_cleaned[2][3]['sukl_kod'] = fnspza.all_tables_list_cleaned[2][3]['sukl_kod1'].str.cat(
    fnspza.all_tables_list_cleaned[2][3]['kod'].astype(str), sep='')

#fnspza.data_check()
fnspza.create_table(stand_column_names=stand_column_names)

# data cleaning
fnspza.df_all = fnspza_data_cleaning(fnspza.df_all)
fnspza.create_columns_w_dict(key='fakultna nemocnica s poliklinikou zilina')

# ## save df
fnspza_df_search = pd.DataFrame(fnspza.df_all[fnspza.final_table_cols])
fnspza.save_tables(table=fnspza_df_search)

df_split_list = split_dataframe(fnspza.df_all.drop(['rok_objednavky', 'rok_objednavky_num'], axis=1), chunk_size=100000)
for i in df_split_list:
    db.insert_table(table_name='priame_objednavky', df=i, if_exists='append', index=False)

# ## FNNR - first load ###

fnnr = PriameObjednavkyMail('fnnitra')

search_result = otl.find_message(path,
                                 "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%" + fnnr.hosp + '.sk' + "' ")
otl.save_attachment(fnnr.hosp_path, search_result)

fnnr.load()
fnnr.clean_tables()
fnnr.data_check()
fnnr.create_table(stand_column_names=stand_column_names)

fnnr.create_columns_w_dict(key='fakultna nemocnica nitra')

# # data cleaning ##
# cena
dict_cena = {"[a-z]|'|\s|[\(\)]+": '', '[,|\.]-.*': '', '-,': '', ",": '.', '\.\.': '.', '/.*/*': '', '\..\.\+': '',
             '/\..*': '', '.*:': ''}
fnnr.df_all = str_col_replace(fnnr.df_all, 'cena', dict_cena)
fnnr.df_all['cena'] = fnnr.df_all['cena'].astype(float)

# datum
fnnr.df_all['datum'] = fnnr.df_all['datum'].str.strip()
fnnr.df_all['datum'] = fnnr.df_all['datum'].apply(get_dates)

# create final table
fnnitra_df_search = pd.DataFrame(fnnr.df_all[fnnr.final_table_cols])


# save tables
fnnr.save_tables(table=fnnitra_df_search)
db.insert_table(table_name='priame_objednavky', df=fnnr.df_all, if_exists='append', index=False)


#  FNTN - first load

fntn = PriameObjednavkyMail('fntn')

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(chromedriver_path2, options=options)
driver.get(dict_all['fakultna nemocnica trencin']['objednavky_link'])

table_html = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH,
                                                                               "//div[contains(@id, 'content')]//table"))).get_attribute(
    "outerHTML")
result_df = pd.read_html(table_html)[0]

while True:
    try:
        WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH,
                                                                   "//div[contains(@class, 'next')]//a[contains(text(), 'Nasled')]"))).click()
        sleep(4)
        next_table_lst = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//div[contains(@id, 'content')]//table"))).get_attribute(
            "outerHTML")
        next_table = pd.read_html(next_table_lst)[0]
        result_df = pd.concat([result_df, next_table], ignore_index=True)
        sleep(3)
    except TimeoutException as ex:
        print('Data were retrieved')
        break
    except:
        driver.quit()
        print('Retrieving data failed')
        break

result_df = func.load_df(os.path.join(fntn.hosp_path + '21_03_2023_11_07_08fntn_web.pkl'), path=os.getcwd())

# data cleaning
result_df = func.clean_str_cols(result_df)
result_df = clean_str_col_names(result_df)

result_df['file'] = 'web'
result_df['insert_date'] = datetime.now()

result_df['ico dodavatela'] = result_df['ico dodavatela'].str.replace('\..*', '', regex=True)
result_df['cena'] = result_df['cena s dph (EUR)'].astype(float)
result_df['datum'] = result_df['datum vyhotovenia'].apply(get_dates)
result_df['cena_s_dph'] = 'ano'
result_df['objednavka_predmet'] = result_df['popis']

fntn.df_all = result_df
fntn.dodavatel_list = ['nazov dodavatela', 'ico dodavatela']
fntn.popis_list = ['popis', 'cislo objednavky', 'cislo zmluvy', 'schvalil', 'cena_s_dph']

fntn.create_columns_w_dict(key='fakultna nemocnica trencin')

fntn_search = pd.DataFrame(fntn.df_all[fntn.final_table_cols])
fntn.save_tables(table=fntn_search)

df = fntn.df_all.drop(columns=['schvalil', 'mesto dodavatela', 'psc dodavatela', 'adresa dodavatela', 'datum vyhotovenia', 'cislo zmluvy', 'cena s dph (EUR)'])
df['objednavka_cislo'] = df['cislo objednavky']
df['dodavatel_ico'] = df['ico dodavatela']
df['dodavatel_nazov'] = df['nazov dodavatela']

df.drop(columns=['cislo objednavky', 'ico dodavatela', 'nazov dodavatela'], inplace=True)

db.insert_table(table_name='priame_objednavky', df=df, if_exists='append', index=False)



# FNSP Presov - first load ###

fnsppresov = PriameObjednavkyMail('fnsppresov')

search_result = otl.find_message(path,
                                 "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%" + fnsppresov.hosp + '.sk' + "' ")
otl.save_attachment(fnsppresov.hosp_path, search_result)

fnsppresov.load()
fnsppresov.clean_tables()
fnsppresov.data_check()
fnsppresov.create_table(stand_column_names=stand_column_names)

fnsppresov.df_all['cena'] = fnsppresov.df_all['cena'].replace('18  255,60 eur', '18255.60')
fnsppresov.df_all['cena'] = fnsppresov.df_all['cena'].astype(float)
fnsppresov.df_all['datum'] = fnsppresov.df_all['datum'].str.strip()
fnsppresov.df_all['datum'] = fnsppresov.df_all['datum'].apply(get_dates)

fnsppresov.create_columns_w_dict(key='fakultna nemocnica s poliklinikou j a reimana presov')

fnsppresov_search = pd.DataFrame(fnsppresov.df_all[fnsppresov.final_table_cols])

# save tables
db.insert_table(table_name='priame_objednavky', df=fnsppresov.df_all, if_exists='append', index=False)
fnsppresov.save_tables(table=fnsppresov_search)

# FNTT - first load ###

fntt = PriameObjednavkyMail('fntt')

search_result = otl.find_message(path,
                                 "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%" + fntt.hosp + '.sk' + "' ")
otl.save_attachment(fntt.hosp_path, search_result)

fntt.load()
fntt.clean_tables()
fntt.data_check()
fntt.create_table(stand_column_names=stand_column_names)

dict_cena = {"[a-z]|'|\s|[\(\)]+": "", ',-': '', ',$': '', ',+': '.', '/.*': ''}
fntt.df_all = str_col_replace(fntt.df_all, 'cena', dict_cena)
fntt.df_all['cena'] = fntt.df_all['cena'].astype(float)

fntt.df_all['datum'] = fntt.df_all['datum'].str.strip()
fntt.df_all['datum'] = fntt.df_all['datum'].apply(get_dates)

fntt.create_columns_w_dict(key='fakultna nemocnica trnava')
fntt_search = pd.DataFrame(fntt.df_all[fntt.final_table_cols])

fntt.save_tables(table=fntt_search)
db.insert_table(table_name='priame_objednavky', df=fntt.df_all, if_exists='append', index=False)

# DFN Kosice - first load ###

dfnkosice = PriameObjednavkyMail('dfnkosice')

search_result = otl.find_message(path,
                                 "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%" + dfnkosice.hosp + '.sk' + "' ")
otl.save_attachment(dfnkosice.hosp_path, search_result)

dfnkosice.load()
dfnkosice.clean_tables()
dfnkosice.data_check()
dfnkosice.create_table(stand_column_names=stand_column_names)

dict_cena = {"[a-z]|'|\s|[\(\)]+": "", ',-': '', ',$': '', ',+': '.', '/.*': ''}
dfnkosice.df_all = str_col_replace(dfnkosice.df_all, 'cena', dict_cena)
dfnkosice.df_all['cena'] = dfnkosice.df_all['cena'].astype(float)

dfnkosice.df_all['datum'] = dfnkosice.df_all['datum'].str.strip()
dfnkosice.df_all['datum'] = dfnkosice.df_all['datum'].apply(get_dates)

# load and append tables from docx files
for file_name in os.listdir(dfnkosice.hosp_path):
    if file_name.split(sep='.')[-1] == 'docx':
        document = Document(os.path.join(dfnkosice.hosp_path, file_name))
        table = document.tables[0]
        data = [[cell.text for cell in row.cells] for row in table.rows]
        df = pd.DataFrame(data)

        # data cleaning
        df.columns = df.iloc[0]
        df.drop(index=df.index[0], axis=0, inplace=True)
        df.drop('P.č.', axis=1, inplace=True)
        df.columns = ['kategoria', 'cena', 'objednavka_predmet', 'dodavatel_nazov']

        df = func.clean_str_cols(df)
        df['cena'] = df['cena'].str.replace('[a-z]|\s', '', regex=True).replace('\.', '', regex=True).replace(',', '.',
                                                                                                              regex=True)
        df['cena'] = df['cena'].astype(float)
        df['file'] = file_name.split(sep='.')[0]
        df['insert_date'] = datetime.now()
        dfnkosice.df_all = pd.concat([dfnkosice.df_all, df], ignore_index=True)

dfnkosice.create_columns_w_dict(key='detska fakultna nemocnica kosice')
dfnkosice_search = pd.DataFrame(dfnkosice.df_all[dfnkosice.final_table_cols])

dfnkosice.save_tables(table=dfnkosice_search)
db.insert_table(table_name='priame_objednavky', df=dfnkosice.df_all, if_exists='append', index=False)

# UNB - first load ###

unb = PriameObjednavkyMail('unb')

search_result = otl.find_message(path,
                                 "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%" + unb.hosp + '.sk' + "' ")
otl.save_attachment(unb.hosp_path, search_result)

unb.load()
unb.clean_tables()
unb.data_check()
unb.create_table(stand_column_names=stand_column_names)

dict_cena = {"[a-z]|'|\s|[\(\)]+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '_':'0'}
unb.df_all = str_col_replace(unb.df_all, 'cena', dict_cena)
unb.df_all['cena'] = np.where(unb.df_all['cena'].str.match(r'\d*\.\d*\.\d*'),
                              unb.df_all['cena'].str.replace('\.', '', 1, regex=True), unb.df_all['cena'])
unb.df_all['cena'] = unb.df_all['cena'].astype(float)

unb.df_all['datum'] = unb.df_all['datum'].apply(get_dates)
unb.df_all.drop(unb.df_all[unb.df_all['kategoria'].isin(['0', '_'])].index, axis=0, inplace=True)

unb.create_columns_w_dict(key='univerzitna nemocnica bratislava')
unb_search = pd.DataFrame(unb.df_all[unb.final_table_cols])

unb.save_tables(table=unb_search)
db.insert_table(table_name='priame_objednavky', df=unb.df_all, if_exists='append', index=False)

# UNLP KE - first load ###

unlp = PriameObjednavkyMail('unlp')

search_result = otl.find_message(path,
                                 "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%" + unlp.hosp + '.sk' + "' ")
otl.save_attachment(unlp.hosp_path, search_result)

unlp.load()
unlp.clean_tables()
unlp.data_check()
unlp.create_table(stand_column_names=stand_column_names)

dict_cena = {"[a-z]|'|\s|[\(\)]|\++": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0'}
unlp.df_all = str_col_replace(unlp.df_all, 'cena', dict_cena)
unlp.df_all['cena'] = np.where(unlp.df_all['cena'].str.match(r'\d*\.\d*\.\d*'),
                               unlp.df_all['cena'].str.replace('\.', '', 1, regex=True), unlp.df_all['cena'])
unlp.df_all['cena'] = unlp.df_all['cena'].astype(float)

unlp.df_all['datum'] = unlp.df_all['datum'].apply(get_dates)
unlp.df_all.drop(unlp.df_all[unlp.df_all['objednavka_predmet']=='spolu'].index, axis=0, inplace=True)

unlp.create_columns_w_dict(key='univerzitna nemocnica l pasteura kosice')
unlp_search = pd.DataFrame(unlp.df_all[unlp.final_table_cols])

unlp.save_tables(table=unlp_search)
db.insert_table(table_name='priame_objednavky', df=unlp.df_all, if_exists='append', index=False)

# UNM - first load ###

unm = PriameObjednavkyMail('unm')

search_result = otl.find_message(path,
                                 "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%" + unm.hosp + '.sk' + "' ")
otl.save_attachment(unm.hosp_path, search_result)

unm.load()
unm.clean_tables()
unm.data_check()
unm.create_table(stand_column_names=stand_column_names)

dict_cena = {"[a-z]|'|\s|[\(\)]|\++": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0'}
unm.df_all = str_col_replace(unm.df_all, 'cena', dict_cena)
unm.df_all['cena'] = np.where(unm.df_all['cena'].str.match(r'\d*\.\d*\.\d*'),
                              unm.df_all['cena'].str.replace('.', '', 1), unm.df_all['cena'])
unm.df_all['cena'] = unm.df_all['cena'].astype(float)

unm.df_all['datum'] = unm.df_all['datum'].apply(get_dates)

unm.create_columns_w_dict(key='univerzitna nemocnica martin')
unm_search = pd.DataFrame(unm.df_all[unm.final_table_cols])

unm.save_tables(table=unm_search)
db.insert_table(table_name='priame_objednavky', df=unm.df_all, if_exists='append', index=False)

# DFNBB - first load ###

dfnbb = PriameObjednavkyMail('dfnbb')

search_result = otl.find_message(path,
                                 "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%" + dfnbb.hosp + '.sk' + "' ")
otl.save_attachment(dfnbb.hosp_path, search_result)

dfnbb.load()
dfnbb.clean_tables()
dfnbb.data_check()
dfnbb.create_table(stand_column_names=stand_column_names)

dict_cena = {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0'}
dfnbb.df_all = str_col_replace(dfnbb.df_all, 'cena', dict_cena)
dfnbb.df_all['cena'] = dfnbb.df_all['cena'].astype(float)

dfnbb.df_all['datum'] = dfnbb.df_all['datum'].apply(get_dates)

dfnbb.create_columns_w_dict(key='detska fakultna nemocnica s poliklinikou banska bystrica')
dfnbb_search = pd.DataFrame(dfnbb.df_all[dfnbb.final_table_cols])

dfnbb.save_tables(table=dfnbb_search)
db.insert_table(table_name='priame_objednavky', df=dfnbb.df_all, if_exists='append', index=False)

# NOU - first load ###

nou = PriameObjednavkyMail('nou')

search_result = otl.find_message(path,
                                 "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%" + nou.hosp + '.sk' + "' ")
otl.save_attachment(nou.hosp_path, search_result)
nou.load()
nou.clean_tables()
nou.data_check()

nou.create_table(stand_column_names=stand_column_names)

# load pdf files
dict_pdf_files = {}
cols = ['kategoria', 'objednavka_predmet', 'cena', 'datum', 'zdroj_financovania', 'poznamka']

for file_name in os.listdir(nou.hosp_path):
    if file_name.split(sep='.')[-1] == 'pdf':
        list_of_pages = camelot.read_pdf(os.path.join(nou.hosp_path, file_name), pages='all')
        df_conc_pages = pd.DataFrame(columns=cols)
        list_of_pages[0].df.drop(list_of_pages[0].df.index[0], inplace=True, axis=0)

        for i in range(len(list_of_pages)):
            list_of_pages[i].df.columns = cols
            list_of_pages[i].df['file'] = file_name
            list_of_pages[i].df['insert_date'] = datetime.now()
            df_conc_pages = pd.concat([df_conc_pages, list_of_pages[i].df], ignore_index=True)

        dict_pdf_files[file_name] = df_conc_pages

# create df with all pdf files and clean it
df_all_pdf = pd.concat([table for table in dict_pdf_files.values()], ignore_index=True)
df_all_pdf = func.clean_str_cols(df_all_pdf)

nou.df_all = pd.concat([nou.df_all, df_all_pdf], ignore_index=True)

dict_cena = {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0'}
nou.df_all = str_col_replace(nou.df_all, 'cena', dict_cena)
nou.df_all['cena'] = nou.df_all['cena'].astype(float)

nou.df_all['datum'] = nou.df_all['datum'].apply(get_dates)

nou.create_columns_w_dict(key='narodny onkologicky ustav')
nou_search = pd.DataFrame(nou.df_all[nou.final_table_cols])

nou.save_tables(table=nou_search)
db.insert_table(table_name='priame_objednavky', df=nou.df_all, if_exists='append', index=False)

# NOUSK - first load ###

nou = PriameObjednavkyMail('nou')

search_result = otl.find_message(path,
                                 "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%" + nou.hosp + 'sk.sk' + "' ")
otl.save_attachment(nou.hosp_path, search_result)

nou.load()
nou.clean_tables()
nou.data_check()
nou.create_table(stand_column_names=stand_column_names)

dict_cena = {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0'}
nou.df_all = str_col_replace(nou.df_all, 'cena', dict_cena)
nou.df_all['cena'] = np.where(nou.df_all['cena'].str.match(r'\d*\.\d*\.\d*'),
                              nou.df_all['cena'].str.replace('\.', '', 1, regex=True), nou.df_all['cena'])
nou.df_all['cena'] = nou.df_all['cena'].astype(float)
nou.df_all['datum'] = nou.df_all['datum'].apply(get_dates)

nou.create_columns_w_dict(key='narodny onkologicky ustav')
nou_search = pd.DataFrame(nou.df_all[nou.final_table_cols])

nou.save_tables(table=nou_search)
db.insert_table(table_name='priame_objednavky', df=nou.df_all, if_exists='append', index=False)


# NUSCH - first load ###

nusch = PriameObjednavkyMail('nusch')

search_result = otl.find_message(path,
                                 "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%" + nusch.hosp + '.sk' + "' ")
otl.save_attachment(nusch.hosp_path, search_result)

nusch.load()
nusch.clean_tables()
nusch.data_check()
nusch.create_table(stand_column_names=stand_column_names)

dict_cena = {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0',
             '=.*': '', '.+\d{2}:\d{2}:\d{2}$': '0'}
nusch.df_all = str_col_replace(nusch.df_all, 'cena', dict_cena)
nusch.df_all['cena'] = nusch.df_all['cena'].astype(float)

nusch.df_all['datum'] = nusch.df_all['datum'].apply(get_dates)
nusch.create_columns_w_dict(key='narodny ustav srdcovych a cievnych chorob as')
nusch_search = pd.DataFrame(nusch.df_all[nusch.final_table_cols])

db.insert_table(table_name='priame_objednavky', df=nusch.df_all, if_exists='append', index=False)
nusch.save_tables(table=nusch_search)

# VOU - first load ###
vou = PriameObjednavkyMail('vou')

search_result = otl.find_message(path,
                                    "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%" + vou.hosp + '.sk' + "' ")
otl.save_attachment(vou.hosp_path, search_result)

vou.load()
vou.clean_tables()
vou.data_check()

vou.create_table(stand_column_names=stand_column_names)

dict_cena = {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0'}
vou.df_all = str_col_replace(vou.df_all, 'cena', dict_cena)
vou.df_all['cena'] = vou.df_all['cena'].astype(float)

vou.df_all['datum'] = vou.df_all['datum'].str.strip()
vou.df_all['datum'] = vou.df_all['datum'].apply(get_dates)

vou.create_columns_w_dict(key='vychodoslovensky onkologicky ustav as')
vou_search = pd.DataFrame(vou.df_all[vou.final_table_cols])

db.insert_table(table_name='priame_objednavky', df=vou.df_all, if_exists='append', index=False)
vou.save_tables(table=vou_search)

# NUDCH - first load ###

nudch = PriameObjednavkyMail('nudch')

search_result = otl.find_message(path, "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%" + nudch.hosp + '.eu' + "' ")
otl.save_attachment(nudch.hosp_path, search_result)

nudch.load()
nudch.clean_tables()
nudch.data_check()
nudch.create_table(stand_column_names=stand_column_names)

dict_cena = {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0'}
nudch.df_all = str_col_replace(nudch.df_all, 'cena', dict_cena)
nudch.df_all['cena'] = nudch.df_all['cena'].astype(float)

nudch.df_all['datum'] = nudch.df_all['datum'].str.strip()
nudch.df_all['datum'] = nudch.df_all['datum'].apply(get_dates)

nudch.create_columns_w_dict(key='narodny ustav detskych chorob')
nudch_search = pd.DataFrame(nudch.df_all[nudch.final_table_cols])

db.insert_table(table_name='priame_objednavky', df=nudch.df_all, if_exists='append', index=False)
nudch.save_tables(table=nudch_search)

# SUSCCH - first load ###

suscch = PriameObjednavkyMail('suscch')

search_result = otl.find_message(path, "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%" + suscch.hosp + '.eu' + "' ")
otl.save_attachment(suscch.hosp_path, search_result)

suscch.load()
suscch.clean_tables()
suscch.data_check()
suscch.create_table(stand_column_names=stand_column_names)

dict_cena = {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0'}
suscch.df_all = str_col_replace(suscch.df_all, 'cena', dict_cena)
suscch.df_all['cena'] = suscch.df_all['cena'].astype(float)

suscch.df_all['datum'] = suscch.df_all['datum'].str.strip()
suscch.df_all['datum'] = suscch.df_all['datum'].apply(get_dates)

suscch.create_columns_w_dict(key='stredoslovensky ustav srdcovych a cievnych chorob as')
suscch_search = pd.DataFrame(suscch.df_all[suscch.final_table_cols])

db.insert_table(table_name='priame_objednavky', df=suscch.df_all, if_exists='append', index=False)
suscch.save_tables(table=suscch_search)

# INMM - first load ###

inmm = PriameObjednavkyMail('inmm')

search_result = otl.find_message(path, "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%" + inmm.hosp + '.sk' + "' ")
otl.save_attachment(inmm.hosp_path, search_result)


inmm.load()
inmm.clean_tables()
inmm.data_check()
inmm.create_table(stand_column_names=stand_column_names)

dict_cena = {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0'}
inmm.df_all = str_col_replace(inmm.df_all, 'cena', dict_cena)
inmm.df_all['cena'] = np.where(inmm.df_all['cena'].str.match(r'\d*\.\d*\.\d*'),
                              inmm.df_all['cena'].str.replace('\.', '', 1, regex=True), inmm.df_all['cena'])
inmm.df_all['cena'] = inmm.df_all['cena'].astype(float)

inmm.df_all['datum'] = inmm.df_all['datum'].str.strip()
inmm.df_all['datum'] = inmm.df_all['datum'].apply(get_dates)

inmm.create_columns_w_dict(key='institut nuklearnej a molekularnej mediciny kosice')
inmm_search = pd.DataFrame(inmm.df_all[inmm.final_table_cols])

db.insert_table(table_name='priame_objednavky', df=inmm.df_all, if_exists='append', index=False)
inmm.save_tables(table=inmm_search)


# NURCH - first load ###

nurch = PriameObjednavkyMail('nurch')

search_result = otl.find_message(path, "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%" + nurch.hosp + '.sk' + "' ")
otl.save_attachment(nurch.hosp_path, search_result)

nurch.load()
nurch.clean_tables()
nurch.data_check()
nurch.create_table(stand_column_names=stand_column_names)

dict_cena = {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0'}
nurch.df_all = str_col_replace(nurch.df_all, 'cena', dict_cena)
nurch.df_all['cena'] = np.where(nurch.df_all['cena'].str.match(r'\d*\.\d*\.\d*'),
                              nurch.df_all['cena'].str.replace('\.', '', 1, regex=True), nurch.df_all['cena'])
nurch.df_all['cena'] = nurch.df_all['cena'].astype(float)

nurch.df_all['datum'] = nurch.df_all['datum'].apply(get_dates)
nurch.create_columns_w_dict(key='narodny ustav reumatickych chorob piestany')
nurch_search = pd.DataFrame(nurch.df_all[nurch.final_table_cols])

db.insert_table(table_name='priame_objednavky', df=nurch.df_all, if_exists='append', index=False)
nurch.save_tables(table=nurch_search)


# Nemocnica PP - first load ###

nemocnicapp = PriameObjednavkyMail('nemocnicapp')

search_result = otl.find_message(path, "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%" + nemocnicapp.hosp + '.sk' + "' ")
otl.save_attachment(nemocnicapp.hosp_path, search_result)

nemocnicapp.load()
nemocnicapp.clean_tables()
nemocnicapp.data_check()
nemocnicapp.create_table(stand_column_names=stand_column_names)

dict_pdf_files = {}
cols = ['kategoria', 'objednavka_predmet', 'cena', 'datum', 'zdroj_financovania']
for file_name in os.listdir(nemocnicapp.hosp_path):
    if file_name.split(sep='.')[-1] == 'pdf':
        list_of_pages = camelot.read_pdf(os.path.join(nemocnicapp.hosp_path, file_name), pages='1')
        df_conc_pages = pd.DataFrame(columns=cols)
        list_of_pages[0].df.drop(list_of_pages[0].df.index[0], inplace=True, axis=0)

        for i in range(len(list_of_pages)):
            list_of_pages[i].df.columns = cols
            list_of_pages[i].df['file'] = file_name
            list_of_pages[i].df['insert_date'] = datetime.now()
            df_conc_pages = pd.concat([df_conc_pages, list_of_pages[i].df], ignore_index=True)

        dict_pdf_files[file_name] = df_conc_pages

# create df with all pdf files and clean it
df_all_pdf = pd.concat([table for table in dict_pdf_files.values()], ignore_index=True)
df_all_pdf = func.clean_str_cols(df_all_pdf)
df_all_pdf.drop_duplicates(inplace=True)

nemocnicapp.df_all = pd.concat([nemocnicapp.df_all, df_all_pdf], ignore_index=True)

dict_cena = {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0', '\.+':'.'}
nemocnicapp.df_all = str_col_replace(nemocnicapp.df_all, 'cena', dict_cena)
nemocnicapp.df_all['cena'] = nemocnicapp.df_all['cena'].astype(float)

nemocnicapp.df_all['datum'] = nemocnicapp.df_all['datum'].str.replace('2055', '2022')
nemocnicapp.df_all['datum'] = nemocnicapp.df_all['datum'].apply(get_dates)
nemocnicapp.create_columns_w_dict(key='nemocnica poprad as')

nemocnicapp_search = pd.DataFrame(nemocnicapp.df_all[nemocnicapp.final_table_cols])

db.insert_table(table_name='priame_objednavky', df=nemocnicapp.df_all, if_exists='append', index=False)
nemocnicapp.save_tables(table=nemocnicapp_search)

df_orig = pd.concat([df_orig, nemocnicapp_search], ignore_index=True)

