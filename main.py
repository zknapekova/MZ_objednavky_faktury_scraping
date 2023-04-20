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
db_cloud = ObjednavkyDB(objednavky_db_connection_cloud)

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

result_df = result_df.assign(file='web', insert_date=datetime.now(), cena_s_dph='ano')
result_df.rename(columns={'cislo objednavky': 'objednavka_cislo', 'nazov dodavatela': 'dodavatel_nazov',
                          'popis':'objednavka_predmet'}, inplace=True)

result_df['dodavatel_ico'] = result_df['ico dodavatela'].str.replace('\..*', '', regex=True)
result_df['cena'] = result_df['cena s dph (EUR)'].astype(float)
result_df['datum'] = result_df['datum vyhotovenia'].apply(get_dates)

fntn.df_all = result_df
fntn.popis_list = ['objednavka_predmet', 'objednavka_cislo', 'cislo zmluvy', 'schvalil', 'cena_s_dph']

fntn.create_columns_w_dict(key='fakultna nemocnica trencin')
fntn.df_all.drop_duplicates(inplace=True)
fntn_search = pd.DataFrame(fntn.df_all[fntn.final_table_cols])
fntn.save_tables(table=fntn_search)

fntn.df_all.drop(columns=['ico dodavatela', 'schvalil', 'mesto dodavatela', 'psc dodavatela', 'adresa dodavatela',
                          'datum vyhotovenia', 'cislo zmluvy', 'cena s dph (EUR)'], inplace=True)
db.insert_table(table_name='priame_objednavky', df=fntn.df_all, if_exists='append', index=False)

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

# vusch - first load ###
vusch = PriameObjednavkyMail('vusch')

search_result = otl.find_message(path, "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%" + vusch.hosp + '.sk' + "' ")
otl.save_attachment(vusch.hosp_path, search_result)

vusch.load()
vusch.clean_tables()
vusch.data_check()
vusch.create_table(stand_column_names=stand_column_names)

dict_cena = {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0'}
vusch.df_all = str_col_replace(vusch.df_all, 'cena', dict_cena)
vusch.df_all['cena'] = np.where(vusch.df_all['cena'].str.match(r'\d*\.\d*\.\d*'),
                              vusch.df_all['cena'].str.replace('\.', '', 1, regex=True), vusch.df_all['cena'])
vusch.df_all['cena'] = vusch.df_all['cena'].astype(float)

vusch.df_all['datum'] = vusch.df_all['datum'].apply(get_dates)
vusch.create_columns_w_dict(key='vychodoslovensky ustav srdcovych a cievnych chorob as')
vusch_search = pd.DataFrame(vusch.df_all[vusch.final_table_cols])

db.insert_table(table_name='priame_objednavky', df=vusch.df_all, if_exists='append', index=False)
vusch.save_tables(table=vusch_search)

# DONSP - first load - no mails found ###

donsp = PriameObjednavkyMail('donsp')

def donsp_table_clean(df, df_rename_dict, set_column_names=False):
    if set_column_names:
        df.drop(index=[df.shape[0] - 1], inplace=True)
        df.dropna(axis=1, how='all', inplace=True)
        df.columns = df.iloc[0]
        df.drop(df.index[0], inplace=True)

    df.rename(columns=df_rename_dict, inplace=True)

    df = func.clean_str_cols(df)
    dict_cena = {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '',
                 '': '0',
                 '\.+': '.'}
    df = str_col_replace(df, 'cena', dict_cena)
    df['cena'] = df['cena'].astype(float)
    df['datum'] = df['datum'].str.replace('29.2.', '27.2.').str.replace('31.04.', '30.04.').str.replace('31.06.', '30.06.')
    df['datum'] = df['datum'].apply(get_dates)
    return df

def donsp_data_download(webapge:str):

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(chromedriver_path2, options=options)
    driver.get(webapge)

    table_html = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH,
                                "//div[contains(@class, 'responsive-table')]//table"))).get_attribute("outerHTML")
    result_df = pd.read_html(table_html)[0]
    result_df = donsp_table_clean(result_df, set_column_names=True, df_rename_dict={'Číslo': 'objednavka_cislo',
                    'Dodávateľ': 'dodavatel_nazov', 'IČO': 'dodavatel_ico', 'Suma celkom': 'cena', 'Popis tovaru': 'objednavka_predmet', 'DÁTUM': 'datum'})

    while True:
        try:
            sleep(4)
            WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH,
                "//div[contains(@class, 'responsive-table')]//table[contains(@class, 'container')]//tr[contains(@class, 'foot')]//a[contains(text(), '»')]"))).click()

            table_html = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH,
                                        "//div[contains(@class, 'responsive-table')]//table"))).get_attribute("outerHTML")
            table_next = pd.read_html(table_html)[0]
            table_next = donsp_table_clean(table_next)
            result_df = pd.concat([result_df, table_next], ignore_index=True)
            sleep(3)
        except TimeoutException:
            print('Data were retrieved')
            break
        except Exception:
            print(traceback.format_exc())
            break
    driver.quit()

    return result_df

# 2021-2023
donsp.df_all = donsp_data_download(donsp_webpages['objednavky_2021_2023'])
donsp.df_all = donsp.df_all.assign(cena_s_dph='nie', file='web', insert_date=datetime.now())

donsp.popis_list = ['objednavka_predmet', 'objednavka_cislo', 'cena_s_dph']
donsp.create_columns_w_dict(key='dolnooravska nemocnica s poliklinikou mudr l nadasi jegeho dolny kubin')

donsp_search = pd.DataFrame(donsp.df_all[donsp.final_table_cols])
db.insert_table(table_name='priame_objednavky', df=donsp.df_all.drop(columns=['Meno schvaľujúceho']), if_exists='append', index=False)

# 2020
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(chromedriver_path2, options=options)
driver.get(donsp_webpages['objednavky_2020'])

n_records_dropdown = driver.find_element(By.XPATH, "//div[contains(@id, 'tablepress-43_length')]//select[contains(@name, 'tablepress-43_length')]")
select = Select(n_records_dropdown)
select.select_by_value('-1')

table_html = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH,
                                "//table[contains(@id, 'tablepress-43')]"))).get_attribute("outerHTML")
driver.quit()
result_df = pd.read_html(table_html)[0]
result_df.rename(
    columns={'Číslo': 'objednavka_cislo', 'Dodávateľ': 'dodavatel_nazov', 'IČO': 'dodavatel_ico',
             'Suma celkom': 'cena', 'Popis tovaru': 'objednavka_predmet', 'DÁTUM': 'datum'},
    inplace=True)

result_df = func.clean_str_cols(result_df)
dict_cena = {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '',
             '': '0', '\.+': '.'}
result_df = str_col_replace(result_df, 'cena', dict_cena)
result_df['cena'] = result_df['cena'].astype(float)
result_df['datum'] = result_df['datum'].str.replace('069', '09').str.replace('101', '10')
result_df['datum'] = result_df['datum'].apply(get_dates)


donsp = PriameObjednavkyMail('donsp')
donsp.df_all = result_df.assign(cena_s_dph='nie', file='web', insert_date=datetime.now())
donsp.popis_list = ['objednavka_predmet', 'objednavka_cislo', 'cena_s_dph']
donsp.dodavatel_list = ['dodavatel_nazov']

donsp.create_columns_w_dict(key='dolnooravska nemocnica s poliklinikou mudr l nadasi jegeho dolny kubin_2020')
donsp.df_all.drop_duplicates(inplace=True)

donsp_search = pd.DataFrame(donsp.df_all[donsp.final_table_cols])

db.insert_table(table_name='priame_objednavky', df=donsp.df_all.drop(columns=['Číslo zmluvy', 'Schválil']))
db_cloud.insert_table(table_name='priame_objednavky', df=donsp.df_all.drop(columns=['Číslo zmluvy', 'Schválil']))


# 2019
driver = webdriver.Chrome(chromedriver_path2, options=options)
driver.get(donsp_webpages['objednavky_2019'])

table_html = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH,
                                "//table[contains(@id, 'tablepress-37')]"))).get_attribute("outerHTML")
driver.quit()
result_df = pd.read_html(table_html)[0]

donsp = PriameObjednavkyMail('donsp')
donsp.df_all = donsp_table_clean(result_df, {'Číslo': 'objednavka_cislo', 'Dodávateľ': 'dodavatel_nazov',
             'Suma celkom': 'cena', 'Popis tovaru': 'objednavka_predmet', 'DÁTUM': 'datum'})

donsp.df_all = donsp.df_all.assign(cena_s_dph='nie', file='web', insert_date=datetime.now())
donsp.popis_list = ['objednavka_predmet', 'objednavka_cislo', 'cena_s_dph']
donsp.dodavatel_list = ['dodavatel_nazov']
donsp.create_columns_w_dict(key='dolnooravska nemocnica s poliklinikou mudr l nadasi jegeho dolny kubin_2019')
donsp.df_all.drop_duplicates(inplace=True)

donsp_search = pd.DataFrame(donsp.df_all[donsp.final_table_cols])

db.insert_table(table_name='priame_objednavky', df=donsp.df_all.drop(columns=['Číslo zmluvy', 'Schválil']))
db_cloud.insert_table(table_name='priame_objednavky', df=donsp.df_all.drop(columns=['Číslo zmluvy', 'Schválil']))

# 2018

driver = webdriver.Chrome(chromedriver_path2, options=options)
driver.get(donsp_webpages['objednavky_2018'])

table_html = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH,
                                "//table[contains(@id, 'tablepress-34')]"))).get_attribute("outerHTML")
driver.quit()
result_df = pd.read_html(table_html)[0]

donsp = PriameObjednavkyMail('donsp')

donsp.df_all = donsp_table_clean(result_df, {'Číslo': 'objednavka_cislo', 'Dodávateľ': 'dodavatel_nazov',
             'Suma celkom': 'cena', 'Popis tovaru': 'objednavka_predmet', 'DÁTUM': 'datum'})

donsp.df_all = donsp.df_all.assign(cena_s_dph='nie', file='web', insert_date=datetime.now())
donsp.popis_list = ['objednavka_predmet', 'objednavka_cislo', 'cena_s_dph']
donsp.dodavatel_list = ['dodavatel_nazov']
donsp.create_columns_w_dict(key='dolnooravska nemocnica s poliklinikou mudr l nadasi jegeho dolny kubin_2018')
donsp.df_all.drop_duplicates(inplace=True)

donsp_search = pd.DataFrame(donsp.df_all[donsp.final_table_cols])


db.insert_table(table_name='priame_objednavky', df=donsp.df_all.drop(columns=['Číslo zmluvy', 'Schválil']))
db_cloud.insert_table(table_name='priame_objednavky', df=donsp.df_all.drop(columns=['Číslo zmluvy', 'Schválil']))

# 2017

driver = webdriver.Chrome(chromedriver_path2, options=options)
driver.get(donsp_webpages['objednavky_2017'])

table_html = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH,
                                "//table[contains(@id, 'tablepress-25')]"))).get_attribute("outerHTML")
driver.quit()
result_df = pd.read_html(table_html)[0]
result_df.columns=['objednavka_cislo', 'objednavka_predmet', 'cena', 'cislo_zmluvy', 'datum', 'dodavatel_nazov', 'schvalil']

donsp = PriameObjednavkyMail('donsp')

donsp.df_all = donsp_table_clean(result_df, {})

donsp.df_all = donsp.df_all.assign(cena_s_dph='nie', file='web', insert_date=datetime.now())
donsp.popis_list = ['objednavka_predmet', 'objednavka_cislo', 'cena_s_dph']
donsp.dodavatel_list = ['dodavatel_nazov']
donsp.create_columns_w_dict(key='dolnooravska nemocnica s poliklinikou mudr l nadasi jegeho dolny kubin_2017')
donsp.df_all.drop_duplicates(inplace=True)

donsp_search = pd.DataFrame(donsp.df_all[donsp.final_table_cols])

db.insert_table(table_name='priame_objednavky', df=donsp.df_all.drop(columns=['cislo_zmluvy', 'schvalil']))
db_cloud.insert_table(table_name='priame_objednavky', df=donsp.df_all.drop(columns=['cislo_zmluvy', 'schvalil']))



# NSP Trstena - no mails found, pdf files available at https://www.nsptrstena.sk/sk/objednavky

# 2017-2022
nsptrstena = PriameObjednavkyMail('nsptrstena')

df_all = pd.DataFrame()

for file_name in os.listdir(nsptrstena.hosp_path):
    list_of_dfs = camelot.read_pdf(os.path.join(nsptrstena.hosp_path, file_name), pages='all', flavor='stream',
                                   row_tol=20)
    for i in range(len(list_of_dfs)):
        df = list_of_dfs[i].df
        mask = df.applymap(lambda x: bool(re.search('(Str:.*)|(Vystavené objednávky za obdobie:)', str(x)))).any(axis=1)
        df.drop(df[mask].index, inplace=True)
        df.reset_index(inplace=True, drop=True)
        df.columns = df.iloc[0]
        df.drop([0], inplace=True)
        if '' in df.columns.values: df.drop(columns=[''], inplace=True)
        df['file'] = file_name
        df_all = pd.concat([df_all, df], ignore_index=True)

df_all = func.clean_str_cols(df_all)
df_all = clean_str_col_names(df_all)
df_all['datum'] = df_all['datumobjednania'].apply(get_dates)

df_all.loc[pd.isna(df_all['predbezna'])==False, 'predbeznacena bez dph'] = df_all['predbezna']

df_all['predbeznapredmet objednaniacena bez dph'] = df_all['predbeznapredmet objednaniacena bez dph'].str.replace('\seur\s', ' ', regex=True)

df_all['cena_extr'] = df_all['predbeznapredmet objednaniacena bez dph'].str.extract(r'(\d{1,}\.\d*)')
df_all['predbeznapredmet objednaniacena bez dph'] = df_all['predbeznapredmet objednaniacena bez dph'].str.replace('\d{1,}\.\d*', ' ', regex=True)


df_all.loc[(pd.isna(df_all['predbeznapredmet objednaniacena bez dph']) == False) & (pd.isna(df_all['predbeznacena bez dph']) == True),
                'predbeznacena bez dph'] = df_all['cena_extr']

df_all.loc[(pd.isna(df_all['predbeznapredmet objednaniacena bez dph']) == False) & (pd.isna(df_all['predmet objednania']) == True),
                'predmet objednania'] = df_all['predbeznapredmet objednaniacena bez dph']

df_all.loc[df_all['predbeznacena bez dph'] == '', 'predbeznacena bez dph'] = '0'
df_all['predbeznacena bez dph'] = df_all['predbeznacena bez dph'].str.replace('[(a-z)|\s]', '', regex=True)
df_all['cena']=df_all['predbeznacena bez dph'].astype(float)

nsptrstena.df_all = df_all.assign(cena_s_dph='nie', insert_date=datetime.now())
nsptrstena.df_all.rename(columns = {'predmet objednania':'objednavka_predmet', 'objednavka':'objednavka_cislo'}, inplace=True)

nsptrstena.popis_list = ['objednavka_predmet', 'objednavka_cislo', 'zmluva', 'cena_s_dph']
nsptrstena.dodavatel_list = ['dodavatel']
nsptrstena.create_columns_w_dict(key='hornooravska nemocnica s poliklinikou trstena')
nsptrstena.df_all.drop_duplicates(inplace=True)

nsptrstena_search = pd.DataFrame(nsptrstena.df_all[nsptrstena.final_table_cols])
db.insert_table(table_name='priame_objednavky', df=nsptrstena.df_all.drop(columns=['predbeznacena bez dph', 'zmluva', 'datumobjednania', 'schvalil', 'predbezna', 'cena bez dph', 'predbeznapredmet objednaniacena bez dph', 'cena_extr']))
db.insert_table(table_name='priame_objednavky', df=nsptrstena.df_all.drop(columns=['predbeznacena bez dph', 'zmluva', 'datumobjednania', 'schvalil', 'predbezna', 'cena bez dph', 'predbeznapredmet objednaniacena bez dph', 'cena_extr']))



# 2023
nsptrstena = PriameObjednavkyMail('nsptrstena')

urlretrieve("https://www.nsptrstena.sk/uploads/fck/document/Objednavky/ROK%202023/Zoznam%20objednavok%202023(1).pdf",
            nsptrstena.hosp_path + current_date_time.strftime("%d-%m-%Y") + str(keysList[63]).replace(" ", "_") + '.pdf')


db.insert_table(table_name='priame_objednavky', df=nsptrstena.df_all.drop(columns=['predbeznacena bez dph', 'zmluva', 'datumobjednania', 'schvalil',  'predbeznapredmet objednaniacena bez dph', 'cena_extr']))
db_cloud.insert_table(table_name='priame_objednavky', df=nsptrstena.df_all.drop(columns=['predbeznacena bez dph', 'zmluva', 'datumobjednania', 'schvalil',  'predbeznapredmet objednaniacena bez dph', 'cena_extr']))

# Kysucka nemocnica

kysuckanemocnica = PriameObjednavkyMail('kysuckanemocnica')

def kn_download_files(year_start, path):
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(chromedriver_path2, options=options)

    year = year_start
    while year <= current_date_time.year:

        link = 'https://www.kysuckanemocnica.sk/zverejnene-dokumenty/objednavky/objednavky-'+str(year)

        driver.get(link)
        #all_tabs = driver.find_elements(By.XPATH, "//table[contains(@class, 'tabledok')]//a[contains(text(), 'Zdravotnícka technika') or contains(text(), 'Oddelenie IT')]")
        all_tabs = driver.find_elements(By.XPATH, "//table[contains(@class, 'tabledok')]//a[contains(text(), '"+str(year)+"')]")

        for element in all_tabs:
            down_link = element.get_attribute('href')
            try:
                urlretrieve(down_link, path + current_date_time.strftime("%d-%m-%Y") + str(element.get_attribute('text'))+'.pdf')
            except Exception as e:
                print(down_link, element.get_attribute('text'))
                print(traceback.format_exc())
                continue
        year += 1
    driver.quit()

kn_download_files(2017, kysuckanemocnica.hosp_path)

# kuchyna
dict_tables = {}
for file_name in os.listdir(kysuckanemocnica.hosp_path):

        print(file_name)
        try:
            if re.match('.*Kuchyňa.*', file_name) and file_name not in ('13-04-2023Kuchyňa - Február 2019.pdf', '13-04-2023Kuchyňa - Január 2019.pdf', '13-04-2023Kuchyňa - Marec 2019.pdf'):
                dict_tables[file_name]= camelot.read_pdf(os.path.join(kysuckanemocnica.hosp_path, file_name), pages='all', flavor='stream',
                                                                  row_tol=13, split_text=True)
            elif file_name in ('13-04-2023Kuchyňa - Február 2019.pdf', '13-04-2023Kuchyňa - Január 2019.pdf', '13-04-2023Kuchyňa - Marec 2019.pdf'):
                dict_tables[file_name] = camelot.read_pdf(os.path.join(kysuckanemocnica.hosp_path, file_name),
                                                          pages='all', flavor='stream',
                                                          row_tol=8, split_text=True)
        except Exception:
            print(file_name)
            print(traceback.format_exc())
            continue

import copy
dict_tables2=copy.deepcopy(dict_tables)

df_all = pd.DataFrame()

for key in dict_tables.keys():
    try:
        for i in range(len(dict_tables[key])):
            df = dict_tables[key][i].df
            print(key)
            df = func.clean_str_cols(df)

            mask = df.applymap(lambda cell: bool(
                re.search('(.*celkova.*)|(bez dph.*)|(.*ratane dph.*)|(identifikacne udaje.+)|(objedna\s*neho.*)|(meno/nazov,.+)|(osoby/ obchodne.+)|(.+miesto podnikania.+)|(meno a priezvisko.+)|(pobytu fo.+)',
                          cell))).any(axis=1)
            df = df.drop(df[mask].index).reset_index(drop=True)
            df['file'] = key

            if not df.empty:
                if (key in ('13-04-2023Kuchyňa - Február 2022.pdf') and i in (0,3,4)) or (key in ('13-04-2023Kuchyňa - Marec 2022.pdf') and i in (0,2)):
                    df.columns = ['objednavka_cislo', 'objednavka_predmet', 'cena_bez_dph', 'cena_zahr_dph',
                                  'datum', 'dodavatel', 'schvalil', 'file']
                elif key == '13-04-2023Kuchyňa - Marec 2022.pdf' and i ==5:
                    df.columns = ['objednavka_cislo', 'objednavka_predmet', 'cena_bez_dph', 'cena_zahr_dph', 'dodavatel', 'file']
                elif key == '13-04-2023Kuchyňa - Máj 2019.pdf' and i ==2:
                    df.columns = ['objednavka_cislo', 'objednavka_predmet', 'cena_zahr_dph', 'datum', 'dodavatel', 'schvalil', 'file']
                elif key == '13-04-2023Kuchyňa - November 2022.pdf' and i ==0:
                    df.columns = ['objednavka_cislo', 'objednavka_predmet', 'cena_bez_dph', 'cena_zahr_dph', 'cislo_zmluvy',
                                  'datum', 'dodavatel', 'file']
                else:
                    df.columns = ['objednavka_cislo', 'objednavka_predmet', 'cena_bez_dph', 'cena_zahr_dph', 'cislo_zmluvy',
                                  'datum', 'dodavatel', 'schvalil', 'file']

                for i in ['cena_zahr_dph', 'cena_bez_dph']:
                    if i in df.columns.values:
                        df[i] = df[i].str.replace('(eur)|\s', '', regex=True).str.replace(',', '.', regex=True)
                        df[i] = np.where(df[i].str.match(r'\d*\.\d*\.\d*'),
                                                      df[i].str.replace('\.', '', 1, regex=True),
                                                      df[i])
                        df.loc[df[i] == '', i] = '0'
                        df.loc[df[i] == '.', i] = '0'
                        df[i] = df[i].astype(float)

                df_all = pd.concat([df_all, df], ignore_index=True)
    except Exception:
        print(traceback.format_exc())
        df2=df
        break

df_all.drop(df_all[df_all['objednavka_predmet']==''].index, inplace=True)
df_all['datum']=df_all['datum'].str.replace('012', '12')

# date
df_all['datum'] = df_all['datum'].apply(get_dates)
df_all['cena'] = np.where(df_all.cena_zahr_dph > df_all.cena_bez_dph, df_all.cena_zahr_dph, df_all.cena_bez_dph)
df_all['cena_s_dph'] = np.where(df_all.cena_zahr_dph > df_all.cena_bez_dph, 'ano', 'nie')
df_all['insert_date'] = datetime.now()


kysuckanemocnica.df_all = df_all
kysuckanemocnica.dodavatel_list=['dodavatel']
kysuckanemocnica.popis_list=['objednavka_predmet', 'objednavka_cislo', 'cislo_zmluvy', 'cena_s_dph']
kysuckanemocnica.create_columns_w_dict('kysucka nemocnica s poliklinikou cadca')

kysuckanemocnica.df_all.drop_duplicates(inplace=True)

kn_search = pd.DataFrame(kysuckanemocnica.df_all[kysuckanemocnica.final_table_cols])

func.save_df(kn_search, 'kysucka_nemocnica_kuchyna.xlsx', kysuckanemocnica.hosp_path)

# oddelenie IT
dict_tables = {}
for file_name in os.listdir(kysuckanemocnica.hosp_path):

        print(file_name)
        try:
            if re.match('.*Oddelenie IT.*', file_name):
                dict_tables[file_name]= camelot.read_pdf(os.path.join(kysuckanemocnica.hosp_path, file_name), pages='all', flavor='stream',
                                                                  row_tol=30, split_text=True)

        except Exception:
            print(file_name)
            print(traceback.format_exc())
            continue


df_all = pd.DataFrame()

for key in dict_tables.keys():
    try:
        for i in range(len(dict_tables[key])):
            df = dict_tables[key][i].df
            print(key)
            df = func.clean_str_cols(df)

            mask = df.applymap(lambda cell: bool(
                re.search('(.*celkova.*)|(bez dph.*)|(.*ratane dph.*)|(identifikacne udaje.+)|(objedna\s*neho.*)|(meno/nazov,.+)|(osoby/ obchodne.+)|(.+miesto podnikania.+)|(meno a priezvisko.+)|(pobytu fo.+)',
                          cell))).any(axis=1)
            df = df.drop(df[mask].index).reset_index(drop=True)
            df['file'] = key

            if not df.empty:
                if (key in ('13-04-2023Oddelenie IT - December 2019.pdf', '13-04-2023Oddelenie IT - November 2019.pdf') and i ==0 ):
                    df.columns = ['objednavka_cislo', 'objednavka_predmet', 'prazdny_stlpec1', 'cena_zahr_dph',
                                  'dodavatel', 'schvalil', 'file']
                elif (key in ('13-04-2023Oddelenie IT - December 2022.pdf') and i ==0 ):
                    df.columns = ['objednavka_cislo', 'prazdny_stlpec1', 'cena_zahr_dph', 'prazdny_stlpec2', 'datum',
                                  'dodavatel', 'schvalil', 'file']
                elif (key in ('13-04-2023Oddelenie IT - Jún 2021.pdf') and i ==0 ):
                    df.columns = ['objednavka_cislo', 'prazdny_stlpec1', 'cena_zahr_dph', 'prazdny_stlpec2',
                                  'datum', 'dodavatel', 'schvalil', 'file']
                elif (key in ('13-04-2023Oddelenie IT - Marec 2019.pdf', '13-04-2023Oddelenie IT - Máj 2019.pdf') and i == 0 ) or \
                        ((key in ('13-04-2023Oddelenie IT - Máj 2019.pdf') and i in (1,2) )):
                    df.columns = ['objednavka_cislo', 'objednavka_predmet', 'prazdny_stlpec1', 'cena_zahr_dph',
                                  'datum', 'dodavatel', 'schvalil', 'file']
                else:
                    df.columns = ['objednavka_cislo', 'objednavka_predmet', 'prazdny_stlpec1', 'cena_zahr_dph', 'prazdny_stlpec2',
                                  'datum', 'dodavatel', 'schvalil', 'file']
                df_all = pd.concat([df_all, df], ignore_index=True)
    except Exception:
        print(traceback.format_exc())
        df2=df
        break


df = dict_tables['13-04-2023Oddelenie IT - November 2019.pdf'][0].df

# datum
df_all['datum_extr'] = df_all['cena_zahr_dph'].str.extract(r'(\d{2}\.\d{2}\.\d{4})')
df_all.loc[pd.isna(df_all.datum), 'datum'] = df_all['datum_extr']
df_all['datum'] = df_all['datum'].str.replace('[a-z]', '', regex=True)
df_all['datum'] = df_all['datum'].apply(get_dates)

# predmet objednavky
df_all.loc[pd.isna(df_all.objednavka_predmet), 'objednavka_predmet'] = df_all['objednavka_cislo']
df_all['objednavka_predmet'] = df_all['objednavka_predmet'].str.replace('mm-\d{4}-\d{3}', '', regex=True)

# cena
df_all['cena'] = df_all['cena_zahr_dph'].str.replace(r'(\d{1,}\.\d{2}\.\d{4})|([a-z])', '', regex=True).str.replace(',',
                                                                                                                    '.',
                                                                                                                    regex=True)
df_all['cena_ext'] = df_all['cena'].str.extract(r'((?<=\s\s)\d+\s*\d*\.\d+)')
df_all['cena'] = df_all['cena'].str.replace(r'((?<=\s\s)\d+\s*\d*\.\d+)', '', regex=True)
df_all['cena'] = df_all['cena'].str.replace(r'\s', '', regex=True)
df_all['cena_ext'] = df_all['cena_ext'].str.replace(r'\s', '', regex=True)
df_all.loc[df_all['cena'] == '', ['cena', 'cena_ext']] = 0
df_all['cena'] = df_all['cena'].astype(float)
df_all['cena_ext'] = df_all['cena_ext'].astype(float)
df_all.loc[pd.isna(df_all['cena_ext']) == False, 'cena'] = df_all['cena'] + df_all['cena_ext']
df_all=df_all.assign(cena_s_dph='ano', insert_date=datetime.now())

kysuckanemocnica.df_all = df_all
kysuckanemocnica.dodavatel_list=['dodavatel']
kysuckanemocnica.popis_list=['objednavka_predmet', 'objednavka_cislo',  'cena_s_dph']
kysuckanemocnica.create_columns_w_dict('kysucka nemocnica s poliklinikou cadca')

kysuckanemocnica.df_all.drop_duplicates(inplace=True)

kn_search = pd.DataFrame(kysuckanemocnica.df_all[kysuckanemocnica.final_table_cols])

func.save_df(kn_search, 'kysucka_nemocnica_oddelenie_it.xlsx', kysuckanemocnica.hosp_path)

# zdravotnícka technika

dict_tables = {}
for file_name in os.listdir(kysuckanemocnica.hosp_path):

        print(file_name)
        try:
            if re.match('.*Zdravotnícka technika.*', file_name):
                dict_tables[file_name]= camelot.read_pdf(os.path.join(kysuckanemocnica.hosp_path, file_name), pages='all', flavor='stream',
                                                                  row_tol=30, split_text=True)

        except Exception:
            print(file_name)
            print(traceback.format_exc())
            continue

df_all = pd.DataFrame()

for key in dict_tables.keys():
    try:
        for i in range(len(dict_tables[key])):
            df = dict_tables[key][i].df
            print(key)
            df = func.clean_str_cols(df)

            mask = df.applymap(lambda cell: bool(
                re.search('(.*celkova.*)|(bez dph.*)|(.*ratane dph.*)|(identifikacne udaje.+)|(objedna\s*neho.*)|(meno/nazov,.+)|(osoby/ obchodne.+)|(.+miesto podnikania.+)|(meno a priezvisko.+)|(pobytu fo.+)',
                          cell))).any(axis=1)
            df = df.drop(df[mask].index).reset_index(drop=True)
            df['file'] = key

            if not df.empty:
                if (key in ('13-04-2023 Zdravotnícka technika - Február 2022.pdf') and i ==1 ) or (key == '13-04-2023 Zdravotnícka technika - Október 2022.pdf' and i ==0)\
                        or (key == '13-04-2023Zdravotnícka technika - August 2019.pdf' and i in (0,1))\
                        or (key == '13-04-2023Zdravotnícka technika - December 2022.pdf' and i ==1)\
                        or (key == '13-04-2023Zdravotnícka technika - Január 2020.pdf' and i ==1)\
                        or (key == '13-04-2023Zdravotnícka technika - Január 2022.pdf' and i ==1)\
                        or (key == '13-04-2023Zdravotnícka technika - Júl 2021.pdf' and i ==2)\
                        or (key == '13-04-2023Zdravotnícka technika - Júl 2022.pdf' and i ==1) or \
                        (key == '13-04-2023Zdravotnícka technika - Máj 2019.pdf' and i ==2) or \
                        (key == '13-04-2023Zdravotnícka technika - Máj 2022.pdf' and i in (1,2)) or \
                        (key == '13-04-2023Zdravotnícka technika - November 2019.pdf' and i ==1) or \
                        (key == '13-04-2023Zdravotnícka technika - November 2021.pdf' and i ==1)\
                        or (key == '13-04-2023Zdravotnícka technika - Október 2019.pdf' and i ==0)\
                        or (key == '13-04-2023Zdravotnícka technika - Október 2021.pdf' and i ==0):
                    df.columns = ['objednavka_cislo', 'objednavka_predmet', 'prazdny_stlpec1', 'cena_zahr_dph',  'datum',
                                  'dodavatel', 'schvalil', 'file']
                elif (key in ('13-04-2023Zdravotnícka technika - Apríl 2018.pdf') and i == 1 ) or \
                        (key in '13-04-2023Zdravotnícka technika - Jún 2018.pdf' and i == 1) or \
                        (key in ('13-04-2023Zdravotnícka technika - Marec 2018.pdf') and i in (1,2)) or \
                        (key in ('13-04-2023Zdravotnícka technika - Marec 2019.pdf') and i ==1) or \
                        (key in ('13-04-2023Zdravotnícka technika - Máj 2018.pdf') and i ==1) or \
                        (key in ('13-04-2023Zdravotnícka technika - Október 2017.pdf') and i == 1) or \
                        (key in ('13-04-2023Zdravotnícka technika - Október 2018.pdf') and i ==1):
                    df.columns = ['objednavka_cislo', 'objednavka_predmet',  'cena_zahr_dph', 'datum',
                                  'dodavatel', 'schvalil', 'file']
                elif (key == '13-04-2023Zdravotnícka technika - August 2018.pdf' and i in (0,1))\
                        or (key=='13-04-2023Zdravotnícka technika - December 2022.pdf' and i == 2)\
                        or (key=='13-04-2023Zdravotnícka technika - február 2023.pdf' and i == 2):
                    df.columns = ['objednavka_cislo', 'objednavka_predmet', 'prazdny_stlpec1', 'cena_zahr_dph', 'dodavatel', 'schvalil', 'file']
                elif (key in ('13-04-2023Zdravotnícka technika - September 2019.pdf') and i == 0 ):
                    df.columns = ['objednavka_cislo', 'prazdny_stlpec1', 'cena_zahr_dph', 'prazdny_stlpec2',
                                      'datum', 'dodavatel', 'schvalil', 'file']
                elif (key in ('13-04-2023Zdravotnícka technika - September 2019.pdf') and i == 1 ):
                    df.columns = ['objednavka_cislo', 'prazdny_stlpec1', 'cena_zahr_dph', 'datum', 'dodavatel', 'schvalil', 'file']
                elif (key in ('13-04-2023Zdravotnícka technika - September 2019.pdf') and i == 2 )\
                        or ( key == '13-04-2023Zdravotnícka technika - September 2022.pdf' and i ==1):
                    df.columns = ['objednavka_cislo', 'objednavka_predmet', 'prazdny_stlpec1', 'cena_zahr_dph', 'datum', 'dodavatel', 'schvalil', 'file']
                else:
                    df.columns = ['objednavka_cislo', 'objednavka_predmet', 'prazdny_stlpec1', 'cena_zahr_dph', 'prazdny_stlpec2',
                                      'datum', 'dodavatel', 'schvalil', 'file']

                # for i in ['cena_zahr_dph', 'cena_bez_dph']:
                #     if i in df.columns.values:
                #         df[i] = df[i].str.replace('(eur)|\s', '', regex=True).str.replace(',', '.', regex=True)
                #         df[i] = np.where(df[i].str.match(r'\d*\.\d*\.\d*'),
                #                                       df[i].str.replace('\.', '', 1, regex=True),
                #                                       df[i])
                #         df.loc[df[i] == '', i] = '0'
                #         df.loc[df[i] == '.', i] = '0'
                #         df[i] = df[i].astype(float)

                df_all = pd.concat([df_all, df], ignore_index=True)
    except Exception:
        print(traceback.format_exc())
        df2=df
        break

func.save_df(df_all, 'kysucka_nemocnica_zdravotnicka_technika.xlsx', kysuckanemocnica.hosp_path)

df = dict_tables['13-04-2023Zdravotnícka technika - Apríl 2018.pdf'][0].df


df_all = func.load_df('kysucka_nemocnica_zdravotnicka_technika.xlsx', kysuckanemocnica.hosp_path)

# date
df_all['datum_extr'] = df_all['cena_zahr_dph'].str.extract('([a-z]+\s\d{1,2}\.\s[a-z]+\s\d{4})', expand=False)
df_all['datum_extr2'] = df_all['cena_zahr_dph'].str.extract(r'(\d{2}\.\d{2}\.\d{3,})')
df_all['datum_extr2'] = df_all['datum_extr2'].str.replace('202', '2023')
df_all.loc[pd.isna(df_all['datum_extr'])==False, 'datum']=df_all['datum_extr']
df_all.loc[pd.isna(df_all['datum_extr2'])==False, 'datum']=df_all['datum_extr2']

df_all['datum'] = df_all['datum'].str.replace(r'(?<=\d{4})\s[a-z]+\,*\s\d{1,2}\.\s[a-z]+\s\d{4}', '', regex=True)
df_all['datum'] = df_all['datum'].str.replace('(?<=\d{4})\s[a-z]+\s\d{1,2}\.\s[a-z]+\s\d{4}', '', regex=True)
df_all['datum'] = df_all['datum'].str.replace('(?<=\d{4})\s\d{1,2}\.\d{1,2}\.\d{4}', '', regex=True)
df_all['datum'] = df_all['datum'].apply(get_dates)

# cena
df_all['cena_zahr_dph'] = df_all['cena_zahr_dph'].str.replace('([a-z]+\s\d{1,2}\.\s[a-z]+\s\d{4})', '', regex=True)
df_all['cena_zahr_dph'] = df_all['cena_zahr_dph'].str.replace('eur', '')

df_all['cena_ext'] = df_all['cena_zahr_dph'].str.extract(r'((?<=\s\s)\d+\s*\d*\,\d+)')
df_all['cena_zahr_dph'] = df_all['cena_zahr_dph'].str.replace(r'(?<=\s\s)\d+\s*\d*\,\d+', '', regex=True)

df_all['cena2'] = df_all['cena_zahr_dph'].str.extract('(\d{0,3}\s*\d{1,3}\,\d{2})')
df_all['cena2']=df_all['cena2'].str.replace(',', '.', regex=True).str.replace('\s', '', regex=True)
df_all['cena_ext']=df_all['cena_ext'].str.replace(',', '.', regex=True).str.replace('\s', '', regex=True)
df_all['cena2']=df_all['cena2'].astype(float)
df_all['cena_ext']=df_all['cena_ext'].astype(float)
# TODO

df_all.loc[(df_all['cena_zahr_dph'].str.contains('[a-z]')==True), 'poznamka_k_cene']=df_all['cena_zahr_dph'].str.strip()
df_all.assign(cena_s_dph='ano', insert_date=datetime.now())

# kysuckanemocnica.df_all = df_all
# kysuckanemocnica.dodavatel_list=['dodavatel']
# kysuckanemocnica.popis_list=['objednavka_predmet', 'objednavka_cislo',  'cena_s_dph', 'poznamka_k_cene', 'schvalil']
# kysuckanemocnica.create_columns_w_dict('kysucka nemocnica s poliklinikou cadca')
#
# kysuckanemocnica.df_all.drop_duplicates(inplace=True)
#
# kn_search = pd.DataFrame(kysuckanemocnica.df_all[kysuckanemocnica.final_table_cols])



#388

df = df_all[['cena_zahr_dph', 'cena2']]

