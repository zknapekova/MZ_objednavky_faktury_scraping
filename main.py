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
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.common.exceptions import ElementNotVisibleException, TimeoutException

current_date_time = datetime.now()
print('Start:', current_date_time)

# load excel and create dictionary
df = func.load_df(source_path)
df['nazov'] = df['Nazov_full']

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
urlretrieve(dict['detska psychiatricka liecebna n o hran']['objednavky_faktury_link'],
            data_path + current_date_time.strftime("%d-%m-%Y") + str(keysList[5]).replace(" ", "_") +
            dict['detska psychiatricka liecebna n o hran']['objednavky_faktury_file_ext'])

# 6 - Fakultna nemocnica Nitra
df_fnnr = pd.read_html(dict['fakultna nemocnica nitra']['objednavky_faktury_link'], encoding='utf-8')[0]
df_fnnr = func.clean_str_cols(df_fnnr)

# 7 - fakultna nemocnica s poliklinikou f d roosevelta banska bystrica

output = FNsP_BB_objednavky(
    link=dict['fakultna nemocnica s poliklinikou f d roosevelta banska bystrica']['objednavky_faktury_link'],
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
driver.get(dict['fakultna nemocnica s poliklinikou zilina']['objednavky_link'])
n_records_dropdown = driver.find_element(By.ID, "limit1")

select = Select(n_records_dropdown)
select.select_by_value('100')
table_lst = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "list_1_com_fabrik_1"))).get_attribute(
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

result_df2 = func.load_df('fnspza_all_web.pkl', path=os.getcwd())
result_df2=clean_str_col_names(result_df2)
result_df2.columns=['objednavka_cislo1', 'cpv kod', 'objednavka_predmet', 'cena', 'pocet mj', 'kategoria', 'pridane']

result_df2['cena'] = result_df2['cena'].str.replace(r"[a-z|'|\s|-|\(\)]+", '', regex=True)
result_df2['cena'] = result_df2['cena'].str.replace(r",", '.', regex=True)
result_df2['cena']=result_df2['cena'].astype(float)

result_df2['rok_objednavky'] = result_df2['objednavka_cislo1'].str.extract(r'(20\d{2})')
result_df2['objednavka_cislo'] = result_df2['objednavka_cislo1'].apply(lambda x: x.split('/')[-1])
result_df2['objednavka_cislo'] = result_df2['objednavka_cislo'].str.replace(r'^0+', '', regex=True)
result_df2['objednavka_cislo'] = result_df2['objednavka_cislo'].apply(
    lambda x: x.split('-')[-1] + '-' + x.split('-')[0] if('-' in x) else x)

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

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
otl = OutlookTools(outlook)

path = outlook.Folders['obstaravanie'].Folders['Doručená pošta'].Folders['Priame objednávky']

### FNsPZA ###

stand_column_names = {
    'objednavatel': ['nazov verejneho obstaravatela'],
    'kategoria': ['kategoria zakazky(tovar/stavebna praca/sluzba)', 'kategoria(tovar/stavebna praca/sluzba)',
                  'kategoria(tovary / prace / sluzby)'],
    'objednavka_predmet': ['nazov predmetu objednavky', 'predmet objednavky'],
    'cena': ['hodnotaobjednavkyv eur bez dph', 's.nc bdph', 'hodnotaobjednavkyv eur bez dph',
                                   'hodnota'],
    'datum': ['datum zadania objednavky', 'datum objednavky'],
    'objednavka_cislo': ['c.obj.', 'cislo objednavky'],
    'zdroj_financovania': ['zdroje financovania'],
    'balenie': ['balenie'],
    'sukl_kod': ['sukl_kod'],
    'mnozstvo': ['mnozstvo'],
    'poznamka': ['kratke zdovodnenie', 'kratke zdovodnenie2'],
    'dodavatel_ico': ['dodavatel - ico'],
    'dodavatel_nazov': ['dodavatel - nazov'],
    'odkaz_na_zmluvu': ['odkaz na zverejnenu zmluvu'],
    'pocet_oslovenych': ['pocet oslovenych']
}

# download all attachements available from outlook
search_result = otl.find_message(path, "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%fnspza.sk' ")
hosp_path = data_path + "fnspza\\"
otl.save_attachement(hosp_path, search_result)

all_tables = load_files(hosp_path)

# remove rows outside of table
all_tables_cleaned = clean_tables(all_tables)
all_tables_cleaned[2][3]['sukl_kod'] = all_tables_cleaned[2][3]['sukl_kod1'].str.cat(
    all_tables_cleaned[2][3]['kod'].astype(str), sep='')

fnspza_all = create_table(all_tables_cleaned, stand_column_names)

### data cleaning ###
fnspza_all['objednavatel']='fnspza'
fnspza_all['link'] = dict['fakultna nemocnica s poliklinikou zilina']['zverejnovanie_objednavok_faktur_rozne']

fnspza_all = func.load_df('fnspza_all.pkl', path=os.getcwd())
fnspza_all2 = func.clean_str_cols(fnspza_all)

# predmet objednavky
fnspza_all2['extr_mnozstvo'] = fnspza_all2['objednavka_predmet'].str.extract(r'(\s+\d+x$)')
fnspza_all2['mnozstvo'] = np.where((pd.isna(fnspza_all2['mnozstvo'])) & (
            pd.isna(fnspza_all2['extr_mnozstvo']) == False), fnspza_all2['extr_mnozstvo'].str.strip(),
                                    fnspza_all2['mnozstvo'])
fnspza_all2['objednavka_predmet'] = fnspza_all2['objednavka_predmet'].str.replace(r'\s+\d+x$', '', regex=True)
fnspza_all2.drop(['extr_mnozstvo'], axis=1, inplace=True)

# cena
fnspza_all2['cena'] = fnspza_all2['cena'].str.replace(r'^mc\s*', '', regex=True)
# remove ',-' and  '.-'
fnspza_all2['cena'] = fnspza_all2['cena'].str.replace(r'[,|\.]-.*', '', regex=True)
fnspza_all2['cena'] = fnspza_all2['cena'].str.replace(r'[a-z]+\.[a-z]*\s*', '', regex=True)
fnspza_all2['cena'] = fnspza_all2['cena'].str.replace(r"[a-z]|'|\s|-|[\(\)]+", '', regex=True)
fnspza_all2['cena'] = fnspza_all2['cena'].str.replace(r",,", '.', regex=True)
fnspza_all2['cena'] = fnspza_all2['cena'].str.replace(r",", '.', regex=True)
fnspza_all2['cena'] = fnspza_all2['cena'].str.replace(r".*[:].*", '0', regex=True)
fnspza_all2['cena'] = fnspza_all2['cena'].str.replace(r"=.*", '', regex=True)

fnspza_all2['cena'] = np.where(fnspza_all2['cena'].str.match(r'\d*\.\d*\.\d*'), fnspza_all2['cena'].str.replace('.', '', 1), fnspza_all2['cena'])

fnspza_all2['cena'] = fnspza_all2['cena'].astype(float)

# rok objednavky
fnspza_all2['rok_objednavky']=fnspza_all2['datum'].str.extract(r'(20\d{2})')

# datum objednavky
fnspza_all2['datum'] = pd.to_datetime(fnspza_all2['datum'], errors='ignore')


# popis
popis_list = ['objednavka_predmet', 'kategoria', 'objednavka_cislo', 'zdroj_financovania', 'balenie',
                                    'sukl_kod', 'mnozstvo', 'poznamka', 'odkaz_na_zmluvu', 'pocet_oslovenych']
fnspza_all2['popis'] = fnspza_all2[popis_list].T.apply(lambda x: x.dropna().to_dict())


# save df
with pd.ExcelWriter('output.xlsx', engine='xlsxwriter', engine_kwargs={'options':{'strings_to_urls': False}}) as writer:
    fnspza_all2.to_excel(writer)
func.save_df(df=fnspza_all2, name='fnspza_all.pkl')

# comparison and pairing

fnspza_all2_unique = fnspza_all2.drop_duplicates()
result_df2_unique = result_df2.drop_duplicates()
# vsetky_unikatne_objednavky = fnspza_all2_unique.shape[0]
# chybajuce_cislo_objednavky = sum(fnspza_all2_unique['objednavka_cislo'].isna())

tab_merged = result_df2_unique.merge(fnspza_all2_unique, on=['objednavka_cislo', 'cena', 'rok_objednavky']).sort_values(
    by='objednavka_cislo')
tab_merged_povod = pd.merge(fnspza_all2_unique, result_df2_unique, left_on='objednavka_cislo',
                            right_on='objednavka_cislo1')

# cisla objednavky ktore sa nepodarilo sparovat
objednavky_paired_list = list(set(tab_merged['objednavka_cislo'])) + list(set(tab_merged_povod['objednavka_cislo_x']))
objednavky_notpaired_list = []
dokumenty_list = []
cena_list = []
datum_list = []
popis_list = []
objednavatel_list = []
dodavatel_nazov_list = []
dodavatel_ico_list = []

tab_filtered = fnspza_all2_unique[pd.isna(fnspza_all2_unique['objednavka_cislo']) == False].reset_index(
    drop=True).sort_values(by='objednavka_cislo')

for i in range(tab_filtered.shape[0]):
    if tab_filtered['objednavka_cislo'].at[i] not in objednavky_paired_list and tab_filtered['rok_objednavky'].at[
        i] == '2022':
        objednavky_notpaired_list.append(tab_filtered['objednavka_cislo'].at[i])
        dokumenty_list.append(tab_filtered['file'].at[i])
        cena_list.append(tab_filtered['cena'].at[i])
        datum_list.append(tab_filtered['datum'].at[i])
        popis_list.append(tab_filtered['popis'].at[i])
        objednavatel_list.append(tab_filtered['objednavatel'].at[i])
        dodavatel_nazov_list.append(tab_filtered['dodavatel_nazov'].at[i])
        dodavatel_ico_list.append(tab_filtered['dodavatel_ico'].at[i])

# nesparovane objednavky z excelov - 22387/31560 ~ 70%
objednavky_notpaired_list = list(set((objednavky_notpaired_list)))

# nesparovane objednavky z webu 32348/41638 ~ 77%
objednavky_web_notpaired_list = []
cpv_kod_list = []
objednavka_web_predmet_list = []
cena_web_list = []
pocet_mj = []
kategoria_web_list = []
pridane_web = []

l = list(set(tab_merged['objednavka_cislo1'])) + list(set(tab_merged_povod['objednavka_cislo_x']))
result_df2_unique = result_df2_unique.reset_index(drop=True)
result_df2_unique['pridane_rok'] = result_df2_unique['pridane'].apply(lambda x: x.split('.')[-1])
for i in range(result_df2_unique.shape[0]):
    if (result_df2_unique['objednavka_cislo1'].at[i] not in l) and (result_df2_unique['pridane_rok'].at[i] == '2022'):
        objednavky_web_notpaired_list.append(result_df2_unique['objednavka_cislo1'].at[i])
        cpv_kod_list.append(result_df2_unique['cpv kod'].at[i])
        objednavka_web_predmet_list.append(result_df2_unique['objednavka_predmet'].at[i])
        cena_web_list.append(result_df2_unique['cena'].at[i])
        pocet_mj.append(result_df2_unique['pocet mj'].at[i])
        kategoria_web_list.append(result_df2_unique['kategoria'].at[i])
        pridane_web.append(result_df2_unique['pridane'].at[i])

objednavky_web_notpaired_list = list(set(objednavky_web_notpaired_list))

### concatenating together ###
df_excels = pd.DataFrame(list(
    zip(objednavky_notpaired_list, dokumenty_list, cena_list, datum_list, objednavatel_list, popis_list,
        dodavatel_nazov_list, dodavatel_ico_list)),
                         columns=['cislo_objednavky', 'zdroj', 'cena', 'datum', 'objednavatel', 'popis',
                                  'dodavatel_nazov', 'dodavatel_ico'])
popis_list3 = ['dodavatel_nazov', 'dodavatel_ico']
df_excels['dodavatel'] = df_excels[popis_list3].T.apply(lambda x: x.dropna().to_dict())
df_web = pd.DataFrame(list(
    zip(objednavky_web_notpaired_list, cpv_kod_list, objednavka_web_predmet_list, cena_web_list, pocet_mj,
        kategoria_web_list, pridane_web)),
                      columns=['cislo_objednavky', 'cpv kod', 'objednavka_predmet', 'cena', 'pocet mj', 'kategoria',
                               'datum'])
df_web['zdroj'] = 'web'
df_web['objednavatel'] = 'fnspza'
df_web['dodavatel'] = np.nan
popis_list2 = ['cpv kod', 'objednavka_predmet', 'kategoria', 'pocet mj']
df_web['popis'] = df_web[popis_list2].T.apply(lambda x: x.dropna().to_dict())

df_final = pd.concat([df_web[['cislo_objednavky', 'zdroj', 'cena', 'datum', 'objednavatel', 'popis']],
                      df_excels[['cislo_objednavky', 'zdroj', 'cena', 'datum', 'objednavatel', 'popis', 'dodavatel']]],
                     ignore_index=True)
df_final.to_excel('nenapojene_objednavky_2022_v2.xlsx')
