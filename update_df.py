import pandas as pd
from config import *
from functions_ZK import *
from schemas import OutlookTools, PriameObjednavkyMail
import win32com.client
import functionss as func
import shutil

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
otl = OutlookTools(outlook)
path = outlook.Folders['obstaravanie'].Folders['Doručená pošta'].Folders['Priame objednávky']

### FNsPZA ###

hosp = 'fnspza'
hosp_path = data_path + hosp + "\\"
hosp_path_hist = historical_data_path + hosp + "\\"

fnspza_df_original = func.load_df(os.path.join(search_data_path + 'fnspza_all.pkl'), path=os.getcwd())

# get last insert date
last_update_str = fnspza_df_original['insert_date'].max().strftime('%d.%m.%Y %I:%M %p')

# filter mails and download attachments
search_result = otl.find_message(path,
                                 "@SQL= (""urn:schemas:httpmail:fromemail"" LIKE '%" + hosp + '.sk' + "') AND (""urn:schemas:httpmail:datereceived"" > '" + last_update_str + "') ")

otl.save_attachment(hosp_path, search_result)

# load and clean data
all_tables = load_files(hosp_path)
all_tables_cleaned = clean_tables(all_tables)
fnspza_df_new = create_table(all_tables_cleaned, stand_column_names)

fnspza_df_new['objednavatel'] = hosp
fnspza_df_new['link'] = dict['fakultna nemocnica s poliklinikou zilina']['zverejnovanie_objednavok_faktur_rozne']

fnspza_df_new2 = fnspza_data_cleaning(fnspza_df_new)

# standardize column names
fnspza_df_search = pd.DataFrame(
    fnspza_df_new2[['objednavatel', 'cena', 'datum_adj', 'dodavatel', 'popis', 'insert_date', 'file', 'link']])
fnspza_df_search.columns = ['objednavatel', 'cena', 'datum', 'dodavatel', 'popis', 'insert_date', 'file', 'link']

# concat with old data and save
fnspza_df_search = pd.concat([fnspza_df_original, fnspza_df_search], ignore_index=True)

with pd.ExcelWriter(os.path.join(search_data_path + 'fnspza_all.xlsx'), engine='xlsxwriter',
                    engine_kwargs={'options': {'strings_to_urls': False}}) as writer:
    fnspza_df_search.to_excel(writer)
func.save_df(df=fnspza_df_search, name=os.path.join(search_data_path + 'fnspza_all.pkl'))

move_all_files(source_path=hosp_path, destination_path=hosp_path_hist)




