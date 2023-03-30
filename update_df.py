import pandas as pd
from config import *
from functions_ZK import *
from schemas import OutlookTools, PriameObjednavkyMail, ObjednavkyDB
import win32com.client
import functionss as func
import shutil
from mysql_config import objednavky_db_connection
import logging
import sys
import traceback
from exceptions import DataNotAvailable
import shutil

logger = logging.getLogger('update_df')
logging.basicConfig(filename="log.txt", format='[%(asctime)s] %(levelname)s:  %(message)s', datefmt="%Y-%m-%d %H:%M:%S")
logger.setLevel(logging.INFO)
console = logging.StreamHandler()
logger.addHandler(console)

try:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    otl = OutlookTools(outlook)
    path = outlook.Folders['obstaravanie'].Folders['Doručená pošta'].Folders['Priame objednávky']
    logger.info(f"Outlook path loaded successfully")
except Exception as e:
    logger.error(traceback.format_exc())
    sys.exit()

try:
    db = ObjednavkyDB(objednavky_db_connection)
except Exception as e:
    logger.error(traceback.format_exc())
    sys.exit()

# create backup and load original df
shutil.copyfile('priame_objednavky_all.pkl', 'priame_objednavky_all_backup.pkl')
df_orig = func.load_df('priame_objednavky_all.pkl', path=os.getcwd())


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
fnspza_df_new['link'] = dict_all['fakultna nemocnica s poliklinikou zilina']['zverejnovanie_objednavok_faktur_rozne']

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

# FNNR
#
# try:
#     obj = PriameObjednavkyMail('fnnitra')
#     logger.info(f'{obj.hosp} data load started')
#
#     # get last insert date
#     last_update_str = df_orig['insert_date'][df_orig['objednavatel'] == obj.hosp].max().strftime('%d.%m.%Y %I:%M %p')
#     # filter mails and download attachments
#     search_result = otl.find_message(path,
#                                      "@SQL= (""urn:schemas:httpmail:fromemail"" LIKE '%" + obj.hosp + '.sk' + "') AND (""urn:schemas:httpmail:datereceived"" > '" + last_update_str + "') ")
#     otl.save_attachment(obj.hosp_path, search_result)
#
#     obj.load()
#     obj.clean_tables()
#     obj.create_table(stand_column_names=stand_column_names)
#     if obj.df_all.empty:
#         raise DataNotAvailable(obj.hosp)
#
#     logger.info(f'Number of rows loaded: {obj.df_all.shape[0]}')
#     obj.create_columns_w_dict(key='fakultna nemocnica nitra')
#
#     # data cleaning #
#     # cena
#     dict_cena = {"[a-z]|'|\s|[\(\)]+": '', '[,|\.]-.*': '', '-,': '', ",": '.', '\.\.': '.', '/.*/*': '', '\..\.\+': '',
#                  '/\..*': '', '.*:': ''}
#     obj.df_all = str_col_replace(obj.df_all, 'cena', dict_cena)
#     obj.df_all['cena'] = obj.df_all['cena'].astype(float)
#     # datum
#     obj.df_all['datum'] = obj.df_all['datum'].str.strip()
#     obj.df_all['datum'] = obj.df_all['datum'].apply(get_dates)
#
#     # create final table
#     df_search = pd.DataFrame(obj.df_all[obj.final_table_cols])
#     df_orig = pd.concat([df_orig, df_search], ignore_index=True)
#
#     # save tables
#     db.insert_table(table_name='priame_objednavky', df=obj.df_all, if_exists='append', index=False)
#     move_all_files(source_path=obj.hosp_path, destination_path=obj.hosp_path_hist)
#     logger.info(f'{obj.hosp} data load finished successfully')
#
# except DataNotAvailable as e:
#     logger.info(e.message)
#     pass
# except Exception:
#     logger.error(f"Error in hosp: {obj.hosp}")
#     logger.error(traceback.format_exc())
#     pass

# # FNsP Presov
# try:
#     obj = PriameObjednavkyMail('fnsppresov')
#     logger.info(f'{obj.hosp} data load started')
#
#     # get last insert date
#     last_update_str = df_orig['insert_date'][df_orig['objednavatel'] == obj.hosp].max().strftime('%d.%m.%Y %I:%M %p')
#
#     # filter mails and download attachments
#     search_result = otl.find_message(path,
#                                          "@SQL= (""urn:schemas:httpmail:fromemail"" LIKE '%" + obj.hosp + '.sk' + "') AND (""urn:schemas:httpmail:datereceived"" > '" + last_update_str + "') ")
#     otl.save_attachment(obj.hosp_path, search_result)
#
#     obj.load()
#     obj.clean_tables()
#     obj.create_table(stand_column_names=stand_column_names)
#
#     if obj.df_all.empty:
#         raise DataNotAvailable(obj.hosp)
#
#     logger.info(f'Number of rows loaded: {obj.df_all.shape[0]}')
#     obj.df_all['cena'] = obj.df_all['cena'].astype(float)
#     obj.df_all['datum'] = obj.df_all['datum'].str.strip()
#     obj.df_all['datum'] = obj.df_all['datum'].apply(get_dates)
#
#     obj.create_columns_w_dict(key='fakultna nemocnica s poliklinikou j a reimana presov')
#     df_search = pd.DataFrame(obj.df_all[obj.final_table_cols])
#     df_orig = pd.concat([df_orig, df_search], ignore_index=True)
#
#     # save tables
#     db.insert_table(table_name='priame_objednavky', df=obj.df_all, if_exists='append', index=False)
#     move_all_files(source_path=obj.hosp_path, destination_path=obj.hosp_path_hist)
#     logger.info(f'{obj.hosp} data load finished successfully')
#
# except DataNotAvailable as e:
#     logger.info(e.message)
#     pass
#
# except Exception:
#     logger.error(f"Error in hosp: {obj.hosp}")
#     logger.error(traceback.format_exc())
#     pass

# # FNTT  ###
#
# try:
#     obj = PriameObjednavkyMail('fntt')
#     logger.info(f'{obj.hosp} data load started')
#
#     # get last insert date
#     last_update_str = df_orig['insert_date'][df_orig['objednavatel'] == obj.hosp].max().strftime('%d.%m.%Y %I:%M %p')
#
#     # filter mails and download attachments
#     search_result = otl.find_message(path,
#                                          "@SQL= (""urn:schemas:httpmail:fromemail"" LIKE '%" + obj.hosp + '.sk' + "') AND (""urn:schemas:httpmail:datereceived"" > '" + last_update_str + "') ")
#     otl.save_attachment(obj.hosp_path, search_result)
#
#     obj.load()
#     obj.clean_tables()
#     obj.create_table(stand_column_names=stand_column_names)
#
#     if obj.df_all.empty:
#         raise DataNotAvailable(obj.hosp)
#
#     logger.info(f'Number of rows loaded: {obj.df_all.shape[0]}')
#     dict_cena = {"[a-z]|'|\s|[\(\)]+": "", ',-': '', ',$': '', ',+': '.', '/.*': ''}
#     obj.df_all = str_col_replace(obj.df_all, 'cena', dict_cena)
#     obj.df_all['cena'] = obj.df_all['cena'].astype(float)
#
#     obj.df_all['datum'] = obj.df_all['datum'].str.strip()
#     obj.df_all['datum'] = obj.df_all['datum'].apply(get_dates)
#
#     obj.create_columns_w_dict(key='fakultna nemocnica trnava')
#     df_search = pd.DataFrame(obj.df_all[obj.final_table_cols])
#     df_orig = pd.concat([df_orig, df_search], ignore_index=True)
#
#     # save tables
#     db.insert_table(table_name='priame_objednavky', df=obj.df_all, if_exists='append', index=False)
#     move_all_files(source_path=obj.hosp_path, destination_path=obj.hosp_path_hist)
#     logger.info(f'{obj.hosp} data load finished successfully')
#
# except DataNotAvailable as e:
#     logger.info(e.message)
#     pass
#
# except Exception:
#     logger.error(f"Error in hosp: {obj.hosp}")
#     logger.error(traceback.format_exc())
#     pass

# DFN Kosice  ###



def get_data(objednavatel:str, dict_cena:dict, dict_key:str, mail_domain_extension='.sk', last_update_str=None):
    try:
        global df_orig

        obj = PriameObjednavkyMail(objednavatel)
        logger.info(f'{obj.hosp} data load started')

        # get last insert date
        if pd.isnull(last_update_str):
            last_update_str = df_orig['insert_date'][df_orig['objednavatel'] == obj.hosp].max().strftime('%d.%m.%Y %I:%M %p')

        # filter mails and download attachments
        search_result = otl.find_message(path,
                                         "@SQL= (""urn:schemas:httpmail:fromemail"" LIKE '%" + obj.hosp + mail_domain_extension + "') AND (""urn:schemas:httpmail:datereceived"" > '" + last_update_str + "') ")
        otl.save_attachment(obj.hosp_path, search_result)

        obj.load()
        obj.clean_tables()
        obj.create_table(stand_column_names=stand_column_names)

        if obj.df_all.empty:
            raise DataNotAvailable(obj.hosp)

        logger.info(f'Number of rows loaded: {obj.df_all.shape[0]}')

        obj.df_all = str_col_replace(obj.df_all, 'cena', dict_cena)
        obj.df_all['cena'] = obj.df_all['cena'].astype(float)

        obj.df_all['datum'] = obj.df_all['datum'].str.strip()
        obj.df_all['datum'] = obj.df_all['datum'].apply(get_dates)

        obj.create_columns_w_dict(key=dict_key)
        df_search = pd.DataFrame(obj.df_all[obj.final_table_cols])
        df_orig = pd.concat([df_orig, df_search], ignore_index=True)

        # save tables
        db.insert_table(table_name='priame_objednavky', df=obj.df_all, if_exists='append', index=False)
        move_all_files(source_path=obj.hosp_path, destination_path=obj.hosp_path_hist)
        logger.info(f'{obj.hosp} data load finished successfully')

    except DataNotAvailable as e:
        logger.info(e.message)
        pass

    except Exception:
        logger.error(f"Error in hosp: {obj.hosp}")
        logger.error(traceback.format_exc())
        pass

get_data('fnnitra', {"[a-z]|'|\s|[\(\)]+": '', '[,|\.]-.*': '', '-,': '', ",": '.', '\.\.': '.', '/.*/*': '', '\..\.\+': '',
                 '/\..*': '', '.*:': ''}, 'fakultna nemocnica nitra', last_update_str='25.03.2023 12:00 AM' )
get_data('fnsppresov', {"[a-z]|'|\s|[\(\)]+": "", ',-': '', ',$': '', ',+': '.', '/.*': ''}, 'fakultna nemocnica s poliklinikou j a reimana presov')
get_data('fntt', {"[a-z]|'|\s|[\(\)]+": "", ',-': '', ',$': '', ',+': '.', '/.*': ''}, 'fakultna nemocnica trnava')



#get_data('dfnkosice', {"[a-z]|'|\s|[\(\)]+": "", ',-': '', ',$': '', ',+': '.', '/.*': ''}, 'detska fakultna nemocnica kosice', last_update_str='24.03.2023 12:00 AM')






func.save_df(df=df_orig, name=os.path.join(os.getcwd(), 'priame_objednavky_all.pkl'))






