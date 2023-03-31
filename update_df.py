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
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.common.exceptions import ElementNotVisibleException, TimeoutException, WebDriverException


logger = logging.getLogger(__name__)
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
    logger.info(f"Connected to database")
except Exception as e:
    logger.error(traceback.format_exc())
    sys.exit()

# create backup and load original df
shutil.copyfile('priame_objednavky_all.pkl', 'priame_objednavky_all_backup.pkl')
df_orig = func.load_df('priame_objednavky_all.pkl', path=os.getcwd())

def get_data(objednavatel:str, dict_cena:dict, dict_key:str, mail_domain_extension:str ='.sk', last_update_str=None):
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

        if objednavatel == 'fnspza':
            obj.df_all = fnspza_data_cleaning(obj.df_all)
        else:
            obj.df_all = str_col_replace(obj.df_all, 'cena', dict_cena)
            obj.df_all['cena'] = np.where(obj.df_all['cena'].str.match(r'\d*\.\d*\.\d*'),
                                           obj.df_all['cena'].str.replace('\.', '', 1, regex=True), obj.df_all['cena'])
            obj.df_all['cena'] = obj.df_all['cena'].astype(float)

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

get_data('fnspza', dict_cena={}, dict_key='fakultna nemocnica s poliklinikou zilina')

get_data('fnnitra',
         {"[a-z]|'|\s|[\(\)]+": '', '[,|\.]-.*': '', '-,': '', ",": '.', '\.\.': '.', '/.*/*': '', '\..\.\+': '',
          '/\..*': '', '.*:': ''}, dict_key='fakultna nemocnica nitra')

get_data('fnsppresov', {"[a-z]|'|\s|[\(\)]+": "", ',-': '', ',$': '', ',+': '.', '/.*': ''},
         dict_key='fakultna nemocnica s poliklinikou j a reimana presov')

get_data('fntt', {"[a-z]|'|\s|[\(\)]+": "", ',-': '', ',$': '', ',+': '.', '/.*': ''},
         dict_key='fakultna nemocnica trnava')

get_data('dfnkosice', {"[a-z]|'|\s|[\(\)]+": "", ',-': '', ',$': '', ',+': '.', '/.*': ''},
         dict_key='detska fakultna nemocnica kosice')

get_data('unb',
         {"[a-z]|'|\s|[\(\)]+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '_':'0'},
         dict_key='univerzitna nemocnica bratislava')

get_data('unlp',
         {"[a-z]|'|\s|[\(\)]|\++": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0'},
         dict_key='univerzitna nemocnica l pasteura kosice')

get_data('unm',
         {"[a-z]|'|\s|[\(\)]|\++": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0'},
         dict_key='univerzitna nemocnica martin')

get_data('dfnbb',
         {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0'},
         dict_key='detska fakultna nemocnica s poliklinikou banska bystrica')

# only new mail domain nousk.sk
get_data('nou',
         {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0'},
         dict_key='narodny onkologicky ustav', mail_domain_extension='sk.sk')

get_data('nusch',
         {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0',
          '=.*': '', '.+\d{2}:\d{2}:\d{2}$': '0'},
         dict_key='narodny ustav srdcovych a cievnych chorob as')
get_data('vou',
         {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0'},
         dict_key='vychodoslovensky onkologicky ustav as')

get_data('nudch',
         {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0'},
         dict_key='narodny ustav detskych chorob', mail_domain_extension='.eu')

get_data('suscch',
         {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0'},
         dict_key='stredoslovensky ustav srdcovych a cievnych chorob as', mail_domain_extension='.eu')

get_data('inmm',
         {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0'},
         dict_key='institut nuklearnej a molekularnej mediciny kosice')

get_data('nurch',
         {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0'},
         dict_key='narodny ustav reumatickych chorob piestany')

get_data('nemocnicapp',
         {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0', '\.+':'.'},
         dict_key='nemocnica poprad as')


# FNTN
logger.info('fntn data load started')
most_recent_date = df_orig['datum'][df_orig['objednavatel'] == 'fntn'].max()
max_date_web = None

try:
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(chromedriver_path2, options=options)
    driver.get(dict_all['fakultna nemocnica trencin']['objednavky_link'])

    table_html = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH,
                                                                                   "//div[contains(@id, 'content')]//table"))).get_attribute(
        "outerHTML")
    table_pd = pd.read_html(table_html)[0]
    table_pd['datum'] = table_pd['Dátum vyhotovenia'].apply(get_dates)
    max_date_web = table_pd['datum'].max()

except Exception as e:
    logger.error(traceback.format_exc())
    driver.quit()
    pass

if max_date_web is not None and most_recent_date <= max_date_web:
    while most_recent_date <= max_date_web:
        try:
            WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH,
                                                                       "//div[contains(@class, 'next')]//a[contains(text(), 'Nasled')]"))).click()
            sleep(4)
            next_table_lst = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//div[contains(@id, 'content')]//table"))).get_attribute(
                "outerHTML")
            next_table = pd.read_html(next_table_lst)[0]
            next_table['datum'] = next_table['Dátum vyhotovenia'].apply(get_dates)
            table_pd = pd.concat([table_pd, next_table], ignore_index=True)
            max_date_web = next_table['datum'].max()
            sleep(3)
        except:
            logger.error(traceback.format_exc())
            driver.quit()
            print('Retrieving data failed')
            break
    driver.quit()
    logger.info('Data were retrieved')

    table_pd.drop(table_pd[table_pd['datum'] < most_recent_date].index, inplace=True)

    table_pd = func.clean_str_cols(table_pd)
    table_pd = clean_str_col_names(table_pd)

    table_pd['file'] = 'web'
    table_pd['insert_date'] = datetime.now()
    table_pd['dodavatel_ico'] = table_pd['ico dodavatela'].str.replace('\..*', '', regex=True)
    table_pd['cena'] = table_pd['cena s dph (EUR)'].astype(float)
    table_pd['cena_s_dph'] = 'ano'
    table_pd['objednavka_predmet'] = table_pd['popis']
    table_pd.rename(columns={'nazov dodavatela': 'dodavatel_nazov', 'cislo objednavky': 'objednavka_cislo'},
                    inplace=True)

    fntn = PriameObjednavkyMail('fntn')
    fntn.df_all = table_pd
    fntn.popis_list = ['popis', 'objednavka_cislo', 'cislo zmluvy', 'schvalil', 'cena_s_dph']
    fntn.create_columns_w_dict(key='fakultna nemocnica trencin')
    fntn.df_all.drop(
        columns=['schvalil', 'mesto dodavatela', 'psc dodavatela', 'adresa dodavatela', 'datum vyhotovenia',
                 'cislo zmluvy', 'cena s dph (EUR)', 'ico dodavatela'], inplace=True)

    fntn_search = pd.DataFrame(fntn.df_all[fntn.final_table_cols])

    df_orig = pd.concat([df_orig, fntn_search], ignore_index=True)
    df_orig.drop_duplicates(inplace=True)

    db.insert_table(table_name='priame_objednavky', df=df_orig[
        (df_orig['objednavatel'] == 'fntn') & (df_orig['insert_date'].dt.date == datetime.now().date())],
                    if_exists='append', index=False)


func.save_df(df=df_orig, name=os.path.join(os.getcwd(), 'priame_objednavky_all.pkl'))






