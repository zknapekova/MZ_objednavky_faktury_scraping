import pandas as pd
from config import *
from functions_ZK import *
from schemas import OutlookTools, PriameObjednavkyMail, ObjednavkyDB
import win32com.client
import functionss as func
import shutil
from mysql_config import objednavky_db_connection, objednavky_db_connection_cloud
import logging
import sys
import traceback
from exceptions import DataNotAvailable
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException
from urllib.request import urlretrieve
import camelot


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
except Exception:
    logger.error(traceback.format_exc())
    sys.exit()

try:
    db = ObjednavkyDB(objednavky_db_connection)
    logger.info(f"Connected to local database")
    db_cloud = ObjednavkyDB(objednavky_db_connection_cloud)
    logger.info(f"Connected to cloud database")
except Exception:
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
        try:
            db.insert_table(table_name='priame_objednavky', df=obj.df_all)
            db_cloud.insert_table(table_name='priame_objednavky', df=obj.df_all)
            logger.info(f'{obj.hosp} data saved to database')
        except Exception:
            logger.error(traceback.format_exc())
            logger.info(f'{obj.hosp} data insert to database failed')

        move_all_files(source_path=obj.hosp_path, destination_path=obj.hosp_path_hist)
        logger.info(f'{obj.hosp} data load finished successfully')

    except DataNotAvailable as e:
        logger.info(e.message)

    except Exception:
        logger.error(f"Error in hosp: {obj.hosp}")
        logger.error(traceback.format_exc())


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

get_data('vusch',
         {"[a-z]|'|\s|[\(\)]|\+|\*+": "", ',-': '', ',$': '', ',+': '.', '/.*': '', '\.-': '', '-': '', '': '0'},
         dict_key='vychodoslovensky ustav srdcovych a cievnych chorob as')


class Locators:
    def __init__(self, hosp):
        if hosp == 'fntn':
            self.table = (By.XPATH, "//div[contains(@id, 'content')]//table")
            self.next_button = (By.XPATH, "//div[contains(@class, 'next')]//a[contains(text(), 'Nasled')]")
        elif hosp == 'donsp':
            self.table = (By.XPATH, "//div[contains(@class, 'responsive-table')]//table")
            self.next_button = (By.XPATH, "//div[contains(@class, 'responsive-table')]//table[contains(@class, 'container')]//tr[contains(@class, 'foot')]//a[contains(text(), '»')]")

class DataHandling:
    def __init__(self, hosp):
        if hosp == 'fntn':
            self.popis_list = ['objednavka_predmet', 'objednavka_cislo', 'cislo zmluvy', 'schvalil', 'cena_s_dph']
            self.dodavatel_list = 'default' # function will use default value in class PriameObjednavkyMail
            self.drop_list = ['ico dodavatela', 'schvalil', 'mesto dodavatela', 'psc dodavatela', 'adresa dodavatela',
                     'datum vyhotovenia', 'cislo zmluvy', 'cena s dph (EUR)']




class GetDataPage:
    def __init__(self, driver, url):
        self.driver = driver
        self.url = url

    def load_web_page(self):
        self.driver.get(self.url)

    def get_table(self, locator, wait_time=10):
        try:
            element = WebDriverWait(self.driver, wait_time).until(EC.visibility_of_element_located(locator)).get_attribute(
            "outerHTML")
            return element
        except NoSuchElementException:
            logger.error(f"Element with locator: {locator} was not found")
            logger.error(traceback.format_exc())

    def click_next_button(self, locator, wait_time=2):
        try:
            WebDriverWait(driver, wait_time).until(EC.element_to_be_clickable(locator)).click()
        except ElementClickInterceptedException:
            logger.error(f"Element with locator: {locator} was not clickable")
            logger.error(traceback.format_exc())


class BaseProcedure:
    def __init__(self, hosp: str, url: str, driver):
        self.driver = driver
        self.loc = Locators(hosp)
        self.get_data_page = GetDataPage(driver, url)
        self.obj = PriameObjednavkyMail(hosp)
        self.data_handling = DataHandling(hosp)

    def download_data(self, date_col: str, most_recent_date: datetime):
        self.most_recent_date=most_recent_date
        self.get_data_page.load_web_page()
        table_html = self.get_data_page.get_table(locator=self.loc.table)
        self.obj.df_all = pd.read_html(table_html)[0]
        self.obj.df_all['datum'] = self.obj.df_all[date_col].apply(get_dates)
        max_date_web = self.obj.df_all['datum'].max()

        if max_date_web is not None:
            while most_recent_date <= max_date_web:
                try:
                    self.get_data_page.click_next_button(locator=self.loc.next_button)
                    sleep(4)
                    next_table_lst = self.get_data_page.get_table(locator=self.loc.table)
                    next_table = pd.read_html(next_table_lst)[0]
                    next_table['datum'] = next_table[date_col].apply(get_dates)
                    self.obj.df_all = pd.concat([self.obj.df_all, next_table], ignore_index=True)
                    max_date_web = next_table['datum'].max()
                    sleep(3)
                except Exception:
                    logger.error(traceback.format_exc())
                    driver.quit()
                    print('Retrieving data failed')
                    break
            driver.quit()
            self.obj.df_all.drop(self.obj.df_all[self.obj.df_all['datum'] < most_recent_date].index, inplace=True)


    def clean_data(self, function, key):
        self.obj.df_all = function(self.obj.df_all)

        if self.data_handling.popis_list and self.data_handling.popis_list != 'default':
            self.obj.popis_list = self.data_handling.popis_list

        if self.data_handling.dodavatel_list and self.data_handling.dodavatel_list != 'default':
            self.obj.dodavatel_list = self.data_handling.dodavatel_list

        self.obj.create_columns_w_dict(key=key)
        if self.data_handling.dodavatel_list:
            self.obj.df_all.drop(columns=self.data_handling.drop_list, inplace=True)

        # remove duplicates

        df_db = db.fetch_records(
            "select * from objednavky.priame_objednavky where objednavatel='"+ self.obj.hosp +"' and datum = '{}'; ".format(
                self.most_recent_date.date().strftime('%Y-%m-%d')))
        df_concat = (pd.concat([self.obj.df_all,
                        df_db[self.obj.df_all.columns]]).drop_duplicates(['popis', 'cena', 'datum', 'dodavatel'], keep=False))

        logger.info('{} rows retrieved'.format(df_concat.shape[0]))
        self.df_search = pd.DataFrame(df_concat[self.obj.final_table_cols])
        # df_orig = pd.concat([df_orig, fntn_search], ignore_index=True)
        # db.insert_table(table_name='priame_objednavky', df=df_concat)
        # db_cloud.insert_table(table_name='priame_objednavky', df=df_concat)
        # logger.info('Data saved to database')


# FNTN
logger.info('fntn data load started')

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(chromedriver_path2, options=options)
fntn = BaseProcedure('fntn', url=dict_all['fakultna nemocnica trencin']['objednavky_link'], driver=driver)

fntn.download_data(date_col='Dátum vyhotovenia', most_recent_date=df_orig['datum'][df_orig['objednavatel'] == 'fntn'].max())
fntn.clean_data(fntn_data_cleaning, key='fakultna nemocnica trencin')




# FNTN
logger.info('fntn data load started')
most_recent_date = df_orig['datum'][df_orig['objednavatel'] == 'fntn'].max()
max_date_web = None

try:
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(chromedriver_path2, options=options)
    driver.get(dict_all['fakultna nemocnica trencin']['objednavky_link'])

    table_html = WebDriverWait(driver, 10).until(EC.visibility_of_element_located(loc.fntn_table)).get_attribute(
        "outerHTML")
    table_pd = pd.read_html(table_html)[0]
    table_pd['datum'] = table_pd['Dátum vyhotovenia'].apply(get_dates)
    max_date_web = table_pd['datum'].max()
except Exception as e:
    logger.error(traceback.format_exc())
    driver.quit()


if max_date_web is not None:
    while most_recent_date <= max_date_web:
        try:
            WebDriverWait(driver, 2).until(EC.element_to_be_clickable(loc.fntn_next_button)).click()
            sleep(4)
            next_table_lst = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located(loc.fntn_table)).get_attribute(
                "outerHTML")
            next_table = pd.read_html(next_table_lst)[0]
            next_table['datum'] = next_table['Dátum vyhotovenia'].apply(get_dates)
            table_pd = pd.concat([table_pd, next_table], ignore_index=True)
            max_date_web = next_table['datum'].max()
            sleep(3)
        except Exception as e:
            logger.error(traceback.format_exc())
            driver.quit()
            print('Retrieving further data failed')
            break
    driver.quit()
    logger.info('Data cleaning started')
    try:
        table_pd.drop(table_pd[table_pd['datum'] < most_recent_date].index, inplace=True)

        fntn = PriameObjednavkyMail('fntn')
        fntn.df_all = fntn_data_cleaning(table_pd)
        fntn.popis_list = ['objednavka_predmet', 'objednavka_cislo', 'cislo zmluvy', 'schvalil', 'cena_s_dph']
        fntn.create_columns_w_dict(key='fakultna nemocnica trencin')
        fntn.df_all.drop(
            columns=['ico dodavatela', 'schvalil', 'mesto dodavatela', 'psc dodavatela', 'adresa dodavatela',
                     'datum vyhotovenia', 'cislo zmluvy', 'cena s dph (EUR)'], inplace=True)

        # remove duplicates

        df_db = db.fetch_records(
            "select * from objednavky.priame_objednavky where objednavatel='fntn' and datum = '{}'; ".format(
                most_recent_date.date().strftime('%Y-%m-%d')))
        df_concat = (pd.concat([fntn.df_all,
                        df_db[fntn.df_all.columns]]).drop_duplicates(['popis', 'cena', 'datum', 'dodavatel'], keep=False))

        logger.info('{} rows retrieved'.format(df_concat.shape[0]))
        fntn_search = pd.DataFrame(df_concat[fntn.final_table_cols])
        df_orig = pd.concat([df_orig, fntn_search], ignore_index=True)
        db.insert_table(table_name='priame_objednavky', df=df_concat)
        db_cloud.insert_table(table_name='priame_objednavky', df=df_concat)
        logger.info('Data saved to database')
    except Exception as e:
        logger.error(traceback.format_exc())
        logger.error('Data insert failed for fntn')
        pass

# DONSP
logger.info('donsp data load started')
most_recent_date = df_orig['datum'][df_orig['objednavatel'] == 'donsp'].max()
max_date_web = None
donsp = PriameObjednavkyMail('donsp')

donsp.df_all = donsp_data_download(webpage=donsp_webpages['objednavky_2021_2023'], most_recent_date=most_recent_date, options=options)

donsp.popis_list = ['objednavka_predmet', 'objednavka_cislo', 'cena_s_dph']
donsp.create_columns_w_dict(key='dolnooravska nemocnica s poliklinikou mudr l nadasi jegeho dolny kubin')

if not donsp.df_all.empty:
    try:
        df_db_donsp = db.fetch_records(
            "select * from objednavky.priame_objednavky where objednavatel='donsp' and datum = '{}'; ".format(
                most_recent_date.date().strftime('%Y-%m-%d')))
        df_concat = (pd.concat([donsp.df_all,
                                df_db_donsp[donsp.df_all.columns]]).drop_duplicates(['popis', 'cena', 'datum', 'dodavatel'],
                                                                                    keep=False))
        if df_concat.empty:
            raise DataNotAvailable('nsptrstena')

        donsp_search = pd.DataFrame(df_concat[donsp.final_table_cols])

        df_orig = pd.concat([df_orig, donsp_search], ignore_index=True)
        db.insert_table(table_name='priame_objednavky', df=df_concat)
        db_cloud.insert_table(table_name='priame_objednavky', df=df_concat)
        logger.info('Data saved to database')
    except DataNotAvailable as exp:
        logger.info(exp.message)
    except Exception as e:
        logger.error(traceback.format_exc())
        logger.error('Inserting data to db failed for nsptrstena')


# NSP Trstena

# the script downloads pdf file from hopsital webpage using urlretrieve and extracts data from it
logger.info('nsptrstena data load started')
nsptrstena = PriameObjednavkyMail('nsptrstena')

nsptrstena.df_all = nsptrstena_data_handling(
    weblink=dict_all['hornooravska nemocnica s poliklinikou trstena_2023']['zverejnovanie_objednavok_faktur_rozne'],
    hosp_object=nsptrstena)

if not nsptrstena.df_all.empty:
    try:
        df_db_nsptrstena = db.fetch_records(
            "select * from objednavky.priame_objednavky where objednavatel='nsptrstena' and file like '%" + str(datetime.now().year) + "%'")
        df_concat = (pd.concat([nsptrstena.df_all,
                                df_db_nsptrstena[nsptrstena.df_all.columns]]).drop_duplicates(
            ['popis', 'cena', 'datum', 'dodavatel'], keep=False))
        if df_concat.empty:
            raise DataNotAvailable('nsptrstena')
        nsptrstena_search = pd.DataFrame(df_concat[nsptrstena.final_table_cols])
        df_orig = pd.concat([df_orig, nsptrstena_search], ignore_index=True)

        db.insert_table(table_name='priame_objednavky', df=df_concat)
        db_cloud.insert_table(table_name='priame_objednavky', df=df_concat)

    except DataNotAvailable as exp:
        logger.info(exp.message)
    except Exception as e:
        logger.error(traceback.format_exc())
        logger.error('Inserting data to db failed for nsptrstena')







func.save_df(df=df_orig, name=os.path.join(os.getcwd(), 'priame_objednavky_all.pkl'))





