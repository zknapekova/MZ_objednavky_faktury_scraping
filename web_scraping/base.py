from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException
import traceback
import datetime
from schemas import PriameObjednavkyMail
import pandas as pd
from web_scraping.locators import Locators
from time import sleep
from functions_ZK import get_dates
from exceptions import DataNotAvailable
import logging

logger = logging.getLogger(__name__)


class DataHandling:
    def __init__(self, hosp):
        if hosp == 'fntn':
            self.popis_list = ['objednavka_predmet', 'objednavka_cislo', 'cislo zmluvy', 'schvalil', 'cena_s_dph']
            self.dodavatel_list = 'default'
            self.drop_list = ['ico dodavatela', 'schvalil', 'mesto dodavatela', 'psc dodavatela', 'adresa dodavatela',
                              'datum vyhotovenia', 'cislo zmluvy', 'cena s dph (EUR)']
        elif hosp == 'donsp':
            self.popis_list = ['objednavka_predmet', 'objednavka_cislo', 'cena_s_dph']
            self.dodavatel_list = 'default'
            self.drop_list = ['Meno schvaľujúceho']


class GetDataPage:
    def __init__(self, driver, url):
        self.driver = driver
        self.url = url

    def load_web_page(self):
        self.driver.get(self.url)

    def get_table(self, locator, wait_time=10):
        try:
            element = WebDriverWait(self.driver, wait_time).until(
                EC.visibility_of_element_located(locator)).get_attribute(
                "outerHTML")
            return element
        except NoSuchElementException:
            logger.error(f"Element with locator: {locator} was not found")
            logger.error(traceback.format_exc())

    def click_next_button(self, locator, wait_time=2):
        try:
            WebDriverWait(self.driver, wait_time).until(EC.element_to_be_clickable(locator)).click()
        except ElementClickInterceptedException:
            logger.error(f"Element with locator: {locator} was not clickable")
            logger.error(traceback.format_exc())


class Base:
    def __init__(self, hosp: str, url: str, driver):
        self.driver = driver
        self.loc = Locators(hosp)
        self.get_data_page = GetDataPage(driver, url)
        self.obj = PriameObjednavkyMail(hosp)
        self.data_handling = DataHandling(hosp)

    def download_data(self, date_col: str, most_recent_date: datetime):
        self.most_recent_date = most_recent_date
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
                    self.driver.quit()
                    break
            self.driver.quit()
            self.obj.df_all.drop(self.obj.df_all[self.obj.df_all['datum'] < most_recent_date].index, inplace=True)

    def clean_data(self, function, key, db_con_cloud, **kwargs):
        try:
            self.obj.df_all = function(self.obj.df_all, **kwargs)

            if self.data_handling.popis_list and self.data_handling.popis_list != 'default':
                self.obj.popis_list = self.data_handling.popis_list

            if self.data_handling.dodavatel_list and self.data_handling.dodavatel_list != 'default':
                self.obj.dodavatel_list = self.data_handling.dodavatel_list

            self.obj.create_columns_w_dict(key=key)
            if self.data_handling.drop_list:
                self.obj.df_all.drop(columns=self.data_handling.drop_list, inplace=True)
            self.obj.df_all = self.obj.df_all.T.drop_duplicates().T

            # remove duplicates
            df_db = db_con_cloud.fetch_records(
                "select * from priame_objednavky where objednavatel='" + self.obj.hosp + "' and datum = '{}'; ".format(
                    self.most_recent_date.date().strftime('%Y-%m-%d')))
            df_concat = (pd.concat([self.obj.df_all,
                                    df_db[self.obj.df_all.columns]]).drop_duplicates(
                ['popis', 'cena', 'datum', 'dodavatel'], keep=False))
            logger.info('{} rows retrieved'.format(df_concat.shape[0]))

            if df_concat.empty:
                raise DataNotAvailable(self.obj.hosp)

            self.df_search = pd.DataFrame(df_concat[self.obj.final_table_cols])
            # df_orig = pd.concat([df_orig, fntn_search], ignore_index=True)
            # db_con_cloud.insert_table(table_name='priame_objednavky', df=df_concat)
            # logger.info('Data saved to database')
        except DataNotAvailable as e:
            logger.info(e.message)
        except Exception:
            logger.error(traceback.format_exc())
            logger.error(f'Inserting data to db failed for {self.obj.hosp}')
