import win32com.client
import os
import pandas as pd
from config import *
import functions_ZK as func2
# import mysql.connector as pyo
from mysql_config import objednavky_db_connection, objednavky_db_connection_cloud
import copy
from sqlalchemy import create_engine, text
import pymysql
import math
import re
import traceback
import logging
logger = logging.getLogger(__name__)

class PriameObjednavkyMail:
    def __init__(self, hosp):
        self.hosp = hosp
        self.hosp_path = data_path + self.hosp + "\\"
        if not os.path.exists(self.hosp_path):
            os.mkdir(self.hosp_path)
        self.hosp_path_hist = historical_data_path + self.hosp + "\\"
        if not os.path.exists(self.hosp_path_hist):
            os.mkdir(self.hosp_path_hist)
        self.final_table_cols = ['objednavatel', 'cena', 'datum', 'dodavatel', 'popis', 'insert_date', 'file', 'link']
        self.all_tables_list = []
        self.all_tables_list_cleaned = []
        self.popis_list = ['objednavka_predmet', 'kategoria', 'objednavka_cislo', 'zdroj_financovania', 'balenie',
                      'sukl_kod', 'mnozstvo', 'poznamka', 'odkaz_na_zmluvu', 'pocet_oslovenych', 'cena_s_dph']
        self.dodavatel_list = ['dodavatel_nazov', 'dodavatel_ico']
        self.all_columns_names = []
        self.df_all = None


    def load(self):
        '''
        The method retrieves data files from the directory specified by the 'hosp_path' parameter.
        '''
        self.all_tables_list = func2.load_files(self.hosp_path)
        self.all_tables_list_cleaned = copy.deepcopy(self.all_tables_list)

    def clean_tables(self):
        '''
        The method cleans the tables stored in "all_tables_list_cleaned" list.
        '''
        for i in range(len(self.all_tables_list_cleaned)):
            doc = self.all_tables_list_cleaned[i][-1]
            for key, value in doc.items():
                doc[key] = func.clean_str_cols(doc[key])
                doc[key] = func2.clean_str_col_names(doc[key])
                for col in doc[key].columns.values:
                    doc[key].drop(doc[key][(doc[key][col].astype(str).str.match(clean_table_regex) == True)].index,
                                  inplace=True)
                doc[key] = doc[key].dropna(axis=1, thresh=math.ceil(doc[key].shape[0]*0.1))
                doc[key] = doc[key].dropna(thresh=2).reset_index(drop=True)
                if ('unnamed' in '|'.join(map(str, doc[key].columns))) or (pd.isna(doc[key].columns).any()):
                    doc[key] = doc[key].dropna(thresh=int(len(doc[key].columns) / 3)).reset_index(drop=True)
                    doc[key] = doc[key].dropna(axis=1, thresh=math.ceil(doc[key].shape[0]*0.1))
                    if not doc[key].empty:
                        if ('vystaveni objednavok' in '|'.join(map(str, doc[key].iloc[0]))): # remove 'Informacie o vystavení objednávok'
                            doc[key] = doc[key].drop(doc[key].index[0])
                        doc[key].columns = doc[key].iloc[0]
                        doc[key] = doc[key].drop(doc[key].index[0])
                if not doc[key].empty:
                    doc[key] = func2.clean_str_col_names(doc[key])
                    self.all_tables_list_cleaned[i].append(doc[key])

    def data_check(self):
        '''
        The method verifies which column names exist in tables stored in the "all_tables_list_cleaned" list.
        '''
        f = open('output.txt', 'w')
        for i in range(len(self.all_tables_list_cleaned)):
            lst = []
            lst.append(self.all_tables_list_cleaned[i][0])
            for j in range(2, len(self.all_tables_list_cleaned[i])):
                lst.append(self.all_tables_list_cleaned[i][j].columns.values)
                f.write(f"{lst}\n")
                for k in range(len(self.all_tables_list_cleaned[i][j].columns.values)):
                    self.all_columns_names.append(self.all_tables_list_cleaned[i][j].columns.values[k])
        f.close()
        print(set(self.all_columns_names))


    def create_table(self, stand_column_names):
        '''
        The method creates pandas data frame containing data stored in "all_tables_list_cleaned".
        :param stand_column_names: the dictionary used for column names standardization
        '''
        df = func2.create_table(self.all_tables_list_cleaned, stand_column_names)
        self.df_all = df.drop_duplicates()
        self.df_all.drop(self.df_all[pd.isna(self.df_all['objednavka_predmet']) & pd.isna(self.df_all['cena']) & pd.isna(
            self.df_all['datum'])].index, axis=0, inplace=True)

    def create_columns_w_dict(self, key: str):
        '''
        The method generates two columns named 'popis' and 'dodavatel' that store a dictionary containing
        all the available data about the order.
        :param key: hospital name from dict_all
        '''
        self.df_all['objednavatel'] = self.hosp
        self.df_all['link'] = dict_all[key]['zverejnovanie_objednavok_faktur_rozne']
        self.df_all['popis'] = self.df_all[self.popis_list].T.apply(lambda x: x.dropna().to_dict())
        self.df_all['dodavatel'] = self.df_all[self.dodavatel_list].T.apply(lambda x: x.dropna().to_dict())

        dict_cols = self.df_all.columns[self.df_all.applymap(lambda x: isinstance(x, dict)).any()]
        self.df_all = self.df_all.apply(lambda x: x.astype(str) if x.name in dict_cols else x)

    def save_tables(self, table, path: str = search_data_path):
        '''
        The method saves given table as xlsx file and pkl file.
        :param table: table name
        :param path: directory path where files will be saved
        '''
        with pd.ExcelWriter(os.path.join(path + self.hosp + '.xlsx'), engine='xlsxwriter',
                            engine_kwargs={'options': {'strings_to_urls': False}}) as writer:
            table.to_excel(writer)
        func.save_df(df=table, name=os.path.join(path + self.hosp + '.pkl'))


class ObjednavkyDB:
    def __init__(self, db_connection):
        self.db_connection = db_connection
        # if self.db_connection['host'] == '127.0.0.1':
        #     self.con = pyo.connect(**db_connection)
        #     self.cursor = self.con.cursor()
        #     self.engine = create_engine(
        #         f"mysql+mysqlconnector://{db_connection['user']}:{db_connection['password']}@{db_connection['host']}:{db_connection['port']}/{db_connection['database']}",
        #         echo=False)
        if self.db_connection['host'] == 'aws.connect.psdb.cloud':
            self.engine = create_engine(
                f"mysql+pymysql://{db_connection['user']}:{db_connection['password']}@{db_connection['host']}/{db_connection['database']}?charset=utf8mb4",
                connect_args={'ssl': {'ssl_ca': '/etc/ssl/cert.pem'}})
            self.con = self.engine.connect()

    def __del__(self):
        self.con.close()

    def fetch_records(self, query: str):
        '''
        Function executes given select statement and fetches the result into pandas data frame.
        :param query: select query
        '''
        return pd.read_sql(query, con=self.con)


    def insert_table(self, table_name: str, df: pd.DataFrame, if_exists: str = 'append', index: bool = False, **kwargs):
        '''
        :param table_name: table in database in which we want to insert data
        :param df: table with data
        :param if_exists: 'append' or 'replace'
        :param index: True or False
        '''
        dict_cols = df.columns[df.applymap(lambda x: isinstance(x, dict)).any()]
        df = df.apply(lambda x: x.astype(str) if x.name in dict_cols else x)
        try:
            df.to_sql(name=table_name, con=self.engine, if_exists='append', index=False, **kwargs) # works for localhost as well as cloud
        except Exception as e:
            logger.error(traceback.format_exc())

    def update(self, query: str, values: list):
        '''
        :param query: update query to be executed on the database
        :param values: values to be used in the query
        '''
        try:
            self.cursor.execute(query, values)
            self.con.commit()
        except Exception as e:
            logger.error(traceback.format_exc())

class OutlookTools:
    def __init__(self, object):
        self.obj = object
        self.n_folders = self.obj.Folders.Count

    def show_all_folders(self):
        '''
        Function prints all folders and subfolders within given outlook account
        '''
        for i in range(0, self.n_folders):
            print(f'Folder: [{i}] {self.obj.Folders[i].Name}')
            n_subfolders = self.obj.Folders[i].Folders.Count
            for j in range(n_subfolders):
                print(f'    Subfolder: [{j}] {self.obj.Folders[i].Folders[j].Name}')
                if self.obj.Folders[i].Folders[j].Folders.Count != 0:
                    for k in range(self.obj.Folders[i].Folders[j].Folders.Count):
                        print(f'        Subfolder: [{k}] {self.obj.Folders[i].Folders[j].Folders[k].Name}')

    def find_message(self, folder_path: str, condition: str):
        '''
        :param folder_path: example - outlook.Folders['zuzana.knapekova@health.gov.sk'].Folders['Doručená pošta']
        :param condition: possible filters to use: subject, sender, to, body, receivedtime etc.
        :return: item object with filtered messages
        '''
        messages_all = folder_path.Items
        return messages_all.Restrict(condition)


    def save_attachment(self, output_path, messages):
        '''
        :param output_path: folder for saving attachments
        :param messages: item object containing at least one message
        '''
        for message in messages:
            ts = pd.Timestamp(message.senton).strftime('%d_%m_%Y_%I_%M_%S')
            for attachment in message.Attachments:
                try:
                    if re.match(r'.*\.((png)|(jpg))', attachment.FileName):
                        continue
                    else:
                        attachment.SaveASFile(os.path.join(output_path, ts+'_'+attachment.FileName))
                    print(f"attachment {attachment.FileName} from {message.Sender} saved")
                except Exception as e:
                    print("error when saving the attachment:" + str(e))
        print('All atachments were saved')


