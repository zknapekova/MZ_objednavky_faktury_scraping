import win32com.client  # pywin32 package needs to be installed
import os
import pandas as pd
from config import *
from functions_ZK import *
import mysql.connector as pyo
from mysql_config import *
import copy


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
                      'sukl_kod', 'mnozstvo', 'poznamka', 'odkaz_na_zmluvu', 'pocet_oslovenych']
        self.dodavatel_list = ['dodavatel_nazov', 'dodavatel_ico']
        self.all_columns_names = []
        self.df_all = None


    def load(self):
        self.all_tables_list = load_files(self.hosp_path)
        self.all_tables_list_cleaned = copy.deepcopy(self.all_tables_list)

    def clean_tables(self):
        for i in range(len(self.all_tables_list_cleaned)):
            doc = self.all_tables_list_cleaned[i][-1]
            for key, value in doc.items():
                doc[key] = doc[key].dropna(axis=1, thresh=3)
                doc[key] = doc[key].dropna(thresh=2).reset_index(drop=True)
                doc[key] = func.clean_str_cols(doc[key])
                doc[key] = clean_str_col_names(doc[key])
                for col in doc[key].columns.values:
                    doc[key].drop(doc[key][(doc[key][col].astype(str).str.match('(vypracov.*)|(schvalil.*)') == True)].index,
                                  inplace=True)
                if ('unnamed' in '|'.join(map(str, doc[key].columns))) or (pd.isna(doc[key].columns).any()):
                    doc[key] = doc[key].dropna(thresh=int(len(doc[key].columns) / 3)).reset_index(drop=True)
                    doc[key] = doc[key].dropna(axis=1, thresh=3)
                    if not doc[key].empty:
                        if ('vystaveni objednavok' in '|'.join(map(str, doc[key].iloc[0]))):
                            doc[key] = doc[key].drop(doc[key].index[0])
                        doc[key].columns = doc[key].iloc[0]
                        doc[key] = doc[key].drop(doc[key].index[0])
                if not doc[key].empty:
                    doc[key] = clean_str_col_names(doc[key])
                    self.all_tables_list_cleaned[i].append(doc[key])

    def data_check(self):
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
        print('File output.txt was saved.')

    def create_table(self, stand_column_names):
        df = create_table(self.all_tables_list_cleaned, stand_column_names)
        self.df_all = df.drop_duplicates()

    def create_columns_w_dict(self, key):
        self.df_all['objednavatel'] = self.hosp
        self.df_all['link'] = dict[key]['zverejnovanie_objednavok_faktur_rozne']
        self.df_all['popis'] = self.df_all[self.popis_list].T.apply(lambda x: x.dropna().to_dict())
        self.df_all['dodavatel'] = self.df_all[self.dodavatel_list].T.apply(lambda x: x.dropna().to_dict())

    def save_tables(self, table, path=search_data_path):
        with pd.ExcelWriter(os.path.join(path + self.hosp + '.xlsx'), engine='xlsxwriter',
                            engine_kwargs={'options': {'strings_to_urls': False}}) as writer:
            table.to_excel(writer)

        func.save_df(df=table, name=os.path.join(path + self.hosp + '.pkl'))

class ObjednavkyDB():
    def __int__(self):
        self.con = pyo.connect(**db_connection)
        self.cursor = self.con.cursor

    def __del__(self):
        self.con.close()

    def fetch_records(self, query):
        self.cursor.execute(query)
        rows = self.cursor.fetchall()
        return rows

    def insert(self, query: str, values: list):
        # example of query: insert into table(col1, ...) values(%s, %s, %s)
        # values =[title, author, isbn]
        self.cursor.execute(query, values)
        self.con.commit()

    def update(self, query: str, values: list):
        self.cursor.execute(query, values)
        self.con.commit()



class OutlookTools:
    def __init__(self, object):
        self.obj = object
        self.n_folders = self.obj.Folders.Count

    def show_all_folders(self):
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
        :return:
        '''
        for message in messages:
            ts = pd.Timestamp(message.senton).strftime('%d_%m_%Y_%I_%M_%S')
            for attachment in message.Attachments:
                try:
                    if re.match(r'.*\.png', attachment.FileName):
                        continue
                    else:
                        attachment.SaveASFile(os.path.join(output_path, ts+'_'+attachment.FileName))
                    print(f"attachment {attachment.FileName} from {message.Sender} saved")
                except Exception as e:
                    print("error when saving the attachment:" + str(e))
        print('All atachments were saved')

# Example of use:
#path = outlook.Folders['zuzana.knapekova@health.gov.sk'].Folders['Doručená pošta']
#OT = OutlookTools(outlook)
#result = OT.find_message(path, "[SenderName] = 'Gajdošová Denisa'")
#OutlookTools(outlook).save_attachement(os.getcwd(), result)

# Useful links:
# another way of getting inbox folder - https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
# all available properties for message - https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.mailitem
# restriction guide - https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook._items.restrict
# advanced filtering - https://learn.microsoft.com/en-us/office/vba/outlook/how-to/search-and-filter/filtering-items-using-query-keywords

