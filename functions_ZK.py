import cv2
from time import sleep
import numpy as np
import pandas as pd
import pytesseract
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.common.exceptions import ElementNotVisibleException, TimeoutException
import re
from config import *
from datetime import datetime
from unidecode import unidecode
import ezodf
import functionss as func
import os
import shutil

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
pd.options.mode.chained_assignment = None  # default='warn'

def update_dict(dict):
    dict['centrum pre liecbu drogovych zavislosti bratislava'][
        'objednavky_faktury_link'] = 'https://cpldz.sk/wp-content/uploads/Fakturyaobjednavky_december.xlsx'

    dict['centrum pre liecbu drogovych zavislosti kosice'][
        'objednavky_faktury_link'] = 'https://ba5c113364.clvaw-cdnwnd.com/a410a7ba52b55860f90d7278a3fc9ce5/200000168-2c40a2c40c/Objedn%C3%A1vky%20tovarov%2C%20slu%C5%BEieb%20a%20pr%C3%A1c%20a%20Fakt%C3%BAry.xlsx?ph=ba5c113364'

    dict['detska fakultna nemocnica kosice'][
        'objednavky_faktury_link'] = 'https://docs.google.com/uc?id=1uNSfzKDWdQfcwNJv5vBNS3yTGkGKlOIsK9FxXpx_edY'

    dict['detska fakultna nemocnica s poliklinikou banska bystrica'][
        'objednavky_faktury_link'] = 'https://www.detskanemocnica.sk/sites/default/files/files/199/objednavky_od_2023.01.01_do_2023.02.17.pdf'
    dict['detska fakultna nemocnica s poliklinikou banska bystrica']['objednavky_faktury_link2022']='https://www.detskanemocnica.sk/sites/default/files/files/199/objednavky_od_2022.01.01_do_2022.12.30.pdf'
    dict['detska fakultna nemocnica s poliklinikou banska bystrica']['objednavky_faktury_link2021']='https://www.detskanemocnica.sk/sites/default/files/files/199/objednavky_od_2022.01.01_do_2022.12.30.pdf'


    dict['detska fakultna nemocnica s poliklinikou banska bystrica']['objednavky_faktury_file_ext'] = '.pdf'


    dict['detska psychiatricka liecebna n o hran'][
        'objednavky_faktury_link'] = 'https://zverejnovanie.mzsr.sk/ministerstvo-zdravotnictva-sr/objednavky/?export=csv&art_rok=2023'
    dict['detska psychiatricka liecebna n o hran']['objednavky_faktury_file_ext'] = '.csv'

    dict['fakultna nemocnica nitra']['objednavky_faktury_link'] = 'https://fnnitra.sk/objd/new/'
    dict['fakultna nemocnica s poliklinikou f d roosevelta banska bystrica'][
        'objednavky_faktury_link'] = 'https://www.fnspfdr.sk/objednavky/zverejnenie.php?akcia=vsetkyobjednavky_internet'
    dict['fakultna nemocnica s poliklinikou zilina']['objednavky_link'] = 'http://www.fnspza.sk/zm2019/objednavky'

    return dict


def preprocess_image(img, resize={'apply': True, 'scale_percent': 220}, gray_scale=True,
                     thresholding={'apply': True, 'threshold': 0},
                     denoise={'apply': True, 'h': 3, 'templateWindowSize': 7, 'searchWindowSize': 21},
                     sharpen={'apply': True}):
    '''

    :param img: input image
    :param resize: dictionary with boolean key apply and scale percent. If resizing or any other method is used, apply needs to be set to True.
    :param gray_scale: converting the image to gray scale
    :param thresholding: dictionary with boolean key apply and threshold.
    :param denoise: dictionary with boolean key apply and input parameters neccessary for function.
    :param sharpen: dictionary with boolean key apply
    :return: transformed image
    '''
    if resize['apply']:
        width = int(img.shape[1] * resize['scale_percent'] / 100)
        height = int(img.shape[0] * resize['scale_percent'] / 100)
        dim = (width, height)
        img = cv2.resize(img, dim, interpolation=cv2.INTER_AREA)
    if gray_scale:
        img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    if thresholding['apply']:
        img = cv2.threshold(img, thresholding['threshold'], 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
        #img = cv2.adaptiveThreshold(img, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
    if denoise['apply']:
        if gray_scale:
            img = cv2.fastNlMeansDenoising(src=img, h=denoise['h'], templateWindowSize=denoise['templateWindowSize'],
                                           searchWindowSize=denoise['searchWindowSize'])
        else:
            img = cv2.fastNlMeansDenoisingColored(src=img, h=denoise['h'],
                                                  templateWindowSize=denoise['templateWindowSize'],
                                                  searchWindowSize=denoise['searchWindowSize'])
    if sharpen['apply']:
        kernel = np.array([[0, -1, 0], [-1, 5, -1],
                           [0, -1, 0]])  # using sharpen kernel https://en.wikipedia.org/wiki/Kernel_(image_processing)
        img = cv2.filter2D(img, -1, kernel)
    return img



def FNsP_BB_objednavky(link:str, search_by: str, value: str, name: str):
    '''

    :link search_by: nazov_dodavatela, cislo_objednavky or ICO
    :param search_by: nazov_dodavatela, cislo_objednavky or ICO
    :param value: text value to be searched
    :return: list with status info and pandas dataframe containing scraped values
    '''

    # set up driver and initialize variables
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(chromedriver_path, options=options)
    driver.get(link)
    img_name = data_path + datetime.now().strftime("%d-%m-%Y") + name + "_image.png"

    # wait until page is loaded
    try:
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//div[contains(@class, 'btn_obj')]//input[contains(@name, 'vyhladanie')]")))
        print('The page was loaded.')
    except ElementNotVisibleException:
        print("The page was not loaded.")
        driver.quit()
        return ['fail', None]

    # make screenshot and read the verification code
    driver.save_screenshot(img_name)
    image = cv2.imread(img_name)
    result = preprocess_image(image)
    text = pytesseract.image_to_string(result, config='--psm 6')
    # TODO - check easyocr package(?) in case more issues occurr
    code = re.findall('\d{5}', text)
    print('Code read by OCR: ', code)

    # fill in date and text fields to obtain result table
    try:
        insert_code = driver.find_element(By.XPATH, "//div[@id='captcha']//input[contains(@name, 'vercode')]")
        insert_code.send_keys(code)
        sleep(2)
        # select date range -- all
        date_dropdown = driver.find_element(By.XPATH,
                                    "//td[contains(@class, 'td_obd')]//select[contains(@name, 'rok')]")

        select = Select(date_dropdown)
        select.select_by_value('-- všetko --')
        sleep(2)

        radio_buttons = driver.find_elements(By.XPATH, "//input[@type='radio']")
        for item in radio_buttons:
            if (search_by == 'ICO' and item.get_attribute('value') == 'radio_dodavico') or (
                    search_by == 'cislo_objednavky' and item.get_attribute('value') == 'radio_cisloobj') or (
                    search_by == 'nazov_dodavatela' and item.get_attribute('value') == 'radio_dodavmeno'):
                item.click()
        sleep(2)

        search_input = driver.find_element(By.XPATH,
                                           "//td[contains(@class, 'td_text')]//input[contains(@name, 'vyhladaj')]")
        search_input.send_keys(value)
        sleep(1)
        driver.find_element(By.XPATH,
                            "//div[contains(@class, 'btn_obj')]//input[contains(@name, 'vyhladanie')]").click()
    except:
        print('FNsP_BB_objednavky set up failed')
        driver.quit()
        return ['fail', None]

    # retrieve the initial table
    try:
        table_lst = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "example"))).get_attribute(
            "outerHTML")
        result_df = pd.read_html(table_lst)[0]
        # click Next
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable(
            (By.XPATH, "//div[contains(@id, 'example_paginate')]//a[contains(@class, 'paginate_enabled_next')]"))).click()
    except:
        print('FNsP_BB_objednavky retrieving data failed')
        driver.quit()
        return ['fail', None]

    while True:
        try:
            next_table_lst = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.ID, "example"))).get_attribute(
                    "outerHTML")
            next_table = pd.read_html(next_table_lst)[0]
            result_df = pd.concat([result_df, next_table])
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,
                                             "//div[contains(@id, 'example_paginate')]//a[contains(@class, 'paginate_enabled_next')]"))).click()
            sleep(3)
        except TimeoutException as ex:
            print('Data were retrieved')
            break
        except:
            print('FNsP_BB_objednavky retrieving data failed')
            driver.quit()
            return ['fail', None]

    driver.quit()
    return ['ok', result_df]

def create_standardized_table(key, table, cols, columns_to_insert):
    '''

    :param key: name of the hospital/institution used in dictionary
    :param table: dataframe containing data
    :param cols: list of columns to be used in the result table
    :param columns_to_insert: values from dictionary
    :return: pandas dataframe
    '''
    objednavky = pd.DataFrame(columns=cols)

    # insert scraped data
    for i in objednavky.columns.values:
        for j in range(len(table.columns.values)):
            if dict[key][i] == table.columns.values[j]:
                print(dict[key][i], table.columns.values[j])
                objednavky[i] = table[table.columns[j]]

    # insert data from dictionary
    for col_name in objednavky.columns.values:
        if col_name in columns_to_insert:
            objednavky[col_name]=dict[key][col_name]

    return objednavky

def clean_str_col_names(df):
    df.columns = [unidecode(str(x).lower().strip().replace('\n', '')) for x in df.columns]
    return df

# def clean_table(df):
#     df = df.dropna(axis=1, thresh=3)
#     filtered_values = list(filter(lambda v: re.match('^Unnamed.*', v), df.columns))
#     indices = [index for (index, item) in enumerate(df.columns.values) if item in filtered_values]
#     for i in range(len(indices)):
#         df.columns.values[indices[i]]=df.iloc[0][indices[i]]
#     return df

def load_files(data_path):
    '''
    :param data_path: path to folder
    :return: list with name of file and dictionary containing data from all sheets
    '''
    # load downloaded files
    loading_start = datetime.now()
    print('Loading start: ', datetime.now())
    all_tables = []
    for file_name in os.listdir(data_path):
        if file_name.split(sep='.')[-1] in ('pdf', 'png', 'jpeg', 'pkl'):
            continue
        elif file_name.split(sep='.')[-1] == 'ods':
            df = pd.read_excel(os.path.join(data_path, file_name), engine='odf', sheet_name=None)
        else:
            df = func.load_df(name=file_name, path=data_path, sheet_name=None)
        all_tables.append([file_name, df])
    print('Loading took: ', datetime.now()-loading_start)
    return all_tables

def clean_tables(input_list):
    for i in range(len(input_list)):
        doc = input_list[i][-1]
        for key, value in doc.items():
            doc[key] = doc[key].dropna(axis=1, thresh=3)
            doc[key] = doc[key].dropna(thresh=2).reset_index(drop=True)
            if ('Unnamed' in '|'.join(map(str, doc[key].columns))) or (pd.isna(doc[key].columns).any()):
                doc[key] = doc[key].dropna(thresh=int(len(doc[key].columns) / 3)).reset_index(drop=True)
                doc[key] = doc[key].dropna(axis=1, thresh=3)
                if not doc[key].empty:
                    doc[key].columns = doc[key].iloc[0]
                    doc[key] = doc[key].drop(doc[key].index[0])
            if not doc[key].empty:
                doc[key] = clean_str_col_names(doc[key])
                input_list[i].append(doc[key])
    return input_list

def create_table(list_of_tables, dictionary):
    final_table = pd.DataFrame(columns=list(dictionary.keys()))
    for i in range(len(list_of_tables)):  # all excel files
        for j in range(2, len(list_of_tables[i])):  # all sheets
            include = False
            for k in range(len(list_of_tables[i][j].columns.values)):  # all columns
                for key, value in dictionary.items():
                    if list_of_tables[i][j].columns.values[k] in value:
                        include = True
                        list_of_tables[i][j].columns = list_of_tables[i][j].columns.str.replace(
                            list_of_tables[i][j].columns.values[k], key, regex=False)
            if include:
                df = list_of_tables[i][j].reset_index(drop=True)
                df['file'] = list_of_tables[i][0]
                df['insert_date'] = datetime.now()
                cols = list(set(list_of_tables[i][j].columns.values).intersection(final_table.columns.values))+['file', 'insert_date']
                final_table = pd.concat([final_table, df[cols]], ignore_index=True)
    return final_table


def get_dates(date_string: str):
    # example: 2022-08-31 00:00:00
    if re.match(r'^20\d{2}-\d{2}-\d{2}.*', str(date_string)):
        date = date_string.split(' ')[0]
        return pd.Timestamp(year=int(date.split('-')[0]), month=int(date.split('-')[1]),
                            day=int(date.split('-')[2]))
    # example: 31.8.2022 or 31. 8. 2022
    elif re.match(r'^\d+\.\s*\d+\.\s*20\d{2}$', str(date_string)):
        date = date_string.strip()
        return pd.Timestamp(year=int(date.split('.')[2]), month=int(date.split('.')[1]),
                            day=int(date.split('.')[0]))
    # example: 31/8/22
    elif re.match(r'\d+/\d+/\d{2}', str(date_string)):
        return pd.Timestamp(year=int('20' + date_string.split('/')[2]), month=int(date_string.split('/')[1]),
                            day=int(date_string.split('/')[0]))
    # example: 05.09.2022-09.09.2022
    elif re.match(r'\d+\.\d+\.20\d{2}.*-.*\d+\.\d+\.20\d{2}.*', str(date_string)):
        date = date_string.split('-')[0].strip()
        return pd.Timestamp(year=int(date.split('.')[2]), month=int(date.split('.')[1]),
                            day=int(date.split('.')[0]))
    # example: 2.-6.10.
    elif re.match(r'^\d+\.-\d+\.\d+.*', str(date_string)):
        date = date_string.split('-')[1].strip()
        return pd.Timestamp(year=2017, month=int(date.split('.')[1]),
                            day=int(date.split('.')[0]))
    # example: 12.-16.10.2018
    elif re.match(r'^\d+.*-.*\d+\.\d+\..*20\d{2}', str(date_string)):
        date = date_string.split('-')[1].strip()
        return pd.Timestamp(year=int(date.split('.')[2]), month=int(date.split('.')[1]),
                            day=int(date.split('.')[0]))
    elif re.match(r'^\d+.*\s.*\d+\.\d+\..*20\d{2}', str(date_string)):
        date = date_string.split(' ')[1].strip()
        return pd.Timestamp(year=int(date.split('.')[2]), month=int(date.split('.')[1]),
                            day=int(date.split('.')[0]))
    else:
        return np.nan


def fnspza_data_cleaning(input_table):
    ### data cleaning ###

    fnspza_all2 = func.clean_str_cols(input_table)

    # predmet objednavky
    fnspza_all2['extr_mnozstvo'] = fnspza_all2['objednavka_predmet'].str.extract(r'(\s+\d+x$)')
    fnspza_all2['mnozstvo'] = np.where((pd.isna(fnspza_all2['mnozstvo'])) & (
            pd.isna(fnspza_all2['extr_mnozstvo']) == False), fnspza_all2['extr_mnozstvo'].str.strip(),
                                       fnspza_all2['mnozstvo'])
    fnspza_all2['objednavka_predmet'] = fnspza_all2['objednavka_predmet'].str.replace(r'\s+\d+x$', '', regex=True)
    fnspza_all2.drop(['extr_mnozstvo'], axis=1, inplace=True)

    # cena
    dict_cena = {'^mc\s*': '', '[,|\.]-.*': '', '[a-z]+\.[a-z]*\s*': '', "[a-z]|'|\s|-|[\(\)]+": '', ",,": '.',
                       ",": '.', ".*[:].*": '0',
                       "=.*": ''}
    for original, replacement in dict_cena.items():
        fnspza_all2['cena'] = fnspza_all2['cena'].replace(original, replacement, regex=True)

    fnspza_all2['cena'] = np.where(fnspza_all2['cena'].str.match(r'\d*\.\d*\.\d*'),
                                   fnspza_all2['cena'].str.replace('.', '', 1), fnspza_all2['cena'])
    fnspza_all2['cena'] = fnspza_all2['cena'].astype(float)
    print('price converted to float successfully')

    # rok objednavky - 4 outlier values 2000, 2048, 2033 and 2026
    fnspza_all2['rok_objednavky'] = fnspza_all2['datum'].str.extract(r'(20\d{2})')

    fnspza_all2['rok_objednavky_num'] = fnspza_all2['rok_objednavky'].apply(
        lambda x: pd.to_numeric(x) if (pd.isna(x) == False) else 0)

    # datum objednavky
    dict_datum_objednavky = {'210': ['201', True], '217': ['2017', True], '[a-z]': ['', True], '\(': ['', True],
                             '\)': ['', True], '\s+': [' ', True], '..': ['.', False], ': ': ['', False], '.,': ['.', False],
                             '5019': ['2019', True], ',': ['.', False], '201/8': ['2018', True], '20\.17': ['2017', True],
                             '209': ['2019', True], '19.12.202$': ['19.12.2022', True]
                             }
    for key, value in dict_datum_objednavky.items():
        fnspza_all2['datum'] = fnspza_all2['datum'].replace(key, value[0], regex=value[1])

    fnspza_all2['datum'] = fnspza_all2['datum'].str.strip()

    fnspza_all2['datum_adj'] = fnspza_all2['datum'].apply(get_dates)

    fnspza_all2['datum_adj'] = fnspza_all2.apply(
        lambda row: pd.Timestamp(year=row['rok_objednavky_num'], month=1, day=1) if (
                (pd.isna(row['rok_objednavky']) == False) & (pd.isnull(row['datum_adj']) == True)) else row[
            'datum_adj'], axis=1)
    print('date converted to timestamp successfully')

    # popis
    popis_list = ['objednavka_predmet', 'kategoria', 'objednavka_cislo', 'zdroj_financovania', 'balenie',
                  'sukl_kod', 'mnozstvo', 'poznamka', 'odkaz_na_zmluvu', 'pocet_oslovenych']
    dodavatel_list = ['dodavatel_nazov', 'dodavatel_ico']

    fnspza_all3=fnspza_all2.drop_duplicates()

    fnspza_all3['popis'] = fnspza_all3[popis_list].T.apply(lambda x: x.dropna().to_dict())
    fnspza_all3['dodavatel'] = fnspza_all3[dodavatel_list].T.apply(lambda x: x.dropna().to_dict())

    return fnspza_all3



def move_all_files(source_path, destination_path):
    for file_name in os.listdir(source_path):
        source = source_path + file_name
        destination = destination_path + file_name
        # move only files
        if os.path.isfile(source):
            shutil.move(source, destination)
            print('Moved:', file_name)
