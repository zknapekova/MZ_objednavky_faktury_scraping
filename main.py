import pandas as pd
import functionss as func
import os
from urllib.request import build_opener, install_opener, urlretrieve, urlopen
from datetime import datetime
#import gdown
import tabula
from selenium import webdriver
from time import sleep
import pytesseract
import cv2
import re

current_date_time = datetime.now()
print('Start:', current_date_time)

# set paths
source_path = 'C:\\Users\\knapekoz\\OneDrive - health.gov.sk\\\Zverejnovanie zmluv linky_subjektyCO.xlsx'
data_path = os.getcwd() + "\\data\\"
vo_path = "C:\\Users\\knapekoz\\health.gov.sk\\OSCM - Posudzovanie žiadostí\\"
chromedriver_path = os.path.join(vo_path, 'chromedriver_win32/chromedriver.exe')
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

if not os.path.exists(data_path):
    os.mkdir(data_path)

# load excel and create dictionary
df = func.load_df(source_path)
# clean data
df = func.clean_str_cols(df)
df['Nazov_full'] = df['Nazov_full'].replace(',|\.', '', regex=True)

dict = df.set_index('Nazov_full').T.to_dict('dict')
keysList = list(dict.keys())

# update
dict['centrum pre liecbu drogovych zavislosti bratislava'][
    'objednavky_faktury_link'] = 'https://cpldz.sk/wp-content/uploads/Fakturyaobjednavky_december.xlsx'

dict['centrum pre liecbu drogovych zavislosti kosice'][
    'objednavky_faktury_link'] = 'https://ba5c113364.clvaw-cdnwnd.com/a410a7ba52b55860f90d7278a3fc9ce5/200000168-2c40a2c40c/Objedn%C3%A1vky%20tovarov%2C%20slu%C5%BEieb%20a%20pr%C3%A1c%20a%20Fakt%C3%BAry.xlsx?ph=ba5c113364'

dict['detska fakultna nemocnica kosice'][
    'objednavky_faktury_link'] = 'https://docs.google.com/uc?id=1uNSfzKDWdQfcwNJv5vBNS3yTGkGKlOIsK9FxXpx_edY'

dict['detska fakultna nemocnica s poliklinikou banska bystrica'][
    'objednavky_faktury_link'] = 'https://www.detskanemocnica.sk/sites/default/files/files/199/objednavky_od_2023.01.01_do_2023.02.17.pdf'
dict['detska fakultna nemocnica s poliklinikou banska bystrica']['objednavky_faktury_file_ext'] = '.pdf'

dict['detska psychiatricka liecebna n o hran']['objednavky_faktury_link'] = 'https://zverejnovanie.mzsr.sk/ministerstvo-zdravotnictva-sr/objednavky/?export=csv&art_rok=2023'
dict['detska psychiatricka liecebna n o hran']['objednavky_faktury_file_ext'] = '.csv'

dict['fakultna nemocnica nitra']['objednavky_faktury_link'] = 'https://fnnitra.sk/objd/new/'
dict['fakultna nemocnica s poliklinikou f d roosevelta banska bystrica']['objednavky_faktury_link'] = 'https://www.fnspfdr.sk/objednavky/zverejnenie.php?akcia=vsetkyobjednavky_internet'


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

list_of_dfs = tabula.read_pdf(file_name, pages='all')
df_conc = pd.DataFrame(columns=list_of_dfs[0].columns)

for i in range(len(list_of_dfs)):
    df_conc = pd.concat([df_conc, list_of_dfs[i]])


# 5 - Detská psychiatrická liecebna n.o. Hráň
urlretrieve(dict['detska psychiatricka liecebna n o hran']['objednavky_faktury_link'],
            data_path + current_date_time.strftime("%d-%m-%Y") + str(keysList[5]).replace(" ", "_") +
            dict['detska psychiatricka liecebna n o hran']['objednavky_faktury_file_ext'])

# 6 - Fakultna nemocnica Nitra
table = pd.read_html(dict['fakultna nemocnica nitra']['objednavky_faktury_link'])[0]

# 7 - fakultna nemocnica s poliklinikou f d roosevelta banska bystrica
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(chromedriver_path, options=options)
driver.get(dict['fakultna nemocnica s poliklinikou f d roosevelta banska bystrica']['objednavky_faktury_link'])
driver.save_screenshot("image.png")

img = cv2.imread('image.png')

#preprocessing
def resizing_image(img, scale_percent:int):
    width = int(img.shape[1] * scale_percent / 100)
    height = int(img.shape[0] * scale_percent / 100)
    dim = (width, height)
    return cv2.resize(img, dim, interpolation=cv2.INTER_AREA)

def thresholding(image):
    return cv2.threshold(image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]


resized = resizing_image(img, scale_percent=220)
gray = cv2.cvtColor(resized, cv2.COLOR_BGR2GRAY)
thresh = thresholding(gray)

#cropped_image = thresh[0:900, 0:2500]
#cv2.imwrite('image_cropped.png', cropped_image)

text = pytesseract.image_to_string(thresh, config='--psm 6')

code = re.findall('\d{5}', text)[-1]
print(code)

