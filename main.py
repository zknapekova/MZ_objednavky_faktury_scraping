import functionss as func
import os
from urllib.request import build_opener, install_opener, urlretrieve
from datetime import datetime
import gdown


current_date_time = datetime.now()
print('Start:', current_date_time)

# set paths
source_path ='C:\\Users\\knapekoz\\OneDrive - health.gov.sk\\\Zverejnovanie zmluv linky_subjektyCO.xlsx'
data_path = os.getcwd() + "\\data\\"

if not os.path.exists(data_path):
    os.mkdir(data_path)

# load excel and create dictionary
df = func.load_df(source_path)
dict = df.set_index('Nazov_full').T.to_dict('dict')
keysList = list(dict.keys())

# update
dict['Centrum pre liečbu drogových závislostí Bratislava']['objednavky_faktury_link']='https://cpldz.sk/wp-content/uploads/Fakturyaobjednavky_december.xlsx'
dict['Centrum pre liečbu drogových závislostí Košice']['objednavky_faktury_link']='https://ba5c113364.clvaw-cdnwnd.com/a410a7ba52b55860f90d7278a3fc9ce5/200000168-2c40a2c40c/Objedn%C3%A1vky%20tovarov%2C%20slu%C5%BEieb%20a%20pr%C3%A1c%20a%20Fakt%C3%BAry.xlsx?ph=ba5c113364'
dict['Detská fakultná nemocnica Košice']['objednavky_faktury_link']='https://docs.google.com/uc?id=1uNSfzKDWdQfcwNJv5vBNS3yTGkGKlOIsK9FxXpx_edY'



opener = build_opener()
opener.addheaders = [('User-agent', 'Mozilla/5.0')]
install_opener(opener)

# 0 - Centrum pre liečbu drogových závislostí Banská\xa0Bystrica - link for scraping not available


# 1 - Centrum pre liečbu drogových závislostí Bratislava
urlretrieve(dict['Centrum pre liečbu drogových závislostí Bratislava']['objednavky_faktury_link'], data_path+current_date_time.strftime("%d-%m-%Y")+str(keysList[1]).replace(" ", "")+'.xlsx')

# 2 - Centrum pre liečbu drogových závislostí Košice
urlretrieve(dict['Centrum pre liečbu drogových závislostí Košice']['objednavky_faktury_link'], data_path+current_date_time.strftime("%d-%m-%Y")+str(keysList[2]).replace(" ", "")+'.xlsx')

# 3 - Detská fakultná nemocnica Košice
# TODO
#gdown.download(dict['Detská fakultná nemocnica Košice']['objednavky_faktury_link'], data_path+current_date_time.strftime("%d-%m-%Y")+str(keysList[3]).replace(" ", "")+'.xlsx')









