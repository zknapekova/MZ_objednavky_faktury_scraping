import os
import functionss as func

# Constants
clean_table_regex = r'(vypracovala*\s*:.*)|(schvalil.*:.*)|(.*riaditel$)|(.*referent .*)|(vysvetlivky:)|(^v stlpci.*uviest.*)|(.*@.*\.sk.*)|(^informacia o.*)|(^kontakt:.*)|(dna: \d+.*)|(judr.*)|(ing.*)|(mgr.*)|(datum zaciatku platnosti)|(^vyhotovil.*)|(anna kasmanova)|(^- za lekaren$)|(^- za referat prevadzky a udrzby$)'

# Paths ###
source_path = 'C:\\Users\\knapekoz\\OneDrive - health.gov.sk\\\Zverejnovanie zmluv linky_subjektyCO.xlsx'

data_path = os.getcwd() + "\\data\\"
if not os.path.exists(data_path):
    os.mkdir(data_path)

search_data_path = os.getcwd() + "\\search_data\\"
if not os.path.exists(search_data_path):
    os.mkdir(search_data_path)

historical_data_path = os.path.join(data_path + "historical_data\\")
if not os.path.exists(historical_data_path):
    os.mkdir(historical_data_path)

vo_path = "C:\\Users\\knapekoz\\health.gov.sk\\OSCM - Posudzovanie žiadostí\\"
chromedriver_path = os.path.join(vo_path, 'chromedriver_win32/chromedriver.exe')
chromedriver_path2 = 'C:\\Users\\knapekoz\\Documents\\Python Scripts\\chromedriver\\chromedriver.exe'


# Dictionaries ###

stand_column_names = {
    'objednavatel': ['nazov verejneho obstaravatela'],
    'kategoria': ['kategoria zakazky(tovar/stavebna praca/sluzba)', 'kategoria(tovar/stavebna praca/sluzba)', 'sluzba', 'kategoria (tovary / prace / sluzby)',
                  'kategoria(tovary / prace / sluzby)', 'kategoria (tovar/stavebna praca/sluzba)', 'kategoria zakazky (tovar/stavebna praca/sluzba)'],
    'objednavka_predmet': ['nazov predmetu objednavky', 'predmet objednavky', 'nazov predmetu zakazky', 'predmet zakazky', 'nazou predmetu objednavky'],
    'cena': ['hodnotaobjednavkyv eur bez dph', 's.nc bdph', 'hodnotaobjednavkyv eur bez dph', 'hodnota zakazky      s dph', 'hodnota objednavky',
                                   'hodnota', 'predpokladana hodnota v eur bez dph', 'hodnota objednavky v eur bez dph', 'cena v eur bez dph', 'hodnota objednavky v eur s dph'],
    'datum': ['datum zadania objednavky', 'datum objednavky', 'cislo oznamenia o vyhlaseni vo/ cislo vestnika/ datum zverejnenia', 'datum', 'platnost zmluvy'],
    'objednavka_cislo': ['c.obj.', 'cislo objednavky'],
    'zdroj_financovania': ['zdroje financovania'],
    'balenie': ['balenie'],
    'sukl_kod': ['sukl_kod'],
    'mnozstvo': ['mnozstvo'],
    'poznamka': ['kratke zdovodnenie', 'kratke zdovodnenie2', 'kratke zdovodnenie (zostavajuca hodnota)', 'kratke zdovodnenie/zostavajuca hodnota'],
    'dodavatel_ico': ['dodavatel - ico'],
    'dodavatel_nazov': ['dodavatel - nazov', 'uspesny uchadzac', 'obchodne meno a sidlo dodavatela'],
    'odkaz_na_zmluvu': ['odkaz na zverejnenu zmluvu'],
    'pocet_oslovenych': ['pocet oslovenych']
}


# load excel and create dictionary
df = func.load_df(source_path)
df['nazov'] = df['Nazov_full']

# clean data
df = func.clean_str_cols(df, cols=['Nazov_full'])
df['Nazov_full'] = df['Nazov_full'].replace(',|\.', '', regex=True)
dict_all = df.set_index('Nazov_full').T.to_dict('dict')

keysList = list(dict_all.keys())

# update

dict_all['centrum pre liecbu drogovych zavislosti bratislava'][
        'objednavky_faktury_link'] = 'https://cpldz.sk/wp-content/uploads/Fakturyaobjednavky_december.xlsx'

dict_all['centrum pre liecbu drogovych zavislosti kosice'][
        'objednavky_faktury_link'] = 'https://ba5c113364.clvaw-cdnwnd.com/a410a7ba52b55860f90d7278a3fc9ce5/200000168-2c40a2c40c/Objedn%C3%A1vky%20tovarov%2C%20slu%C5%BEieb%20a%20pr%C3%A1c%20a%20Fakt%C3%BAry.xlsx?ph=ba5c113364'

dict_all['detska fakultna nemocnica kosice'][
        'objednavky_faktury_link'] = 'https://docs.google.com/uc?id=1uNSfzKDWdQfcwNJv5vBNS3yTGkGKlOIsK9FxXpx_edY'

dict_all['detska fakultna nemocnica s poliklinikou banska bystrica'][
        'objednavky_faktury_link'] = 'https://www.detskanemocnica.sk/sites/default/files/files/199/objednavky_od_2023.01.01_do_2023.02.17.pdf'
dict_all['detska fakultna nemocnica s poliklinikou banska bystrica']['objednavky_faktury_link2022']='https://www.detskanemocnica.sk/sites/default/files/files/199/objednavky_od_2022.01.01_do_2022.12.30.pdf'
dict_all['detska fakultna nemocnica s poliklinikou banska bystrica']['objednavky_faktury_link2021']='https://www.detskanemocnica.sk/sites/default/files/files/199/objednavky_od_2022.01.01_do_2022.12.30.pdf'


dict_all['detska fakultna nemocnica s poliklinikou banska bystrica']['objednavky_faktury_file_ext'] = '.pdf'


dict_all['detska psychiatricka liecebna n o hran'][
        'objednavky_faktury_link'] = 'https://zverejnovanie.mzsr.sk/ministerstvo-zdravotnictva-sr/objednavky/?export=csv&art_rok=2023'
dict_all['detska psychiatricka liecebna n o hran']['objednavky_faktury_file_ext'] = '.csv'

dict_all['fakultna nemocnica nitra']['objednavky_faktury_link'] = 'https://fnnitra.sk/objd/new/'
dict_all['fakultna nemocnica s poliklinikou f d roosevelta banska bystrica'][
        'objednavky_faktury_link'] = 'https://www.fnspfdr.sk/objednavky/zverejnenie.php?akcia=vsetkyobjednavky_internet'
dict_all['fakultna nemocnica s poliklinikou zilina']['objednavky_link'] = 'http://www.fnspza.sk/zm2019/objednavky'
dict_all['fakultna nemocnica trencin']['objednavky_link'] = 'https://www.fntn.sk/zverejnovanie/objednavky/zobraz/triedenie/order_date/smer/zostupne'



