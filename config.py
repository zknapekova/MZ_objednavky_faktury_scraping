import os

source_path = 'C:\\Users\\knapekoz\\OneDrive - health.gov.sk\\\Zverejnovanie zmluv linky_subjektyCO.xlsx'

data_path = os.getcwd() + "\\data\\"
if not os.path.exists(data_path):
    os.mkdir(data_path)

search_data_path = os.getcwd() + "\\search_data\\"
if not os.path.exists(search_data_path):
    os.mkdir(search_data_path)

vo_path = "C:\\Users\\knapekoz\\health.gov.sk\\OSCM - Posudzovanie žiadostí\\"
chromedriver_path = os.path.join(vo_path, 'chromedriver_win32/chromedriver.exe')
chromedriver_path2 = 'C:\\Users\\knapekoz\\Documents\\Python Scripts\\chromedriver\\chromedriver.exe'

stand_column_names = {
    'objednavatel': ['nazov verejneho obstaravatela'],
    'kategoria': ['kategoria zakazky(tovar/stavebna praca/sluzba)', 'kategoria(tovar/stavebna praca/sluzba)',
                  'kategoria(tovary / prace / sluzby)'],
    'objednavka_predmet': ['nazov predmetu objednavky', 'predmet objednavky'],
    'cena': ['hodnotaobjednavkyv eur bez dph', 's.nc bdph', 'hodnotaobjednavkyv eur bez dph',
                                   'hodnota'],
    'datum': ['datum zadania objednavky', 'datum objednavky'],
    'objednavka_cislo': ['c.obj.', 'cislo objednavky'],
    'zdroj_financovania': ['zdroje financovania'],
    'balenie': ['balenie'],
    'sukl_kod': ['sukl_kod'],
    'mnozstvo': ['mnozstvo'],
    'poznamka': ['kratke zdovodnenie', 'kratke zdovodnenie2'],
    'dodavatel_ico': ['dodavatel - ico'],
    'dodavatel_nazov': ['dodavatel - nazov'],
    'odkaz_na_zmluvu': ['odkaz na zverejnenu zmluvu'],
    'pocet_oslovenych': ['pocet oslovenych']
}

