import pandas as pd
import xlwings as xw
import os, sys
import xlrd
import time
import numpy as np
from unidecode import unidecode
import requests
import re
import olefile
import matplotlib.pyplot as plt
if not re.split('\\\\|/', os.path.realpath(__file__))[2]=='nelka':
    import winsound
    frequency = 2500  # Set Frequency To 2500 Hertz
    duration = 500  # Set Duration To 1000 ms == 1 second


def timeit(t=None, s='', interval=[-1,-2], ending=''):
    '''
    Parameters
    ----------
    t : list, optional
        List of times which is returned appended. The default is None.
    s : string, optional
        String to be printed. The default is ''.
    interval : list of length 2, optional
        Printed s is followed by t[interval[0]] - t[interval[1]]. The default is [-1,-2].

    Returns
    -------
    list
        List of times.

    Use example
    -------
    times = fn.timeit(t=None, s='', interval=[-1,-2])
    times = fn.timeit(t=times, s='Saved', interval=[-1,-2])
    '''
    if t is None:
        return [time.time()]
    else:
        t.append(time.time())
        int_sec = (t[interval[0]] - t[interval[1]])
        if not s=='':
            print(f'{s} {secs_to_nicer_time(int_sec)}{ending}')
        return t


def secs_to_nicer_time(secs):    
        if secs < 100:
            leading_zeros = 0
            try:
                while str(secs).replace('.','')[leading_zeros] == '0':
                    leading_zeros += 1
            except: pass
            nice_time = f'{secs:.{2+leading_zeros}f}s'
        elif secs < 3600:
            nice_time = f'{secs//60:.0f}m {secs%60:.0f}s'
        else:
            nice_time = f'{secs//3600:.0f}h {secs%3600/60:.0f}m'
        return nice_time



def new_progress_bar (iteration, total, 
                      prefix = '', 
                      suffix = '', 
                      decimals = 1, 
                      length = 100, 
                      fill = '█', 
                      times = None,
                      printEnd = "\r"):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        times       - Optional  : start time (list of times)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    use at the beginning of for loop
    """
    remaining_time = ''
    ratio = iteration/total
    if times and iteration==1:
        timeit(t=times, s='', interval=[-1,-2])
    elif times and iteration>1:
        int_sec = (time.time() - times[-1])
        remaining_sec = int_sec/(iteration-1) * (total-(iteration-1))
        remaining_time = secs_to_nicer_time(remaining_sec)
    
    percent = f'{100*(ratio):.{decimals}f}'
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    # Print New Line on Complete
    if iteration == total: 
        if times:
            print(f'\r{prefix} |{bar}| {percent}%, {iteration}/{total} {suffix} finished in {secs_to_nicer_time(int_sec)}     ', end = printEnd)
            print()
    else:
        print(f'\r{prefix} |{bar}| {percent}%, {iteration}/{total} {suffix} {remaining_time} left     ', end = printEnd)
#fn.new_progress_bar(i+1, len(files), prefix='Downloading', suffix='', times=times, length=25)


def printProgressBar (iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print('\r%s |%s| %s%% %s' % (prefix, bar, percent, suffix), end = printEnd)
    # Print New Line on Complete
    if iteration == total: 
        print()


def save_df(df, name, 
            path = None, 
            replace = True, 
            dict_of_sheets = False,
            sheet_form_func = None,
            **kwargs):
    """
    Saves pandas dataframe into .xlsx, .csv, .dta files
    @params:
        df           - Required  : pandas dataframe object
        name        - Required  : name with file extension (Str)
        path         - Optional  : path to save folder (Str)
        kwargs      - Optional  : passed to individual save functions
    """
    
    def find_free_f_name(name, path):
        f_path = os.path.join(path, name)
        if not os.path.exists(f_path):
            return f_path
        else:
            return find_free_f_name('.'.join(name.split('.')[:-1])+'1.'+name.split('.')[-1], path)
    
    if not path:
        path = os.path.dirname(os.path.abspath(__file__))
    f_path = os.path.join(path, name)
    
    if os.path.exists(f_path) and replace:
        os.remove(f_path)
    elif os.path.exists(f_path) and (not replace):
        f_path = find_free_f_name(name, path)

    if name.split('.')[-1][:3] == 'xls':
        if not dict_of_sheets:
            #df.to_excel(f_path, **kwargs)
            writer = pd.ExcelWriter(f_path, engine='xlsxwriter')
            df.to_excel(writer, sheet_name='Harok 1', **kwargs)  # send df to writer
            worksheet = writer.sheets['Harok 1']  # pull worksheet object
            (max_row, max_col) = df.shape
            worksheet.autofilter(0, 0, max_row, max_col - 1)
            worksheet.freeze_panes(1, 0)
            try:
                for idx, col in enumerate(df):  # loop through all columns
                    series = df[col]
                    max_len = max((
                        series.astype(str).map(len).max(),  # len of largest item
                        len(str(series.name))  # len of column name/header
                        )) + 1  # adding a little extra space
                    worksheet.set_column(idx, idx, min((60, max_len)))  # set column width
            except:
                pass            
            
            if sheet_form_func:
                sheet_form_func(writer, worksheet)

            writer.save()
            #writer.close()
        elif dict_of_sheets:
            writer = pd.ExcelWriter(f_path, engine='xlsxwriter')
            for sheetname, df_part in df.items():  # loop through `dict` of dataframes
                df_part.to_excel(writer, sheet_name=sheetname, **kwargs)  # send df to writer
                worksheet = writer.sheets[sheetname]  # pull worksheet object
                for idx, col in enumerate(df_part):  # loop through all columns
                    series = df_part[col]
                    max_len = max((100,
                        series.astype(str).map(len).max(),  # len of largest item
                        len(str(series.name))  # len of column name/header
                        )) + 1  # adding a little extra space
                    worksheet.set_column(idx, idx, min((60, max_len)))  # set column width
            writer.save()
    elif name.split('.')[-1][:3] == 'csv':
        df.to_csv(path_or_buf=f_path, **kwargs)
    elif name.split('.')[-1][:3] == 'dta':
        df.to_stata(f_path, **kwargs)
    elif name.split('.')[-1][:3] == 'pkl':
        if 'index' in kwargs:
            del kwargs['index']
        df.to_pickle(f_path, **kwargs)
    elif name.split('.')[-1][:3] == 'fea':
        df.to_feather(f_path, **kwargs)
    else:
        print('unknown file extension:', name)
    '''    
        saves = dict(
            prehlad = pd.DataFrame({'Keys':one_of.keys(),'Values':one_of.values()}),
            pacienti = pacienti,
            vykony = vykony,
            diagnozy = diagnozy,
            diag_vyk_pary = diag_vyk_pair,
            vsetko = anom_loca,
            )
        writer = pd.ExcelWriter(save_loca, engine='xlsxwriter')
        for sheetname, df in saves.items():  # loop through `dict` of dataframes
            df.to_excel(writer, sheet_name=sheetname, index=False)  # send df to writer
            worksheet = writer.sheets[sheetname]  # pull worksheet object
            for idx, col in enumerate(df):  # loop through all columns
                series = df[col]
                max_len = max((
                    series.astype(str).map(len).max(),  # len of largest item
                    len(str(series.name))  # len of column name/header
                    )) + 1  # adding a little extra space
                worksheet.set_column(idx, idx, max_len)  # set column width
        writer.save()'''


def load_df(name, path = None, verbose=False, fail_loud=True, **kwargs):
    """
    Loads .xls, .csv, .dta files into pandas dataframe
    @params:
        name        - Required  : name with file extension (Str)
        path         - Optional  : path to save folder (Str)
        kwargs      - Optional  : passed to individual save functions
    """
    if not path:
        if len(name.split('\\')) > 1:
            path = '\\'.join(name.split('\\')[:-1])
        else:
            path = os.path.dirname(os.path.abspath(__file__))    
    f_path = os.path.join(path, name)
    if name.split('.')[-1][:4] == 'xlsb':
        df = pd.read_excel(f_path, engine='pyxlsb',**kwargs)
    elif name.split('.')[-1][:3] == 'xls':
        try:
            workbook = xlrd.open_workbook(f_path, ignore_workbook_corruption=True)
            df = pd.read_excel(workbook, **kwargs)
            if verbose: print('Loadujem cez xlrd pandas reader')
        except:
            try:
                df = pd.read_excel(f_path, **kwargs)
                if verbose: print('Loadujem cez pandas reader')
            except:
                try:
                    with olefile.OleFileIO(f_path) as ole:
                        if ole.exists('Workbook'):
                            d = ole.openstream('Workbook')
                            df = pd.read_excel(d, engine='xlrd')
                    if verbose: print('Loadujem cez olefile')
                except:
                    try:
                        app = xw.App(visible=False, add_book=False)
                        xls = app.books.open(f_path)
                        sheet = xls.sheets[0]
                        df1 = sheet.range('A1').options(pd.DataFrame, header=1, index=False, expand='table').value
                        df = df1.copy()
                        del df1
                        app.quit()
                        if verbose: print('Loadujem cez xw.app')
                    except:
                        if verbose: print('fakt neviem precitat ten excel')
                        if fail_loud:
                            print(f'failed with {f_path}')
                            exit()
    elif name.split('.')[-1][:3] == 'csv':
        df = pd.read_csv(f_path, **kwargs)
    elif name.split('.')[-1][:3] == 'dta':
        df = pd.read_stata(f_path, **kwargs)
    elif name.split('.')[-1][:3] == 'xml':
        df = pd.read_xml(f_path, **kwargs)
    elif name.split('.')[-1][:3] == 'pkl':
        df = pd.read_pickle(f_path, **kwargs)
    elif name.split('.')[-1][:3] == 'fea':
        df = pd.read_feather(f_path, **kwargs)
    elif name.split('.')[-1][:4] == 'json':
        df = pd.read_json(f_path, **kwargs)
        '''
    elif name.split('.')[-1][:3] == 'sql':
        import sqlite3
        conn = sqlite3.connect(f_path)
        df = pd.read_sql(f_path, conn, **kwargs)'''
    else:
        if verbose: print('unknown file extension:', name)    
    return df


def concat_xl_files(DB_dir, name_filter = [0,''], f_names=None ):
    """
    Used this to check sheet names and first rows of sheets (keys in database) in multiple excel files.
    @params:
        DB_dir      - Required  : Database folder path to be listed (Str)
        n_sheets     - Required  : total iterations (Int)
        name_filter - Optional  : example: [5,'Lieky'] -> file name[:5] has to equal 'Lieky'
    Returns appended excel sheet in pandas dataframe format
    """    
    # Read file names in DB_dir and filter them
    if f_names is None:
        f_names = os.listdir(DB_dir)
        f_names = [f_name for f_name in f_names if f_name[:name_filter[0]] == name_filter[1]]
    #f_names = f_names[:3]
    
    #Read and join all sheets in all files into merged dataframe
    dfs = []
    printProgressBar(0, len(f_names), prefix = 'file iter:', suffix = f_names[0], length = 25)
    for i, f_name in enumerate(f_names):
        if f_name.split('.')[-1][0] == 'o':
            xls = pd.ExcelFile(os.path.join(DB_dir, f_name), engine = 'odf')
        else:
            xls = pd.ExcelFile(os.path.join(DB_dir, f_name))
        sheets = xls.sheet_names 
        for sheet in sheets:
            df = xls.parse(sheet)
            df['Typ'] = sheet.lower().strip()
            df['file'] = f_name
            df.columns = df.columns.str.lower().str.strip()
            #df.columns = df.columns.str.strip()
            dfs.append(df)
        printProgressBar(i+1, len(f_names), prefix = 'file iter:', suffix = f_names[i], length = 25)        
    empty = pd.DataFrame()
    merged = empty.append(dfs, ignore_index = True)
    return merged


def simpler_concat_xl_files(DB_dir, name_filter = [0,''], f_names=None, times=None, **kwargs):
    """
    Used this to check sheet names and first rows of sheets (keys in database) in multiple excel files.
    @params:
        DB_dir      - Required  : Database folder path to be listed (Str)
        n_sheets     - Required  : total iterations (Int)
        name_filter - Optional  : example: [5,'Lieky'] -> file name[:5] has to equal 'Lieky'
    Returns appended excel sheet in pandas dataframe format
    """    
    # Read file names in DB_dir and filter them
    if f_names is None:
        f_names = os.listdir(DB_dir)
        f_names = [f_name for f_name in f_names if f_name[:name_filter[0]] == name_filter[1]]
    #f_names = f_names[:3]
    
    #Read and join all sheets in all files into merged dataframe
    dfs = []
    #printProgressBar(0, len(f_names), prefix = 'file iter:', suffix = f_names[0], length = 25)
    for i, f_name in enumerate(f_names):
        new_progress_bar(i+1, len(f_names), prefix='Concatenating', suffix='', times=times, length=25)
        df = load_df(f_name, path=DB_dir, **kwargs)
        df['file'] = f_name
        dfs.append(df)
        #printProgressBar(i+1, len(f_names), prefix = 'file iter:', suffix = f_names[i], length = 25)        
    empty = pd.DataFrame()
    merged = empty.append(dfs, ignore_index = True)
    return merged

def to_numeric(df):
    pd.options.mode.chained_assignment = None
    for col in df.columns:
        old_isna = df[col].isna().sum()
        df[col] = pd.to_numeric(df[col].astype('string')
                           .str.replace(u'\xa0', u' ')
                           .str.replace('EUR','')
                           .str.replace(' ','')
                           .str.replace(',','.'), errors='coerce')
        new_isna = df[col].isna().sum()
        diff = new_isna - old_isna
        if diff > 0:
            print(f'Bolo pridanych {diff} Nans ({diff/len(df):.2f}%) v {col}')
    pd.options.mode.chained_assignment = 'warn'
    return df




def clean_str_cols(df, cols=None, unicode=True, lower=True, strip=True, xa0=True):
    #dfo = clean_str_cols(dfo, cols=None, unicode=True, lower=True, strip=True, xa0=True)
    if cols is None:
        cols = df.columns
    for col in cols:
        if df[col].dtype == 'O':
            df[col] = df[col].astype(str)
            if unicode: df[col] = df[col].map(unidecode)
            if lower: df[col] = df[col].str.lower()
            if strip: df[col] = df[col].str.strip()
            if xa0: df[col] = df[col].str.replace(u'\xa0', u' ')
            df.loc[df[col]=='nan', col] = np.nan
    return df

def clean_float_cols(df, cols=None):
    #dfo = clean_str_cols(dfo, cols=None, unicode=True, lower=True, strip=True, xa0=True)
    if cols is None:
        cols = df.columns
    for col in cols:
        if df[col].dtype =='float64' or df[col].dtype =='float32':
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(u'\xa0', u' ').str.replace('EUR','').str.replace(' ','').str.replace(',','.'), errors='coerce')
    return df

def round_to_x_cifers_after_nonzero(df, col, n):
    df['num_text'] = df[col].astype(str).str.replace('\.','', regex=True)
    df['dot_pos'] = df[col].astype(str).str.split('.').str[0].str.len()
    df['lead'] = df['num_text'].str.split(r'1|2|3|4|5|6|7|8|9').str[0]
    df['strip'] = df['num_text'].str.replace('^0+','', regex=True)
    df['num'] = df['strip'].str[:n].astype(int)
    df['oldnum'] = df['num'].values
    loca = (df['strip'].str[n].fillna('0').astype(int)>=5)
    df.loc[loca, 'num'] = df.loc[loca]['num'] + 1
    df['dot_pos'] = df['dot_pos'] + df['num'].astype(str).str.len() - df['oldnum'].astype(str).str.len()
    df['trailing0s'] = [(x-n)*'0' for x in df['dot_pos']]
    df['numstr'] = df['lead'] + df['num'].astype(str) + df['trailing0s']
    df[col] = [x[1][:x[0]]+'.'+x[1][x[0]:] for x in df[['dot_pos','numstr']].values]
    df[col] = df[col].str.replace('--','-')
    df[col] = df[col].astype(float)
    df.drop(['num_text','dot_pos','lead','strip','oldnum','num','trailing0s','numstr'], axis = 1, inplace = True)
    return df


def time_stuff():
    #leaps = [2008, 2012, 2016, 2020, 2024, 2028]
    years = [2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023]
    #days_in_months = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    days_in_months_cum = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365]
    #days_in_months_leap = [0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    days_in_months_leap_cum = [0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366]
    hols_in_months_dict = {
        '2013': [[1,6],[],[29 ],[1 ],[1,8],[],[5],[29],[1,15],[],[1,17],[24,25,26]],
        '2014': [[1,6],[],[],[18,21],[1,8],[],[5],[29],[1,15],[],[1,17],[24,25,26]],
        '2015': [[1,6],[],[],[3 ,6 ],[1,8],[],[5],[29],[1,15],[],[1,17],[24,25,26]],
        '2016': [[1,6],[],[25,28],[],[1,8],[],[5],[29],[1,15],[],[1,17],[24,25,26]],
        '2017': [[1,6],[],[],[14,17],[1,8],[],[5],[29],[1,15],[],[1,17],[24,25,26]],
        '2018': [[1,6],[],[30 ],[2 ],[1,8],[],[5],[29],[1,15],[],[1,17],[24,25,26]],
        '2019': [[1,6],[],[],[19,22],[1,8],[],[5],[29],[1,15],[],[1,17],[24,25,26]],
        '2020': [[1,6],[],[],[10,13],[1,8],[],[5],[29],[1,15],[],[1,17],[24,25,26]],
        '2021': [[1,6],[],[],[2 ,5 ],[1,8],[],[5],[29],[1,15],[],[1,17],[24,25,26]],
        '2022': [[1,6],[],[],[15,18],[1,8],[],[5],[29],[1,15],[],[1,17],[24,25,26]],
        '2023': [[1,6],[],[],[7 ,10],[1,8],[],[5],[29],[1,15],[],[1,17],[24,25,26]]}
    first_weekend_dict = {
        '2013': [5,6],
        '2014': [4,5],
        '2015': [3,4],
        '2016': [2,3],
        '2017': [1,7],
        '2018': [6,7],
        '2019': [5,6],
        '2020': [4,5],
        '2021': [2,3],
        '2022': [1,2],
        '2023': [1,7]}
    return years, days_in_months_cum, days_in_months_leap_cum, hols_in_months_dict, first_weekend_dict


def work_days_in_year(year, verbose=False):
    _, days_in_months_cum, days_in_months_leap_cum, hols_in_months_dict, first_weekend_dict = time_stuff()
    if year%4 == 0:
        DinM = days_in_months_leap_cum
        days_tot = 366
    else:
        DinM = days_in_months_cum
        days_tot = 365
    df = pd.DataFrame(np.arange(1,days_tot+1))
    df.columns = ['days']
    df['dates'] = pd.date_range(start=f'{year}-1-1', end=f'{year+1}-1-1', closed='left')
    df['work_day'] = 1
    df.loc[(df['days']-first_weekend_dict[str(year)][0])%7 == 0, 'work_day'] = 0
    df.loc[(df['days']-first_weekend_dict[str(year)][1])%7 == 0, 'work_day'] = 0
    holidays = [[idk+days for idk in hols_in_months_dict[str(year)][i]] for i,days in enumerate(DinM[:-1])]
    holidays = [item for sublist in holidays for item in sublist]
    for holiday in holidays:
        if (not (holiday - first_weekend_dict[str(year)][0])%7 == 0) and \
            (not (holiday - first_weekend_dict[str(year)][1])%7 == 0):
            df.loc[df['days'] == holiday, 'work_day'] = 0
    if verbose:
        print(f"Total days = {df['days'].max()} in {year}")
        print(f"Work days = {df['work_day'].sum()}, holidays + weekends = {df['days'].max()-df['work_day'].sum()}")
    return df

def work_days_in_years(start=2020, end=2021):
    year = start
    df = pd.DataFrame({})
    while year < end:
        df = df.append(work_days_in_year(year, verbose=False))
        year += 1
    return df
    
       
def date_to_days(df, date_col='DATUM', work_days=False, start=None):
    years, days_in_months_cum, days_in_months_leap_cum, hols_in_months_dict, first_weekend_dict = time_stuff()
    
    days_col = 'DAYS_'+date_col
    str_date_col = 'str'+date_col
    
    df[str_date_col] = df[date_col].astype(str)
    df[days_col] = df[str_date_col].str[4:6].astype(int)
    df[days_col] = df[days_col] - 1
    
    df['int_day_col'] = df[str_date_col].str[6:].astype(int)
    df['int_year_col'] = df[str_date_col].str[:4].astype(int)
    df_years = df['int_year_col'].unique()
    df_years = np.sort(df_years)
    
    for year in df_years:
        if not year in years: print('dataset has years date_to_days function is not ready for')
    if not start is None:
        if start < df_years[0]:
            years = years[years.index(start):years.index(df_years[-1])+1]
        else: print('date_to_days function is not ready to start later then the year in dataset, only sooner')
    else:
        years = years[years.index(df_years[0]):years.index(df_years[-1])+1]
        
    days_in_years = [0]
    for year in years:
        if year%4 == 0: days_in_years.append(366)
        else: days_in_years.append(365)
    days_in_years = [sum(days_in_years[:i+1]) for i in range(len(days_in_years))]
    
    for i, year in enumerate(years):
        if year in df_years:
            loca = df['int_year_col'] == int(year)
            print(year, i, loca.sum())
            if year%4 == 0:
                DinM = days_in_months_leap_cum
            else:
                DinM = days_in_months_cum
            df.loc[loca, days_col] = df.loc[loca][days_col].apply(lambda x: DinM[x])
            df.loc[loca, days_col] = df.loc[loca][days_col] + df.loc[loca]['int_day_col'] + days_in_years[i]
            if work_days:
                df['WORK_DAY'] = 1
                df.loc[(loca) & ((df[days_col]-first_weekend_dict[str(year)][0])%7 == 0), 'WORK_DAY'] = 0
                df.loc[(loca) & ((df[days_col]-first_weekend_dict[str(year)][1])%7 == 0), 'WORK_DAY'] = 0
                holidays = [[idk+days for idk in hols_in_months_dict[str(year)][i]] for i,days in enumerate(DinM[:-1])]
                holidays = [item for sublist in holidays for item in sublist]
                for holiday in holidays:
                    if (not (holiday - first_weekend_dict[str(year)][0])%7 == 0) and \
                        (not (holiday - first_weekend_dict[str(year)][1])%7 == 0):
                        df.loc[(loca) & (df[days_col] == holiday), 'WORK_DAY'] = 0
            
    df = df.drop([str_date_col,'int_year_col','int_day_col'], axis=1)
    return df


def myplot(*args,
           ptype = 'plot',
           x_lab = '',
           y_lab = '',
           title='',
           leg = [],
           grid = True,
           **kwargs):
    '''
    Parameters
    ----------
    *args : array-like
        Data passed to plt functions for plotting.
    ptype : 'plot'/'hist'/'scatter', optional
        Specifies type of plot. The default is 'plot'.
    x_lab : string, optional
        X axis label. The default is ''.
    y_lab : string, optional
        Y axis label. The default is ''.
    leg : list of strings, optional
        Legend. The default is [].
    **kwargs : TYPE
        Passed to plt functions.

    Returns
    -------
    None.
    
    Use example
    -------
    fn.myplot(x1, y1, ptype='plot', x_lab='Time [date]', y_lab='Column [-]')
    '''
    
    if ptype == 'plot':
        plt.plot(*args,*kwargs)
    if ptype == 'hist':
        plt.hist(*args,*kwargs)
    if ptype == 'scatter':
        plt.scatter(*args,*kwargs)
    if ptype == 'imshow':
        plt.imshow(*args,*kwargs)
    plt.rcParams["figure.dpi"] = 300
    plt.ylabel(y_lab)
    plt.xlabel(x_lab)
    if not leg=='':
        plt.legend(leg, loc='best')
    if not title=='':
        plt.title(title)
    plt.grid(grid)
    plt.show()    


def analyse_df(df, txt_f_name, group_by = None, values_c = 20 ):
    """
    Used this to check sheet names and first rows of sheets (keys in database) in multiple excel files.
    @params:
        df           - Required  : pandas dataframe object
        txt_f_name  - Required  : txt file name (Str)
        group_by     - Optional  : group output by unique values in 'group_by' column (str)
        values_c     - Optional  : length of printed most common values
    """    
    dfs = []
    vals = ['']
    if group_by:
        vals = df[group_by].unique()
        for val in vals:
            dfs.append(df[df[group_by] == val])
    else:
        dfs.append(df)
            
    
    lines = []
    for i, df in enumerate(dfs):
        lines.append('({}/{})  {}: {}\n'.format(i+1, len(dfs), group_by, vals[i]))
        max_key = max([len(str(key)) for key in df.dtypes.keys()])
        for key, dtype in df.dtypes.items():
            line = ''
            nunique = df[key].nunique()
            nans = df[key].isna().sum()
            #print(key, dtype, nunique, nans, nans/len(df)*100, max_key)
            line += '{:<{fill}} {:<15s} nunique = {:<10d} Nans = {:<10d} ({:<5.1f}%)\t'.format(key, str(dtype), nunique, nans, nans/len(df)*100, fill = max_key)
            if not dtype == 'O':
                line += 'Min = {:<10}\tMax = {:<10}'.format(df[key].min(), df[key].max())        
            
            # Get most common values and their counts
            values_counts = df[key].value_counts()
            val = [str(v) for v in values_counts.index.tolist()[:values_c]]
            counts = [str(c) for c in values_counts.tolist()[:values_c]]
            line += 'Values: '
            for val_c in zip(val, counts):
                line += val_c[0]+' ('+val_c[1]+'), '
            
            lines.append(line+'\n')
        lines.append('\n')
    
    keys_f = open(txt_f_name,"w")
    keys_f.writelines(lines)
    keys_f.close()    



def request_DB_liekov():
    '''
    200: Everything went okay, and the result has been returned (if any).
    301: The server is redirecting you to a different endpoint. This can happen when a company switches domain names, or an endpoint name is changed.
    400: The server thinks you made a bad request. This can happen when you don’t send along the right data, among other things.
    401: The server thinks you’re not authenticated. Many APIs require login ccredentials, so this happens when you don’t send the right credentials to access an API.
    403: The resource you’re trying to access is forbidden: you don’t have the right permissions to see it.
    404: The resource you tried to access wasn’t found on the server.
    503: The server is not ready to handle the request.
    '''
    response = requests.get('https://api.sukl.sk/json/lieky.php')
    print(response.status_code)
    print(response.json())
    return response

def prep_name_pair_file(DB_dir, txt_f_name):
    '''
    Used this when renaming files from hospitals, it creates txt file
    which pairs to old names to prepped Lieky_SZM_2019-___.xls names

    @params:
        DB_dir      - Required  : Database folder path to be listed (Str)
        txt_f_name  - Required  : txt life name (Str)
    '''
    print(os.listdir(DB_dir))
    names = os.listdir(DB_dir)
    names = [name+'\t\tLieky_SZM_2019-.'+name.split('.')[-1]+'\n' for name in names]
    
    names_f = open(os.path.join(DB_dir,txt_f_name),"w")
    names_f.writelines(names)
    names_f.close()


def print_xl_keys(DB_dir, txt_f_name, n_sheets, name_filter = [0,''] ):
    """
    Used this to check sheet names and first rows of sheets (keys in database) in multiple excel files.
    @params:
        DB_dir      - Required  : Database folder path to be listed (Str)
        txt_f_name  - Required  : txt life name (Str)
        n_sheets     - Required  : total iterations (Int)
        name_filter - Optional  : example: [5,'Lieky'] -> file name[:5] has to equal 'Lieky'
    """    
    # Read file names in DB_dir and filter them
    f_names = os.listdir(DB_dir)
    f_names = [f_name for f_name in f_names if f_name[:name_filter[0]] == name_filter[1]]
    
    # Open files and and put their sheet names and respective sheet keys into dictionary
    xl_keys = dict()
    printProgressBar(0, len(f_names), prefix = 'file iter:', suffix = f_names[0], length = 25)
    for i, f_name in enumerate(f_names):
        if f_name.split('.')[-1][0] == 'o':
            xls = pd.ExcelFile(os.path.join(DB_dir, f_name), engine = 'odf')
        else:
            xls = pd.ExcelFile(os.path.join(DB_dir, f_name))
        sheets = xls.sheet_names 
        xl_keys[f_name] = dict(sheets = sheets)
        for sheet in sheets:
            df = xls.parse(sheet)
            #print(df)
            xl_keys[f_name][sheet] = df.keys().to_list()
        #break
        printProgressBar(i+1, len(f_names), prefix = 'file iter:', suffix = f_names[i], length = 25)        
    
    # Write xl_keys dictionary into txt file so that names of sheets and keys of sheet are neatly underneath each other
    n = 1 + n_sheets
    lines = [[] for _ in range(n)]
    for f_name, k_dict in xl_keys.items():
        lines[0].append(f'{f_name:<50s}' +' '.join(k_dict['sheets'])+'\n')
        for i in range(n_sheets):
            try:
                sheet = k_dict['sheets'][i]
                lines[i+1].append(f'{f_name:<50s}  {sheet:<20s}  ' + ' '.join(k_dict[sheet])+'\n')
            except:
                lines[i+1].append(f'{f_name:<50s}'+ '\n')
                pass
    keys_f = open(os.path.join(DB_dir, txt_f_name),"w")
    for line in lines:
        keys_f.writelines(line)
    keys_f.close()

def preset_pandas(max_rows=100, max_cols=200, disp_width=320):
    #pd.set_option('display.max_rows', max_rows)
    #pd.set_option('display.max_columns', max_cols)
    #pd.set_option('display.width', disp_width)
    #pd.options.mode.use_inf_as_na = True
    times = timeit(t=None, s='', interval=[-1,-2])
    return times
    
    
    
def join_dfs(df, df_dict, on='', cols=[], add_on_prefix=False):
    df_dict = df_dict[cols]
    if add_on_prefix:
        cols[1:] = [str(on)+'_'+col for col in cols[1:]]
        df_dict.columns = cols
    if on not in cols: 
        df_dict = df_dict.rename(columns={cols[0]:on})
    
    df = df.join(df_dict.set_index(on), how='left', on=on)
    return df

#df = fn.join_dfs(df, df_zle_diagnozy, on='icd_10dg', cols=['icd_10dg', 'dummy_column'], add_on_prefix=False)

def pretty_print_num(num):
    b = f'{num:.2f}'.replace('.',',')[::-1]
    if len(b)>12: b = b[:6] + ' ' + b[6:9] + ' ' + b[9:12] + ' ' + b[12:]
    elif len(b)>9: b = b[:6] + ' ' + b[6:9] + ' ' + b[9:]
    elif len(b)>6: b = b[:6] + ' ' + b[6:] 
    return b[::-1]
    
    

def get_attributes(driver, element):
    return driver.execute_script('var items = {}; for (index = 0; index < arguments[0].attributes.length; ++index) { items[arguments[0].attributes[index].name] = arguments[0].attributes[index].value }; return items;', element)

def find_by(parent, tag, by_str, 
            by='class', 
            exact=True, 
            first=True, 
            max_wait=10, 
            sleep_step=0.5,
            verbose=True):
    something = None if first else []
    wait = 0
    while ((something is None) or (something == [])) and (wait < max_wait):
        try:
            for element in parent.find_elements_by_tag_name(tag):
                if exact:
                    if by == 'text':
                        if element.text == by_str:
                            if first:
                                something = element
                                break
                            else:
                                something.append(element)
                                
                    else:
                        if element.get_attribute(by) == by_str:
                            if first:
                                something = element
                                break
                            else:
                                something.append(element)
                else:
                    if by == 'text':
                        if by_str in element.text:
                            if first:
                                something = element
                                break
                            else:
                                something.append(element)
                    else:
                        if by_str in element.get_attribute(by):
                            if first:
                                something = element
                                break
                            else:
                                something.append(element)
            if first:
                if something is None:
                    exit()
            else:
                if something is []:
                    exit()
        except:
            time.sleep(sleep_step)
            wait += sleep_step
            
    if (something is None) and (verbose):
        print(f'Failed to find {by_str} in {tag}s in {max_wait}s')
    return something


def open_from_sidebar(driver, sidebar):
    bars={}
    vsetky_ziadosti = driver.find_element_by_xpath("//a[@title='Všetky Žiadosti']")
    bars['vsetky_ziadosti'] = vsetky_ziadosti
    try:
        na_vecne_posudenie = driver.find_element_by_xpath("//a[@title='Žiadosti na vecné posúdenie']")
        bars['na_vecne_posudenie'] = na_vecne_posudenie
    except:
        pass
    try:
        na_komisiu = driver.find_element_by_xpath("//a[@title='Žiadosti na posúdenie členom komisie']")
        bars['na_komisiu'] = na_komisiu
    except:
        pass
    mnou_posudene = driver.find_element_by_xpath("//*[contains(text(), 'Mnou posúdené')]")
    bars['mnou_posudene'] = mnou_posudene
    side_bar_toggle = find_by(driver, 'button', 'd-lg-block d-none navbar-toggler', by='class')
    '''bars = {
            'vsetky_ziadosti':vsetky_ziadosti,
            'na_vecne_posudenie':na_vecne_posudenie,
            'na_komisiu':na_komisiu,
            'mnou_posudene':mnou_posudene
            }'''
    choice = bars[sidebar]
   
    # Open ziadosti na specified sidebar
    if not 'sidebar-lg-show' in driver.find_elements_by_tag_name('body')[0].get_attribute('class'):
        side_bar_toggle.click()
        time.sleep(0.5)
    if not 'open' in driver.find_elements_by_tag_name('app-sidebar-nav-dropdown')[0].get_attribute('class'):
        driver.find_element_by_xpath("//a[@title='Žiadosti so súhlasom MZ SR']").click()
        time.sleep(0.5)
    choice.click()
    time.sleep(0.5)


def get_kody_na_vecne_posudenie(driver):
    open_from_sidebar(driver, 'na_vecne_posudenie')
    body = find_by(driver, 'tbody', '', by='class', exact=False)
    rows = body.find_elements_by_tag_name('tr')
    kody_na_posudenie = []
    for row in rows:    
        tds = row.find_elements_by_tag_name('td')
        kod_mzsr = tds[0].find_elements_by_tag_name('a')[0] # click on kod ziadosti
        kody_na_posudenie.append(kod_mzsr.get_attribute('text').strip())
    return kody_na_posudenie


def open_ziadost(driver, ziadost, kde='na_vecne_posudenie'):
    open_from_sidebar(driver, kde)
    found = False
    body = find_by(driver, 'tbody', '', by='class', max_wait=2)
    if not body is None:
        # Check for other pages
        prev_page = find_by(driver, 'li', 'pagination-prev', by='class', exact=False)
        next_page = find_by(driver, 'li', 'pagination-next', by='class', exact=False)
        if 'disabled' in prev_page.get_attribute('class') and 'disabled' in next_page.get_attribute('class'):
            print('No other pages')
        rows = body.find_elements_by_tag_name('tr')
        #ziadosti =[] # list of [kod string, clickable element]
        for row in rows:
            tds = row.find_elements_by_tag_name('td')
            kod_mzsr = tds[0].find_elements_by_tag_name('a')[0] # click on kod ziadosti
            #ziadosti.append([kod_mzsr.get_attribute('text').strip(), kod_mzsr])
            if ziadost==kod_mzsr.get_attribute('text').strip():
                found = True
                kod_mzsr.click()
                break
    if not found:
        print(f'Ziadost {ziadost} nebola najdena v {kde}')

#open_ziadost('2021-00571', kde='na_vecne_posudenie')
def wait_for_stale(element, max_wait=5, interval=0.5):
    wait = 0
    while wait < max_wait:
        try:
            element.text
            time.sleep(interval)
            wait += interval            
        except:
            break
        
def flatten_list_of_lists(lis):
    return [item for sublist in lis for item in sublist]

# Disable
def block_print():
    sys.stdout = open(os.devnull, 'w')

# Restore
def enable_print():
    sys.stdout = sys.__stdout__










def clean_df(df, tolerance = 0.01):
    """
    cleans df, too df specific for reuse
    @params:
        df             - Required  : pandas dataframe to be cleaned
    """
    # Replace special characters in column names
    df.rename(unidecode, axis='columns', inplace=True)
    df.drop(['nc b dph', 'unnamed: 16', 'ico'],axis = 1, inplace = True)
    
    # Replace special characters lower() and strip() string values in all columns 
    for key, dtype in df.dtypes.items():
        if not key == 'datum dodavky':
            #print(key)
            loca = ~df[key].isnull() & ~df[key].astype(str).str.isnumeric()
            df.loc[loca, key] = df.loc[loca][key].astype(str).map(unidecode).str.lower().str.strip()
    
    # Columns which should be numeric
    num_columns = ['pocet mj v jednom baleni', 'cena za mj v eur s dph',\
                   'cena za 1 balenie v eur s dph', 'pocet nakupenych mj',\
                   'pocet nakupenych baleni', 'celkovy nakup s dph']
    
    # Force columns to be numeric and datetime
    for key in num_columns:
        df[key] = pd.to_numeric(df[key],errors='coerce')    
    df['datum dodavky'] = pd.to_datetime(df['datum dodavky'], errors='coerce')
    
    ''' To check negative numbers across files
    df['neg'] = (df['cena za 1 balenie v eur s dph']<0.).astype(int)
    print(df.groupby(['file','typ'])['neg'].sum())
    '''
    # Drop retarded numeric and string values....
    df.loc[df['pocet mj v jednom baleni']<=0, 'pocet mj v jednom baleni'] = np.nan
    df.loc[df['cena za mj v eur s dph']<=0, 'cena za mj v eur s dph'] = np.nan
    df.loc[df['cena za 1 balenie v eur s dph']<=0, 'cena za 1 balenie v eur s dph'] = np.nan
    df.loc[df['celkovy nakup s dph']==0, 'celkovy nakup s dph'] = np.nan
    df.dropna(axis = 0, subset = ['celkovy nakup s dph'], inplace = True)

    # Drops row if both are NA
    df['essential'] = df[['cena za 1 balenie v eur s dph', 'pocet nakupenych baleni']].isna().sum(axis = 1)
    ''' To check mess in essential columns, 0:great, 2:both columns are missing, not good
    print(df.groupby(['file','typ'])['essential'].mean())
    '''
    df.drop(df.loc[df['essential']>1].index, inplace = True)
    
    df.reset_index(drop=True, inplace=True)
    
    # Create ideal price per package
    df['cen_bal'] = df['celkovy nakup s dph']/df['pocet nakupenych baleni']
    df.loc[df['cen_bal'].abs()==np.inf,'cen_bal'] = np.nan
    # Adjust price per package if it is within tolerance 
    df['adj_cen_bal'] = df['cena za 1 balenie v eur s dph']
    df.loc[(df['cen_bal']-df['cena za 1 balenie v eur s dph']).abs() < tolerance, 'adj_cen_bal'] = df['cen_bal']
    df['cena za 1 balenie v eur s dph'] = df['adj_cen_bal']
    
    ''' Veci pre PZZ excel, Jakub Slobodnik
    group = df[df['typ']=='lieky'].groupby(['file'])['celkovy nakup s dph'].sum().reset_index()
    print(group.sort_values('celkovy nakup s dph',ascending=False),group.sum())
    save_df(group, 'nemocnice.xlsx', replace = True)
    print(df['celkovy nakup s dph'].sum())
    '''
    
    # (re)Create ideal price per package and ideal number of packages  
    df['cen_bal'] = df['celkovy nakup s dph']/df['pocet nakupenych baleni']
    df.loc[df['cen_bal'].abs()==np.inf,'cen_bal'] = np.nan
    df['poc_bal'] = df['celkovy nakup s dph']/df['cena za 1 balenie v eur s dph']
    df.loc[df['poc_bal'].abs()==np.inf,'poc_bal'] = np.nan
    
    # Fill in price per package and number of packages from ideals, if real is unknown
    loca = df['cena za 1 balenie v eur s dph'].isna()
    df.loc[loca, 'cena za 1 balenie v eur s dph'] = df.loc[loca]['cen_bal']
    loca = df['pocet nakupenych baleni'].isna()
    df.loc[df['pocet nakupenych baleni'].isna(), 'pocet nakupenych baleni'] = df.loc[loca]['poc_bal']
    
    # Find discrepancies between total_price and price_per_package * ideal_number_of_packages 
    df['calc_cen'] = df['cena za 1 balenie v eur s dph']*df['pocet nakupenych baleni']
    df['abs_cen_diff'] = (df['calc_cen'] - df['celkovy nakup s dph']).abs()
    #print(df.groupby(['file','typ'])['abs_cen_diff'].sum())
    #group = df.groupby(['file'])['abs_cen_diff'].sum().reset_index()
    #print(group.sort_values('abs_cen_diff',ascending=False))
    #print(df.groupby(['file'])['abs_cen_diff'].sum().sum())
    
    # Drop prices with discrepancies above tolerance and unnecessary columns
    df.drop(df.loc[df['abs_cen_diff'] > tolerance].index, inplace = True)
    df.drop(['calc_cen', 'essential', 'cen_bal', 'poc_bal', 'adj_cen_bal', 'abs_cen_diff'],axis = 1, inplace = True)
    
    
    df['sukl'] = df['sukl'].str.replace(' ','')
    suklshit = ['-', ' ', '', '0', '-1', 'p-----', 'p----', 'unknown', 'nema', 'x', 'm', 'xx', 'xxx', 'xxxxx', 'enseal', 'md', 'cmd']
    for shi in suklshit:
        df.loc[df['sukl'] == shi, 'sukl'] = np.nan
    
    loca = ((df['typ'] == 'lieky') & ~df['sukl'].isna()) & df['sukl'].astype(str).str[0].str.isalpha()
    df.loc[loca, 'sukl'] = df.loc[loca]['sukl'].astype(str).str[1:]
    loca = ((df['typ'] == 'lieky') & ~df['sukl'].isna()) & df['sukl'].astype(str).str[0].str.isalpha()
    df.loc[loca, 'sukl'] = np.nan
    
    
    for i in range(1,4):
        loca = df.loc[(df['typ'] == 'lieky') &(~df['sukl'].isna()) & (df['sukl'].astype(str).str.len() == i)].index
        df.loc[loca, 'sukl'] = (5-i)*'0' + df.loc[loca]['sukl']

    loca = df.loc[(df['typ'] == 'lieky') & (~df['sukl'].isna()) & (df['sukl'].str.len()>5)].index
    df.loc[loca, 'sukl'] = np.nan



    #df.loc[(df['typ'] == 'lieky') &(~df['sukl'].isna()) & (df['sukl'].astype(str).str.len() == 2)].sukl
    df.reset_index(drop=True, inplace=True)
    #print(df['sukl'].value_counts()[:100])
    #print(df[~df['sukl'].isna() & df['sukl'].str.isalpha()]['sukl'])
    
    df['nazov lieku'] = df['nazov lieku'].fillna(df['nazov tovaru'])
    df.drop(['nazov tovaru'],axis = 1, inplace = True)
    df.columns = ['sukl', 'atc', 'nazov', 'velkost mj', 'N mj v bal', 'C MJ s dph',\
                  'C bal s dph', 'N mj', 'N bal', 'HZ s dph', 'nemocnica','lekaren', 'datum', \
                  'dodavatel','poznamka', 'typ', 'file', 'kod mzsr', 'kod tovaru', 'poskytovatel']
    df['typ'].replace('szm nekategorizovane', 'szm nekat', inplace=True)
    df['typ'].replace('szm kategorizovane', 'szm kat', inplace=True)
    return df


def get_lieky(df, save_name):
    df = df.loc[(df['typ'] == 'lieky') & (~df['sukl'].isna()) & (~df['celkovy nakup s dph'].isna())]
    groupby = df.groupby(['sukl'])
    group = groupby['cena za 1 balenie v eur s dph'].mean().reset_index()
    group.columns = ['sukl', 'mean_cena_bal']
    group['n_nakupov'] = groupby.size().values
    group['nakupenych_baleni'] = groupby['pocet nakupenych baleni'].sum().values
    group['celkovy_objem_EUR'] = groupby['celkovy nakup s dph'].sum().values
    group['min_cena_bal'] = groupby['cena za 1 balenie v eur s dph'].min().values
    group['max_cena_bal'] = groupby['cena za 1 balenie v eur s dph'].max().values
    group['median_cena_bal'] = groupby['cena za 1 balenie v eur s dph'].median().values

    cols=['sukl',
         'min_cena_bal',
         'median_cena_bal',
         'mean_cena_bal',
         'max_cena_bal',
         'n_nakupov',
         'nakupenych_baleni',
         'celkovy_objem_EUR']
    
    save_df(group[cols], save_name+'.xlsx', path = None, replace = True)
    
    

def analyse_float_cols(df, n_outs = 30):
    """
    Used this to check sheet names and first rows of sheets (keys in database) in multiple excel files.
    @params:
        df           - Required  : pandas dataframe object
        group_by     - Optional  : group output by unique values in 'group_by' column (str)
        values_c     - Optional  : length of printed most common values
    """    
    for key, dtype in df.dtypes.items():
        if str(dtype)[:5] == 'float':
            print(key)
            print('Min = {:<10}\tMax = {:<10}'.format(df[key].min(), df[key].max()))
            print(df.nlargest(n_outs,key)[key])
            print(df.nsmallest(n_outs,key)[key])
            #df['cena za 1 balenie v eur s dph'].hist(bins=100)
    #return df


def load_large_dta(fname, chunk_size = 1000000):
    # chunk size is the number of lines to be read at once
    import sys

    reader = pd.read_stata(fname, iterator=True)
    df = pd.DataFrame()
    
    try:
        chunk = reader.get_chunk(chunk_size)
        while len(chunk) > 0:
            df = df.append(chunk, ignore_index=True)
            chunk = reader.get_chunk(chunk_size)
            print('.')
            sys.stdout.flush()
    except (StopIteration, KeyboardInterrupt):
        pass

    print('\nloaded {} rows'.format(len(df)))

    return df


def weighted_median(df, val, weights = []):
    if not weights:
        return df[val].median()
    elif len(weights) == 1:
        df_sorted = df.sort_values(val)
        cumsum = df_sorted[weights[0]].fillna(1).cumsum()
        cutoff = df_sorted[weights[0]].fillna(1).sum() / 2.
        return df_sorted[cumsum >= cutoff][val].iloc[0]
    else:
        df_sorted = df.sort_values(val)
        df_sorted['new weight'] = df_sorted[weights[0]].fillna(1)
        for weight in weights[1:]:
            df_sorted['new weight'] = df_sorted['new weight'] * df_sorted[weight].fillna(1)
        cumsum = df_sorted['new weight'].cumsum()
        cutoff = df_sorted['new weight'].sum() / 2.
        return df_sorted[cumsum >= cutoff][val].iloc[0]


def fdf(df,kde =['nazov'], l=[],c = ['sukl','nazov','N mj v bal','N bal','C bal s dph','typ'], ret = False):
    loca = ~df['C bal s dph'].isna()
    for i, column in enumerate(kde):    
        for item in l[i]:
            if not item == '':
                if column == 'sukl':
                    new_loc = df[column].str.contains(unidecode(item).lower(),na=False)
                    sukls = new_loc.sum()
                    print('sukl found ', sukls)
                    if sukls > 0:
                        loca = (loca) & new_loc
                else:
                    #re.split('(\d+)',string) keby nahodou chcem splitovat numbers
                    loca = (loca) & df[column].str.contains(unidecode(item).lower(),na=False)
        
    if c == []:
        idk = df.loc[loca]
        idk['c_p_MJ'] = idk['C bal s dph'] / idk['N mj v bal'].fillna(1) / 1.1
        print(idk)
    else:
        idk = df.loc[loca][c]
        try:
            idk['c_p_MJ'] = idk['C bal s dph'] / idk['N mj v bal'].fillna(1) / 1.1
        except:
            pass
        print(idk)
    print('\npocet: ', len(idk))
    if len(idk) > 0:
        try:
            print('c_p_MJ bez DPH, plain median: {:.4f}'.format(weighted_median(idk, 'c_p_MJ', weights=[])))
            print('c_p_MJ bez DPH, na balenia median: {:.4f}'.format(weighted_median(idk, 'c_p_MJ', weights=['N bal'])))
            print('c_p_MJ bez DPH, na bal*MJ median: {:.4f}'.format(weighted_median(idk, 'c_p_MJ', weights=['N bal', 'N mj v bal'])))
        except:
            pass
    if ret:
        return loca
'''
fdf(df,kde = ['nazov'], l=[['']])
fdf(df,kde = ['sukl'], l=[['']])
fdf(df,kde = ['sukl','nazov'], l=[[''],['']])
fdf(df,kde = ['sukl','atc','nazov'], l=[[''],[''],['']])
fdf(df,kde = ['atc'], l=[['']])


    
    Oxid arzenitý 12 mg    L01XX27     1863D         balenie     16,00     5 257,2200         84 115,52     
    
    
'''


#Pantoprazol 40mg plv ifo    A02BC02     C39397         balenie     2 500,00     21,7700 

if __name__ == '__main__':
    '''['sukl', 'atc ozncenie ucinnej latky', 'nazov lieku', 'velkost mj',
       'pocet mj v jednom baleni', 'cena za mj v eur s dph',
       'cena za 1 balenie v eur s dph', 'pocet nakupenych mj',
       'pocet nakupenych baleni', 'celkovy nakup s dph', 'nemocnica',
       'lekaren', 'datum dodavky', 'dodavatel', 'poznamka', 'typ', 'file',
       'kod mz sr', 'kod tovaru', 'nazov tovaru', 'poskytovatel']
    '''
    '''['sukl', 'atc', 'nazov', 'velkost mj',
       'N mj v bal', 'C MJ s dph',
       'C bal s dph', 'N mj',
       'N bal', 'HZ s dph', 'nemocnica',
       'lekaren', 'datum', 'dodavatel', 'poznamka', 'typ', 'file',
       'kod mzsr', 'kod tovaru', 'poskytovatel']
    '''
    start = time.time()
    project_dir = 'C:\\Users\\klukaa\\Desktop\\praca\\_projects\\Benchmarking L SZM'
    DB_dir = 'C:\\Users\\klukaa\\Desktop\\praca\\_projects\\Benchmarking L SZM\\cleanDB'
    
    pd.set_option('max_colwidth', 65)
    pd.set_option('display.max_rows', 200)
    pd.set_option('display.max_columns', 200)
    pd.set_option('display.width', 320 )
    pd.options.mode.use_inf_as_na = True
    #pd.set_option('display.width', 1000)
    # Dont run again as it would replace file with already manually set pairs of clean and dirty xls files
    #prep_name_pair_file(DB_dir, '_names.txt')
    
    #print_xl_keys(DB_dir,'_keys.txt', 3, name_filter = [5,'Lieky'] )
    
    #df = concat_xl_files(DB_dir, name_filter = [5,'Lieky'] )
    #analyse_df(df, 'analyses.txt', group_by = 'typ', values_c = 20 )
    #save_df(df, 'dirtyDB.csv', path = None, replace = True)


    #df_lieky = load_df('zoznamy_liekov.xlsx', path = None)

    df = load_df('dirtyDB.csv', path = project_dir)
    df = clean_df(df, tolerance = 0.01)

    #get_lieky(df, 'lieky')
    #save_df(df, 'cleanDB.csv', path = None, replace = True)
    
    #save_df(df, 'dirtyDB.dta', path = None, replace = False)
    #save_df(df, 'dirtyDB.csv', path = None, replace = True)
    
    
    # JSON API pre SUKL nefunguje, shame
    #r = request_DB_liekov()
    
    #df = load_large_dta('D:\\_praca\\DB\\birthscut.dta')
    
    
    # Takto hladat konkretne polozky (HISTO TRAY ABC 72)
    #df[df['nazov tovaru'].str.contains("histo tray",na=False)][['nazov tovaru','pocet nakupenych baleni','cena za 1 balenie v eur s dph','celkovy nakup s dph']]
    #P31978
    print('{:.4f}s, run time'.format(time.time() - start))
    winsound.Beep(frequency, duration)





#    https://www.health.gov.sk/?zoznamy-uradne-urcenych-cien





''' DEAD CODE

and: &    or: |    not:~

    #df['pocet nakupenych mj'].astype(float)
    #df.loc[~df['Field1'].str.isdigit(), 'Field1'].tolist()
    #df.loc[~df['Field1'].astype(str).str.isdigit(), 'Field1'].tolist()
    #df.loc[~df['lekaren'].isnull()]['lekaren'].unique()
    #df.loc[~df['velkost mj'].isnull() & df['velkost mj'].astype(str).str.isalpha()]['velkost mj'].map(unidecode).str.lower().value_counts()

    #df.loc[~df['lekaren'].isnull()]['lekaren'].rename(unidecode,  inplace=True)
df.dtypes.keys()
    
    df.columns = df.columns.str.lower()
    for key, dtype in df.dtypes.items():
        pass

np.nan

df1 = pd.DataFrame({'A': ['A0', 'A1', 'A2', 'A3'],
                    'B': [0., np.nan, np.nan, np.nan],
                    'C': [-1, 2, np.nan, np.nan],
                    'D': ['D0', 'D1', 'D2', np.nan]},
                    index=[0, 1, 2, 3])

df1 = pd.DataFrame({'A': ['A0', 'A1', 'A2', 'A3'],
                    'B': ['B0', 'B1', 'B2', 'B3'],
                    'C': ['C0', 'C1', 'C2', 'C3'],
                    'D': ['D0', 'D1', 'D2', 'D3']},
                    index=[0, 1, 2, 3])

df2 = pd.DataFrame({'A': ['A4', 'A5', 'A6', 'A7'],
                    'B': ['B4', 'B5', 'B6', 'B7'],
                    'C': ['C4', 'C5', 'C6', 'C7'],
                    'D': ['D4', 'D5', 'D6', 'D7']},
                    index=[0, 1, 2, 3])

df4 = pd.DataFrame({'B': ['B2', 'B3', 'B6', 'B7'],
                    'D': ['D2', 'D3', 'D6', 'D7'],
                    'F': ['F2', 'F3', 'F6', 'F7']},
                    index=[2, 3, 6, 7])

print (df1,'\n', df2,'\n', df4)
print(df1.append([df2,df4], ignore_index = True))

'''