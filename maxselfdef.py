''' author: Max Lin

'''

import time
from datetime import datetime
import os

def js_wait(driver):
    a = driver.find_element_by_xpath("//div[@class='appian-indicator-message']")
    while(a.get_attribute("aria-hidden") != 'true'):
        time.sleep(0.5)
        a = driver.find_element_by_xpath("//div[@class='appian-indicator-message']")
    time.sleep(1)

def digit_adj(number):
    number = str(number)
    if len(number) ==1:
        number = "0"+number
    return number

def time_format(filetype = 'xlsx'):
    '''return a time stamp postfix. Eg: _2018_10_01_09_34 AM.xlsx'''
    time_obj = datetime.now()
    if time_obj.hour>11:
        if time_obj.hour == 12:
            return f"_{time_obj.year}_{digit_adj(time_obj.month)}_{digit_adj(time_obj.day)}_12_{digit_adj(time_obj.minute)} PM.{filetype}"
        else:
            return f"_{time_obj.year}_{digit_adj(time_obj.month)}_{digit_adj(time_obj.day)}_{digit_adj(time_obj.hour%12)}_{digit_adj(time_obj.minute)} PM.{filetype}"
    else:
        return f"_{time_obj.year}_{digit_adj(time_obj.month)}_{digit_adj(time_obj.day)}_{digit_adj(time_obj.hour)}_{digit_adj(time_obj.minute)} AM.{filetype}"

def change_NO_to_hyperlink(text):
    '''change to text string as a excel internal hyperlink'''
    return f'=HYPERLINK("#{text}!A1","{text}")'

def apply(array,function):
    '''this is apply function for list'''
    for element in range(0,len(array)):
        array[element] = function(array[element])
    return array

def extract_text(react_text):
    '''this function is able to return text or react text from soup object. If it is none, return a space '''
    import re
    pattern = r'-->(?P<col>.*)<!--'
    if re.search(pattern,str(react_text)) is not None:
        return re.search(pattern,str(react_text)).group('col')
    elif react_text.text is not None:
        return react_text.text
    else:
        return " "

def check_environment_var(need_environment_var_list = ["chromedriver"]):
    '''This function return a environment variable dictionary'''
    environment_var_dict  = {}
    for var_key in need_environment_var_list:
        try:
            environment_var_dict[var_key] = os.environ[var_key]
        except:
            print(f"Please add environment variable for {var_key}, and run the APP")
            time.sleep(10)
            raise AttributeError
    return environment_var_dict

def crop(image_path, coords, saved_location):
    """
    @param image_path: The path to the image to edit
    @param coords: A tuple of x/y coordinates (x1, y1, x2, y2)
    @param saved_location: Path to save the cropped image
    """
    from PIL import Image
    image_obj = Image.open(image_path)
    cropped_image = image_obj.crop(coords)
    cropped_image.save(saved_location)
    #cropped_image.show()
def list_spliter(listObj,chunkSize=4):
    return [listObj[i:i + chunkSize] for i in range(0, len(listObj), chunkSize)]

def Maxcopyfolder(src, dst):
    from shutil import copyfile
    if os.path.isfile(src):
        try:
            copyfile(src, dst)
        except:
            print(f'{src} failed!')

    if os.path.isdir(src):
        if os.path.isdir(dst):
            shutil.rmtree(dst)
        os.makedirs(dst)
        files = os.listdir(src)
        for file in files:
            Maxcopyfolder(src=src + f'\\{file}', dst=dst + f'\\{file}')

def check_folder_exist(folderName):
    """
    :param folderName: the folder(directorary we are going to check whether exist)
    :return: None (if it does not exist, the def will create one automatically)
    """
    if not os.path.isdir(folderName):
        os.mkdir(folderName)

def copypaste_latest_file(Download_route,Dest_route,fileNameContains,postfix = ".xlsx"):
    '''
    :param Download_route: Download route
    :param Dest_route: Destination route
    :param fileNameContains: the part string of downloaded file name
    :param postfix: extension type profix (".xlsx" is default)
    :return: None (copy paste the lastest download file form download to target folder)
    '''
    from shutil import copyfile
    files = [file for file in os.listdir(Download_route) if file.find(fileNameContains) != -1]
    latest = files[0]
    for file in files:
        if os.path.getctime(Download_route+'\\'+file) > os.path.getctime(Download_route+'\\'+latest):
            latest = file
    copyfile(Download_route+"\\"+latest,Dest_route+'\\'+fileNameContains+postfix)

def find_latest_file(folders,filenamecontains):
    '''
    :param folders: the folders we are going to search
    :param filenamecontains: part of string file name contains
    :return: the full route of the latest file
    '''
    if type(folders) ==str:
        folders = [folders]
    latest_folder = folders[0]
    while(len([file for file in os.listdir(latest_folder) if file.find(filenamecontains) != -1])==0): ## if file is not in the first folder then remove the folder and assign to next folder until find it
        folders.pop(0)
        if len(folders)==0:
            raise FileNotFoundError(f'We can not find file: {filenamecontains}')
        latest_folder = folders[0]
    latest = [file for file in os.listdir(latest_folder) if file.find(filenamecontains) != -1][0] ## define the first folder and first file in that folder as latest file
    for folder in folders:
        files = [file for file in os.listdir(folder) if file.find(filenamecontains) != -1]
        for file in files:
            if os.path.getctime(folder+'\\'+file) > os.path.getctime(latest_folder+'\\'+latest):
                latest = file
                latest_folder = folder
    return latest_folder +'\\'+ latest


def similar_compare(x,y,comprise=True,threshold = 0.88):
    '''

    :param x: string
    :param y: string
    :param comprise:  x comprise y or x comprise x will return True
    :param threshold: similarity percentage threshold
    :return: Boolean
    '''
    from difflib import SequenceMatcher
    def simliar(a,b):
        return SequenceMatcher(None,a,b).ratio()>threshold
    flag = False
    flag = (((x in y) or (y in x)) and comprise) or simliar(x,y)
    return (flag)

def similar_mapping_key_dictionary(x, y, comprise=False, threshold=0.98):
    '''
    :param x: pd.Series x
    :param y: pd.Series y
    :param comprise:  x comprise y or x comprise x will think they are matching key
    :param threshold: similarity percentage threshold
    :return: dictionary of mapping key like keyDictionary[x] = y
    Note y must contain all items of x
    '''
    import pandas as pd
    x = pd.Series(x.unique()).copy()
    y = pd.Series(y.unique()).copy()
    x.sort_values(inplace=True)
    y.sort_values(inplace=True)
    dictionary = {}
    for i in range(x.shape[0]):
        for j in range(y.shape[0]):
            if similar_compare(x.iloc[i], y=y.iloc[j], comprise=comprise, threshold=threshold):
                dictionary[x.iloc[i]] = y.iloc[j]
                break

        # dictionary[x[x.apply(lambda x: similar_compare(x,y=y.iloc[i],comprise=comprise,threshold= threshold))].values[0]] = y.iloc[i]
    return dictionary

def to_universe_date_format(PdSeries, errors='coerce',to_string = True, NaT_replace = ' '):
    '''

    :param PdSeries: pandas Series (str or object type)
    :param errors: if we encounter errors, what should we do
    :param to_string:  Return str or pd time series
    :param NaT_replace: if we choose str, what should we replace NaT
    :return: pd Series with universal data format
    '''
    import pandas as pd
    PdSeries = pd.to_datetime(PdSeries,errors)
    if to_string:
        PdSeries = PdSeries.astype('str')
        PdSeries = PdSeries.str.replace('NaT',NaT_replace)
    return PdSeries

def to_traditional_date_format_string(string):
    import re
    research_result = re.search(r'(?P<year>\d{4}?)-(?P<month>\d{2}?)-(?P<day>\d{2}?)(?P<other>.*)',string)
    if research_result==None:
        return ' '
    else:
        return research_result.group('month')+'/'+research_result.group('day')+'/'+research_result.group('year')+research_result.group('other')

def to_traditional_date_format(PdSeries,NaT_replace = ' '):
    import pandas as pd
    PdSeries = PdSeries.astype('str')
    return PdSeries.apply(to_traditional_date_format_string)

def linefeed_remove_strip(PdSeries):
    '''
    replace '\n' and apply strip() on a string pdSeries
    '''
    return PdSeries.str.replace('\n',' ').str.strip()

def getwebdriver(url=""):
    # Start the chrome driver with selenium
    # https://stackoverflow.com/questions/43079018/selenium-chromedriver-failed-to-load-extension
    capabilities = {'chromeOptions': {'useAutomationExtension': False}}
    driver = webdriver.Chrome(os.environ['chromedriver'], desired_capabilities=capabilities)
    driver.maximize_window()
    if url != "":
        driver.get(url)
    return driver
