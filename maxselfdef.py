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
    :param y: pd.Series x
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