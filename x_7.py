print('--START-LIBRARIES-IMPORTING')
print('\tOS-LIB')
import requests
import os
import subprocess
import time
from io import StringIO, BytesIO
print('\tSELENIUM-LIB')
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager

print('\tPANDAS-LIB')

from pandas import DataFrame
from pandas import read_csv
import win32clipboard
import pyperclip
from PIL import Image

print('--END-LIBRARIES-IMPORTING')


def send_msg(element, category, slice, driver):
    msg = create_msg(slice)
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(msg, win32clipboard.CF_UNICODETEXT)
    win32clipboard.CloseClipboard()
    element.send_keys(Keys.CONTROL + 'v')
    time.sleep(0.5)

    def input_img(element, filename):
        image = Image.open(filename)

        output = BytesIO()
        image.convert("RGB").save(output, "BMP")
        data = output.getvalue()[14:]
        output.close()

        # data = image.tobitmap()

        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
        win32clipboard.CloseClipboard()

        element.send_keys(Keys.CONTROL + 'v')

    # --------------Закоментируй ниже--------------
    #images = [
    #    'Mix.jpg',
    #	 'octopusoil.jpg',
    #]
    #for img in images:
    #    input_img(element, img)
    #    time.sleep(1.5)
    # -----------Вплоть до этой строчки------------

    ActionChains(driver).key_down(Keys.ENTER).key_up(Keys.ENTER).perform()
    time.sleep(1.5)


def create_msg(slice):
    df = read_csv('category/cat1.csv', sep=';', encoding='utf-8')
    content = df[df.cat == slice.category].cont.iloc[0]
    return content.format(slice['name']).encode('unicode-escape').replace(b'\\\\u', b'\\u').replace(b'\\\\U', b'\\U').decode('unicode-escape')


def write(element, slice, driver):
    msg = create_msg(slice)
    send_msg(element, msg, slice, driver)


def main():
    filename = r'data.csv'
    df = read_csv(filename, sep=';', encoding='windows-1251')  # 'utf-8')
    
    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.get('https://web.whatsapp.com')

    input("Press Enter when u'll be ready : ")
    print("Starting...")

    search_field = driver.find_element_by_id(
        'side').find_element_by_xpath('//div[@data-tab=3]')

    for pos, elem in enumerate(df.number):
        search_field.send_keys(elem)
        time.sleep(1)

        search_field.send_keys(Keys.TAB)
        coords = driver.switch_to.active_element.location
        if coords['x'] > coords['y'] or coords == search_field.location:
            search_field.send_keys(Keys.CONTROL + 'a')
            search_field.send_keys(Keys.BACKSPACE)
            search_field.send_keys(Keys.BACKSPACE)
            search_field.send_keys(Keys.BACKSPACE)
            search_field.send_keys(Keys.BACKSPACE)
            search_field.send_keys(Keys.BACKSPACE)
            search_field.send_keys(Keys.BACKSPACE)
            search_field.send_keys(Keys.BACKSPACE)
            search_field.send_keys(Keys.BACKSPACE)
            search_field.send_keys(Keys.BACKSPACE)
            search_field.send_keys(Keys.BACKSPACE)
            search_field.send_keys(Keys.BACKSPACE)
            continue
        search_field.send_keys(Keys.ENTER)
        try:
            time.sleep(1)
            text_field = driver.find_element_by_tag_name(
                'footer').find_elements_by_xpath('//div[@contenteditable="true"]')[-1]
        except NoSuchElementException:
            print(pos, elem)
            search_field.send_keys(Keys.CONTROL + 'a')
            search_field.send_keys(Keys.BACKSPACE)
            search_field.send_keys(Keys.BACKSPACE)
            search_field.send_keys(Keys.BACKSPACE)
            search_field.send_keys(Keys.BACKSPACE)
            search_field.send_keys(Keys.BACKSPACE)
            search_field.send_keys(Keys.BACKSPACE)
            search_field.send_keys(Keys.BACKSPACE)
            search_field.send_keys(Keys.BACKSPACE)
            search_field.send_keys(Keys.BACKSPACE)
            search_field.send_keys(Keys.BACKSPACE)
            search_field.send_keys(Keys.BACKSPACE)
            continue
        slice = df.iloc[pos]
        write(text_field, slice, driver)
        search_field.send_keys(Keys.CONTROL + 'a')
        search_field.send_keys(Keys.DELETE)

    while True:
        input('Эта прога закрывается только через крестик')

main()
