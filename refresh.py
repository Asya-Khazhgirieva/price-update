#Код для обновления прайсо нескольких вендоров. В рабочем скрипте более 25 вендоров и 10 дистрибьюторов

#Импорт модулей
#работа с Windows
import os, sys
sys.coinit_flags = 0
from os import system
import win32com.client
from win32com.client import Dispatch
import win32clipboard #чистка кэша
import shutil #работа с папками и файлами Windows
import pywinauto
from pywinauto.application import Application #для Outlook
import win32com.client as win32

#для zip-архивов
import zipfile

#Курс ЦБ РФ
import pycbrf
from pycbrf.toolbox import ExchangeRates

#дата и время
import datetime
from datetime import datetime
from datetime import *
import time

#чтение Excel
import xlrd
import pyexcel as p #конвертация .xls в .xlsx
import xlsxwriter as xl #конвертация .csv в .xlsx
import csv #конвертация .csv в .xlsx
import linecache #работа со строками

#работа в браузере
from selenium import webdriver 
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

#логирование процессов
import logging
import requests
#Импорт модулей

#указание параметров логирования


#Обновление курса валют
dataQuery = datetime.now().strftime("%Y-%m-%d")
rates = ExchangeRates(dataQuery)
print('Получение курса валют на ' + dataQuery)
rates['USD'].name
rates['EUR'].name
print('Курс USD:')
print(rates['USD'].value)
print('Курс EUR:')
print(rates['EUR'].value)

#Тайминг начала кода
start_time = datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S")
now = datetime.now()
print('Обновление запущено: ' + start_time)

#указание параметров окна браузера
options = Options()
options.add_argument("--start-maximized") #полноэкранный режим окна
print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Chrome>: Открытие браузера')
chromedriver = Service("C:\\Users\user\chromedriver\chromedriver.exe")
#Если скрипт крашнулся на этом этапе, то скорее всего нужно просто обновить драйвер.
driver = webdriver.Chrome(service=chromedriver, options=options)

print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Отладка>: удаление кэша gen_py')
shutil.rmtree('C:/Users/user/AppData/Local/Temp/gen_py')
#данная папка препятствует нормальной работе буфера обмена, в том числе функций Range и PasteSpecial в Excel.

print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Outlook>: Создание COM-объекта')
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

print('Открываем файл для записи')
f = open('C:/Users/user/Desktop/log.txt', 'a', encoding='utf-8')

#Обновление стока 
print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Ноутбуки>: Обновление стока')
print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Входящие>: Поиск папки Ноутбуки')
inbox = outlook.Folders["user@domain.ru"].Folders["Входящие"].Folders["Ноутбуки"]
print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Ноутбуки>: Получение писем')
messages = inbox.Items
print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Ноутбуки>: Просмотр последнего письма')
message = messages.GetLast()
print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Ноутбуки>: Проверка даты письма')
date = message.SentOn.strftime("%d.%m.%Y")
yest = datetime.now() - timedelta(days=1)
yest = yest.strftime("%d.%m.%Y")
print(yest)
print(message.SentOn.strftime("%d.%m.%Y"))
try:
    if message.SentOn.strftime("%d.%m.%Y") == yest:
        print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Ноутбуки>: Получение вложений письма')
        attachments = message.Attachments
        attachment = attachments.Item(1)
        print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Ноутбуки>: Сохранение вложения Notebook_price.zip в Desktop')
        src = 'C:/Users/user/Desktop/'
        filename = 'Notebook_price.zip'
        attachment.SaveASFile(src + filename)
        print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Ноутбуки>: Извлечение содержимого архива Notebook_price.zip в Desktop')
        z = zipfile.ZipFile('C:/Users/user/Desktop/Notebook_price_OCS.zip', 'r')
        z.extractall('C:/Users/user/Desktop/')
        print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Ноутбуки>: Создание COM-объекта Excel')
        Notebook = win32com.client.DispatchEx('Excel.Application')
        Notebook.Visible = 1
        print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Ноутбуки>: Назначение имени файла Notebook_price.xlsx')
        list_files=list()
        for name in z.namelist(): 
            list_files.append(name) 
        wb = Notebook.Workbooks.Open('C:/Users/user/Desktop/' + name, None, True)
        print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Ноутбуки>: Сохранение Notebook_price.xlsx в Склады')
        wb.SaveCopyAs('C:/Users/user/Desktop/!_GPLи для компани/Notebook_price.xlsx')
        print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Ноутбуки>: Завершение')
        wb.Close(False)
        Notebook.Quit()
        z.close()
except Exception as err:
    print(err)
    err = str(err)
    f.write(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <OCS Ноутбуки>: Обновление стока' + '-' + err + '\n')
    pass

#Скачивание стока
try:
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <вендор>: Переход на сайт поставщика')
    driver.get('https://вендор.ru/auth/')
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <вендор>: Ввод логина и пароля')
    login = driver.find_element(By.NAME, 'USER_LOGIN')
    password = driver.find_element(By.NAME, 'USER_PASSWORD')
    sign = driver.find_element(By.XPATH, '/html/body/main/section/div/form/button')
    login.send_keys('login')
    password.send_keys('password')
    sign.click()
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <вендор>: Скачивание вендор.xlsx')
    driver.find_element(By.CLASS_NAME, "header__user-name").click()
    driver.find_element(By.CLASS_NAME, "js-downloadPrice").click()
    time.sleep(5)
except Exception as err:
    print(err)
    err = str(err)
    f.write(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <вендор>: Обновление стока' + '-' + err + '\n')
    pass

#Открытие прайс-листа 1
try:
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <PriceRUVendor.xlsx>: Создание COM-объекта Excel')
    Vendor = win32com.client.Dispatch('Excel.Application')
    Vendor.Visible = 1
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <PriceRUVendor.xlsx>: Принудительное обновление связей')
    wb = Vendor.Workbooks.Open('C:/Users/user/Desktop/PriceRU/PriceRUVendor.xlsx')
    wb.RefreshAll()
    sheet = wb.ActiveSheet
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <PriceRUVendor.xlsx>: Обновление валюты [USD]')
    sheet.Cells(1,23).value = rates['USD'].value
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <PriceRUVendor.xlsx>: Сохранение')
    wb.Save()
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <PriceRUVendor.xlsx>: Получение данных из ячеек [E:T]')
    sheet.Columns('E:T').Copy()
except Exception as err:
    print(err)
    err = str(err)
    f.write(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <PriceRUVendor.xlsx>: Открытие прайс-листа Vendor' + '-' + err + '\n')
    pass
time.sleep(1)
#Открытие файла загрузки Vendor
try:
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <4533Vendor.xls>: Создание COM-объекта Excel')
    Vendor1 = win32com.client.Dispatch('Excel.Application')
    Vendor1.Visible = 1
    wb1 = Vendor1.Workbooks.Open('C:/Users/user/Desktop/!_Заливка на сайт/4533Vendor.xls')
    sheet = wb1.ActiveSheet
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <4533Vendor.xls>: Обновление данных в ячейках [A:P]')
    sheet.Columns('A:P').PasteSpecial(Paste=-4163) #4163 = "вставить только значения"
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <4533Vendor.xls>: Сохранение')
    wb1.Save()
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <PriceRUVendor.xlsx>: Завершение')
    wb.Close(True)
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <4533Vendor.xls>: Завершение')
    wb1.Close(True)
    Vendor.Quit()
    Vendor1.Quit()
except Exception as err:
    print(err)
    err = str(err)
    f.write(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <4533Vendor.xls>: Открытие файла загрузки Vendor' + '-' + err + '\n')
    wb.Close(True)
    wb1.Close(True)
    Vendor.Quit()
    Vendor1.Quit()
    #logging.info('<4533Vendor.xls>')
    pass
    
shutil.copy('C:/Users/user/Desktop/!_Заливка на сайт/4533Vendor.xls','C:/Users/user/Desktop/!_Заливка на сайт/4555Vendor.xls')

try:
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.CloseClipboard()
except Exception as err:
    print(err)
    pass

#Обновление стока Dist
try:
    src = 'C:/Users/user/Downloads/dealerD.zip'
    dst = 'C:/Users/user/Desktop/dealerD.zip'
    shutil.move(src, dst)
    print('<Dist>: Извлечение содержимого архива dealerD.zip в Desktop')
    z = zipfile.ZipFile('C:/Users/user/Desktop/dealerD.zip', 'r')
    z.extractall('C:/Users/user/Desktop/')
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Dist>: Создание COM-объекта Excel')
    Dist = win32com.client.Dispatch('Excel.Application')
    Dist.Visible = 1
    wb = Dist.Workbooks.Open('C:/Users/user/Desktop/DealerD.xlsx')
    sheet = wb.ActiveSheet
    sheet.Rows.EntireRow.Hidden=False
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Dist>: Добавление столбца A:A')
    rangeObj = sheet.Range('A1:A2')
    rangeObj.Value = ['a', 'b']
    rangeObj = sheet.Range('A1:A2')
    rangeObj.Value = [1,2]
    rangeObj.EntireColumn.Insert()
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Dist>: Добавление столбца B:B')
    rangeObj = sheet.Range('A1:A2')
    rangeObj.Value = ['a', 'b']
    rangeObj = sheet.Range('B1:B2')
    rangeObj.Value = [1,2]
    rangeObj.EntireColumn.Insert()
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Dist>: Получение данных из столбца F')
    sheet.Columns('E').Copy()
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Dist>: Добавление данных из столбца F')
    sheet.Columns('A').PasteSpecial()
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Dist>: Добавление формулы в ячейку B1')
    sheet.Cells(1,2).Value = 'f=ЕСЛИ(C1="+";"+";ЕСЛИ(D1="+";"+";0))'
    sheet.Cells(1,2).Replace('f=', '=')
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Dist>: Автозаполнение данных в столбце B')
    sheet.Cells(1,2).Copy()
    #Подсчет заполненных ячеек в столбце G
    op_cell = Dist.WorksheetFunction.CountA(Dist.Range("G:G"))
    op_cell = int(op_cell) + 13
    #Продлить формулу в столбце B до конца столбца G
    ranfil = "B1:B" + str(op_cell)
    Dist.Range("B1").AutoFill(Dist.Range(ranfil), 0)
    #sheet.Columns('B').PasteSpecial()
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Dist>: Сохранение Dist.xlsx в Склады')
    wb.SaveCopyAs('C:/Users/user/Desktop/Склады/Dist.xlsx')
    wb.Close(False)
    Dist.Quit()
    z.close()
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Dist>: Удаление временной копии dealerD.xlsx в Desktop')
    os.remove('C:/Users/user/Desktop/DealerD.xlsx')
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Dist>: Удаление временной копии dealerD.zip в Desktop')
    os.remove('C:/Users/user/Desktop/dealerD.zip')
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Dist>: Завершение')
except Exception as err:
    print(err)
    err = str(err)
    f.write(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Dist>' + '-' + err + '\n')
    

#Скачивание стока Dist2
try:
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Dist2>: Переход на сайт поставщика')
    driver.execute_script('''window.open("https://ecom.Dist2.ru/login","_blank");''') #Открывает ссылку в новой вкладке. Далее драйвер возвращается к старой вкладке для продолжения работы.
    driver.switch_to.window(driver.window_handles[-1]) #переключение на 2 вкладку
    time.sleep(2)
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Dist2>: Ввод логина и пароля')
    login = driver.find_element(By.ID, 'mail')
    login.send_keys('login')
    password = driver.find_element(By.ID, 'password')
    password.send_keys('password')
    sign = driver.find_element(By.XPATH, '/html/body/div/div[1]/div/div[2]/form/div[4]/button')
    sign.click()
    time.sleep(5)
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Dist2>: Скачивание Dist2.xlsx')
    try:
        price = driver.find_element(By.XPATH, '/html/body/div/aside/div[2]/div[2]/div[1]').click()
        elementToFocus = driver.find_element(By.XPATH, "/html/body/div/aside/div[3]/div[2]")
        driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", elementToFocus)
        time.sleep(3)
        price = driver.find_element(By.XPATH, '/html/body/div/aside/div[3]/div[2]/div[4]/span/span').click()
    except Exception as err:
        print(err)
        price = driver.find_element(By.LINK_TEXT, "Каталог").click()
        elementToFocus = driver.find_element(By.XPATH, "/html/body/div/aside/div[3]/div[2]") #Делаю div, который нужно проскроллить, активным
        driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", elementToFocus) #Скролл дива
        time.sleep(3)
        price = driver.find_element(By.LINK_TEXT, "Скачать прайс-лист").click()
except Exception as err:
    print(err)
    err = str(err)
    f.write(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Dist2>' + '-' + err + '\n')

#Тайминг
end_time = datetime.strftime(datetime.now(), "%Y.%m.%d %H:%M:%S")
print('Обновление прайс-листов завершено ' + end_time)
start_time = datetime.strftime(datetime.now(), "%Y.%m.%d %H:%M:%S")
print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S ") + 'Загрузка прайсов в NetCat site ' + start_time)

#Импорт скрипта загрузки прайсов
import uploadct

print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Яндекс.Вебмастер>: Обновление webmaster.xml')
try:
    driver.execute_script('''window.open("https://site.ru/netcat/modules/netshop/export/webmaster_build_xml.php?key=*****","_blank");''')
    time.sleep(2)
    driver.get('https://site.ru/netcat/modules/netshop/export/webmaster_build_xml.php?key=*****')
    time.sleep(2)
except Exception as err:
    print(err)
    err = str(err)
    f.write(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Яндекс.Вебмастер>' + '-' + err + '\n')

#Тайминг
end_time = datetime.strftime(datetime.now(), "%Y.%m.%d %H:%M:%S")
print('Обновление полностью завершено: ' + end_time)
try:
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'user@domain.ru'
    mail.Subject = 'Обновление полностью завершено: ' + end_time
    mail.Body = 'Скрипт отработал успешно'
    mail.Send()
except Exception as err:
    print(err)
    err = str(err)
    f.write(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' Сообщение на почту о завершении обновления' + '-' + err + '\n')

#Конец кода
f.write('Обновление полностью завершено: ' + end_time + '\n' + '\n')
f.close()
driver.quit()
system("taskkill /f /im py.exe")
system("taskkill /f /im chromedriver.exe")
system("taskkill /f /im excel.exe")

sys.exit(0)
quit()
