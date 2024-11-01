from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from time import sleep
import openpyxl

#Настройка параметров селениума
option = webdriver.ChromeOptions()
option.add_argument('--headless')
option.add_argument("--log-level=3")
option.add_argument("--dns-prefetch-disable")
option.add_argument('--no-sandbox')
option.add_argument('--mute-audio')
option.add_argument("--ignore-certificate-errors")

#Инициализация данных пользовавтеля с Excel
bookq = openpyxl.open("loginandpassword.xlsx",read_only=True)
sheetq = bookq.active

book = openpyxl.open("sittings.xlsx",read_only=True)
sheet = book.active

#Функция захода в личный кабинет
def vhod():
    driver.get('https://kaspi.kz/mc/#/login')
    sleep(3)
    emailbutton = driver.find_element(By.XPATH, '//*[@id="app"]/div/div/div/div/form/div[1]/div/div/section/div/nav/ul/li[2]/a')
    emailbutton.click()
    sleep(0.8)
    email = driver.find_element(By.XPATH, '//*[@id="user_email"]')
    email.send_keys(sheetq["A1"].value)
    sleep(0.8)
    vhodbutton = driver.find_element(By.XPATH, '//*[@id="app"]/div/div/div/div/form/div[2]/div/button')
    vhodbutton.click()
    password = driver.find_element(By.XPATH, '//*[@id="app"]/div/div/div/div/form/div[1]/div/div/div[1]/input')
    password.send_keys(sheetq["B1"].value)
    sleep(0.5)
    vhodbutton2 = driver.find_element(By.XPATH, '//*[@id="app"]/div/div/div/div/form/div[2]/div/button')
    vhodbutton2.click()
    sleep(3)

#Заход программы на товар для получения данных с сайта
def start_pars(pars, minprice, stepprice, shopname):
    try:
        global itemname
        global minbool
        global onename
        global newpriceitem
        global oneprice
        driver.get(pars)
        sleep(3)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        itemname = driver.find_element(By.CLASS_NAME, 'item__heading').text
        onename = driver.find_element(By.XPATH, '//*[@id="offers"]/div/div/div[1]/table/tbody/tr[1]/td[1]/a')
        onename = onename.text
        search = driver.find_element(By.CLASS_NAME,'sellers-table__price-cell-text')
        search1 = search.text.replace(" ", "")
        search1 = search1[:-1]
        oneprice = int(search1)
        minbool = True
        if oneprice <= minprice:
            minbool = False
        elif oneprice > minprice:
            newpriceitem = oneprice - stepprice
    except NoSuchElementException:
        start_pars(pars, minprice, stepprice, shopname)

#Заход программы в личный кабинет для изменения цены в зависимости от цены на карточке товара
def newpricepast(newpricesite, shopname):
    try:
        driver.get(newpricesite)
        sleep(4)
        inputprice = driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div/div[2]/div/div[2]/div[2]/div/table/thead/tr[2]/th[2]/div/span/div[2]/input')
        sleep(0.5)
        inputprice.send_keys(Keys.CONTROL, 'a')
        sleep(0.5)
        inputprice.send_keys(Keys.DELETE)
        sleep(0.5)
        inputprice.send_keys(newpriceitem)
        print("Старая цена:", oneprice)
        print("Ставлю новую цену:", newpriceitem)
        sleep(0.5)
        buttonsave = driver.find_element(By.XPATH, '//*[@id="app"]/div/section/div/div[2]/div/div[2]/div[3]/button[1]')
        buttonsave.click()
        sleep(4)
        print("Я успешно изменил цену :)")
    except NoSuchElementException:
        newpricepast(newpricesite, shopname)

#Проверка на самую низкую цену на товар
def newprice(newpricesite, shopname):
    if onename != shopname:
        print("Меняю цену на:", itemname)
        newpricepast(newpricesite, shopname)
    elif onename == shopname:
        print("У вас самая низкая цена на", itemname)

#Проверка на минимально выставленную цену на товар
def mainwhile(pars, minprice, shopitemname, stepprice, shopname):
    start_pars(pars, minprice, stepprice, shopname)
    if minbool == True:
        newprice(F"https://kaspi.kz/mc/#/products/{shopitemname}", shopname)
    elif minbool == False:
        print("Цена достигла минимума на", itemname)
    else:
        pass


#Основная функция
def main():
    global driver
    while True:
        #Установка драйвера хром для парсинга
        driver = webdriver.Chrome(service=Service(ChromeDriverManager(driver_version="116.0.5845.111").install()), options=option)
        driver.set_window_size(1280, 720)
        vhod()
        #Цикл проходит по всем товарам написанные пользователем
        for row in range(1, sheet.max_row + 1):
            mainwhile(sheet[row][0].value, sheet[row][1].value, sheet[row][2].value, sheet[row][3].value, sheetq["C1"].value)
        driver.quit()
        driver = None
        itemname = None
        onename = None
        newpriceitem = None
        oneprice = None
        try:
            timesleep = sheetq["D1"].value
            print(f"Всё проверенно теперь я жду {timesleep} Секунд")
            sleep(timesleep)
        except KeyboardInterrupt:
            pass

if __name__ == "__main__":
    main()