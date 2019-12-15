# Note: For proper working of this Script Good and Uninterepted Internet Connection is Required
# Keep all contacts unique
# Can save contact with their phone Number

# Import required packages
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import datetime
import time
import openpyxl as excel
import requests
import json
from bitlyshortener import Shortener


signos = ["aquario", "peixes", "aries", "touro", "gemeos", "cancer", "leao", "virgem", "libra", "escorpiao",
          "sagitario", "capricornio"]

moedas = ["USD", "EUR", "GBP", "ARS", "BTC"]

bolsas = ["IBOVESPA", "NASDAQ"]

stringFinal = " *BOM DIA FLOR DO DIA* \n"

stringFinal += "&"


def encurtarLink(url):
    tokens_pool = ['4e37257b19b91aed295cc013e995a2a79f0a3b2a']
    shortener = Shortener(tokens=tokens_pool, max_cache_size=8192)
    urls = []
    urls.append(url)

    return shortener.shorten_urls(urls)[0]



def getFinancas():
    cot = requests.get("https://api.hgbrasil.com/finance")
    c = json.loads(cot.text)
    texto = ""
    lista = c["results"]["currencies"]
    for val in moedas:
        info = lista[val]
        texto += info["name"] + ": R$" + str(info["buy"]).replace(".", ",") + " (" + str(info["variation"]) + "%)\n"

    lista = c["results"]["stocks"]
    for val in bolsas:
        info = lista[val]
        nome = ""
        if "BOVESPA" in info["name"]:
            nome = "Bovespa"
        else:
            nome = "NASDAQ"

        texto += nome + ": " + str(info["points"]).replace(".", ",") + " pontos (" + str(
            info["variation"]) + "%)\n"

    tax = requests.get("https://api.hgbrasil.com/finance/taxes?key=46198f46")
    t = json.loads(tax.text)
    taxas = t["results"][0]

    texto += "CDI: " + str(taxas["cdi"]).replace(".", ",") + "%\n"
    texto += "Selic: " + str(taxas["selic"]).replace(".", ",") + "%\n"

    return texto


def getSignos():
    texto = ""
    for val in signos:
        req = requests.get("http://babi.hefesto.io/signo/" + val + "/dia")
        d = json.loads(req.text)
        texto += "*" + str(d["signo"]).capitalize() + "*: " + str(d["texto"]).replace("      ", "") + "\n"
    return texto


def getNoticias():
    texto = ""
    req = requests.get(
        "https://newsapi.org/v2/top-headlines?sources=google-news-br&apiKey=122a4be6cc4648f0988aaf154abd18e3")
    noticias = json.loads(req.text)
    artigos = noticias["articles"]

    for val in artigos:
        texto += "" + val["title"] + "\n" + "_Leia mais em: " + encurtarLink(val["url"]) + "_\n"

    return texto


def getClima():
    texto = ""
    cli = requests.get(
        "https://api.hgbrasil.com/weather?fields=only_results,temp,description,max,city_name,min,humidity,sunrise,sunset,date&key=46198f46&city_name=Prudente,SP")
    c = json.loads(cli.text)

    texto += "Cidade: " + c["city_name"] + "\n"
    texto += "Temperatura: " + str(c["temp"]) + " °C\n"
    texto += "Umidade: " + str(c["humidity"]) + "%\n"
    texto += "Tempo: " + c["description"] + "\n"
    texto += "Nascer do sol: " + c["sunrise"] + "\n"
    texto += "Por do sol: " + c["sunset"] + "\n"

    return texto


# function to read contacts from a text file
def readContacts(fileName):
    lst = []
    file = excel.load_workbook(fileName)
    sheet = file.active
    firstCol = sheet['A']
    for cell in range(len(firstCol)):
        contact = str(firstCol[cell].value)
        contact = "\"" + contact + "\""
        lst.append(contact)
    return lst

# Target Contacts, keep them in double colons
# Not tested on Broadcast
targets = readContacts("contacts.xlsx")

# can comment out below line
print(targets)

# Driver to open a browser
chrome = '/usr/local/bin/chromedriver'
driver = webdriver.Chrome(chrome)

# link to open a site
driver.get("https://web.whatsapp.com/")

# 10 sec wait time to load, if good internet connection is not good then increase the time
# units in seconds
# note this time is being used below also
wait = WebDriverWait(driver, 10)
wait5 = WebDriverWait(driver, 5)
input("Scan the QR code and then press Enter")

# Message to send list
# 1st Parameter: Hours in 0-23
# 2nd Parameter: Minutes
# 3rd Parameter: Seconds (Keep it Zero)
# 4th Parameter: Message to send at a particular time
# Put '\n' at the end of the message, it is identified as Enter Key
# Else uncomment Keys.Enter in the last step if you dont want to use '\n'
# Keep a nice gap between successive messages
# Use Keys.SHIFT + Keys.ENTER to give a new line effect in your Message
msgToSend = [
    [15, 39, 0, stringFinal]
]

# Count variable to identify the number of messages to be sent
count = 0
while 1 == 1:

    # Identify time
    curTime = datetime.datetime.now()
    curHour = curTime.time().hour
    curMin = curTime.time().minute
    curSec = curTime.time().second

    # if time matches then move further
    if msgToSend[count][0] == curHour and msgToSend[count][1] == curMin and msgToSend[count][2] == curSec:
        # utility variables to tract count of success and fails
        success = 0
        sNo = 1
        failList = []

        stringFinal += "*Cotações:*\n"
        stringFinal += getFinancas()
        stringFinal += "&"
        stringFinal += "*Previsão do tempo*\n"
        stringFinal += getClima()
        stringFinal += "&"
        stringFinal += "*Notícias:*\n"
        stringFinal += getNoticias()
        stringFinal += "&"
        stringFinal += "*Horóscopo:*\n"
        stringFinal += getSignos()

        msgToSend[count][3] = stringFinal

        # Iterate over selected contacts
        for target in targets:
            print(sNo, ". Target is: " + target)
            sNo += 1
            try:
                # Select the target
                x_arg = '//span[contains(@title,' + target + ')]'
                try:
                    wait5.until(EC.presence_of_element_located((
                        By.XPATH, x_arg
                    )))
                except:
                    # If contact not found, then search for it
                    searBoxPath = '//*[@id="input-chatlist-search"]'
                    wait5.until(EC.presence_of_element_located((
                        By.ID, "input-chatlist-search"
                    )))
                    inputSearchBox = driver.find_element_by_id("input-chatlist-search")
                    time.sleep(0.5)
                    # click the search button
                    driver.find_element_by_xpath('/html/body/div/div/div/div[2]/div/div[2]/div/button').click()
                    time.sleep(1)
                    inputSearchBox.clear()
                    inputSearchBox.send_keys(target[1:len(target) - 1])
                    print('Target Searched')
                    # Increase the time if searching a contact is taking a long time
                    time.sleep(4)

                # Select the target
                driver.find_element_by_xpath(x_arg).click()
                print("Target Successfully Selected")
                time.sleep(2)

                # Select the Input Box
                inp_xpath = "//div[@contenteditable='true']"
                input_box = wait.until(EC.presence_of_element_located((
                    By.XPATH, inp_xpath)))
                time.sleep(1)

                # Send message and break lines
                for part in msgToSend[count][3].split('\n'):
                    breakManyLines = 0

                    if "&" in part:
                        breakManyLines = 1
                        part = part.replace("&", "")

                    if breakManyLines == 1:
                        ActionChains(driver).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.SHIFT).key_up(
                            Keys.ENTER).perform()
                        ActionChains(driver).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.SHIFT).key_up(
                            Keys.ENTER).perform()
                        input_box.send_keys(part)
                        ActionChains(driver).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.SHIFT).key_up(
                            Keys.ENTER).perform()
                    else:
                        input_box.send_keys(part)
                        ActionChains(driver).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.SHIFT).key_up(
                            Keys.ENTER).perform()

                # Link Preview Time, Reduce this time, if internet connection is Good
                time.sleep(1)
                input_box.send_keys(Keys.ENTER)
                print("Successfully Send Message to : " + target + '\n')
                success += 1
                time.sleep(0.5)

            except:
                # If target Not found Add it to the failed List
                print("Cannot find Target: " + target)
                failList.append(target)
                pass

        print("\nSuccessfully Sent to: ", success)
        print("Failed to Sent to: ", len(failList))
        print(failList)
        print('\n\n')
driver.quit()
