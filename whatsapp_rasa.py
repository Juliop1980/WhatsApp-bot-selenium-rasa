# Note: For proper working of this Script Good and Uninterepted Internet Connection is Required
# Keep all contacts unique
# Can save contact with their phone Number

# Import required packages
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import datetime
import time
import openpyxl as excel
import requests
from datetime import datetime, timedelta


driver = webdriver.Firefox()
wait = WebDriverWait(driver, 150)
wait5 = WebDriverWait(driver, 999999999)

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

# Función que recibe texto y lo envía al respectivo contacto de whatsapp
def to_whatsapp(txt, contacto):
    select_contacto(contacto)
    # Select the Input Box
    inp_xpath = "//div[(@contenteditable='true') and (@data-tab='1')]"
    input_box = wait.until(EC.presence_of_element_located((
        By.XPATH, inp_xpath)))
    #time.sleep(1)

    # Send message
    # contacto is your target Name and msgToSend is you message
    input_box.send_keys(txt) # + Keys.ENTER (Uncomment it if your msg doesnt contain '\n')
    # Link Preview Time, Reduce this time, if internet connection is Good
    #time.sleep(10)
    input_box.send_keys(Keys.ENTER)
    print("Successfully Send Message to : "+ contacto + '\n')
    #success+=1
    #time.sleep(0.5)
    return True

# Función que envia un mensaje a rasa y retorna el string correspondiente a la salida de rasa
def to_rasa(txt):

    sender = "whatsapp"
    bot_message = ""
    print("Sending message now...")
    r = requests.post('http://localhost:5002/webhooks/rest/webhook', json={"sender": sender, "message": txt})
    for i in r.json():
        bot_message = bot_message + i['text']
    #clear_whatsapp_chat
    return bot_message

def get_last_whatsapp_message():
    #select_contacto(contacto)
    #clear_whatsapp_chat()
    first_message_glance = (driver.find_element_by_class_name("_274yw").text).splitlines()
    return first_message_glance[0]

def get_last_whatsapp_time():
    first_message_glance = (driver.find_element_by_class_name("_274yw").text).splitlines()
    return first_message_glance[1]


def clear_whatsapp_chat(contacto):
    select_contacto(contacto)
    #Selecciona menu para luego presionar boton para borrar chat
    menu_xpath = "(//div[(@role='button') and (@title='Menu')])[2]"
    menu_button = wait5.until(EC.presence_of_element_located((
        By.XPATH, menu_xpath)))
    #time.sleep(1)
    menu_button.click()
    #time.sleep(10)
    #seleccionar boton para borrar chat
    borrar_xpath = "//div[(@class='Ut_N0 n-CQr') and (@role='button') and (@title='Delete chat')]"
    #driver.find_element_by_xpath("//div[(@class='Ut_N0 n-CQr') and (@role='button') and (@title='Delete chat')]").click()
    borrar_button = wait5.until(EC.presence_of_element_located((
        By.XPATH, borrar_xpath)))
    #time.sleep(1)
    borrar_button.click()
    #time.sleep(10)
    #confirmar borrar
    confirmar_borrar_xpath = "//div[(@class='S7_rT FV2Qy') and (@role='button')]"
    #driver.find_element_by_xpath("//div[(@class='S7_rT FV2Qy') and (@role='button')]").click()
    confirmar_borrar_button = wait5.until(EC.presence_of_element_located((
        By.XPATH, confirmar_borrar_xpath)))
    #time.sleep(1)
    confirmar_borrar_button.click()
    return True



# selecciona el contacto en Whatsapp Web
def select_contacto(contacto):
    # Select the target
    x_arg = '//span[contains(@title,' + contacto + ')]'

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

    driver.find_element_by_xpath(x_arg).click()
    print("Target Successfully Selected")
    time.sleep(2)


def es_reciente(whatsapp_time):
    whatsapp_datetime = whatsapp_time_string_to_datetime(whatsapp_time)
    now = datetime.now()
    print(whatsapp_datetime)
    difference = (now - whatsapp_datetime)
    total_seconds = difference.total_seconds()
    if total_seconds <= 3600:
        #print("True")
        return True
    else:
        #print("False")
        return False
    # aux_whatsapp_time = now.strftime("%Y-%m-%d")()
    # whatsapp_hour= (whatsapp_time.split(":"))[0]
    # time = (whatsapp_time.split(" "))[1]
    # time_number =
    # if time == "PM":


    #print(whatsapp_hour)


def whatsapp_time_string_to_datetime(time_string):
    #now = datetime.now()
    #aux_whatsapp_time = now.strftime("%Y-%m-%d")
    aux_whatsapp_time = whatsapp_datetime_string()
    #print(aux_whatsapp_time)
    whatsapp_hour = (time_string.split(":"))[0]
    time = (time_string.split(" "))[1]
    whatsapp_hour_number = int(whatsapp_hour)
    if time == "PM":
        whatsapp_hour = str(whatsapp_hour_number+12)
    else:
        whatsapp_hour= "0"+whatsapp_hour

    aux_whatsapp_minutes = (((time_string.split(":"))[1]).split(" "))[0]

    aux_whatsapp_time= aux_whatsapp_time + " " + whatsapp_hour + ":"+ aux_whatsapp_minutes + ":00"

    datetime_object = datetime.strptime(aux_whatsapp_time, '%Y-%m-%d %H:%M:%S')
    return datetime_object

# devuelve la ultima fecha de whatsapp en un string. Si la fecha es anterior al dia anterior devuelve None
def whatsapp_datetime_string():
    date_xpath = "//span[@class='_3Whw5']"
    #driver.find_element_by_xpath(date_xpath).text
    whatsapp_date = driver.find_element_by_xpath(date_xpath).text
    datetime_string = whatsapp_date.rstrip()
    #whatsapp_datetime_to_datetime_string(datetime_string)
    #print ("hola")
    now = datetime.now()
    #print(datetime_string)
    if datetime_string == "TODAY":
        #print(now.strftime("%Y-%m-%d"))
        return now.strftime("%Y-%m-%d")
    elif datetime_string == "YESTERDAY":
        days = timedelta(1)
        new_date = now - days
        return new_date.strftime("%Y-%m-%d")
    else:
        days = timedelta(3)
        new_date = now - days
        return new_date.strftime("%Y-%m-%d")

    #aux_whatsapp_time = now.strftime("%Y-%m-%d")


# devuelve la ultima fecha de whatsapp segun el string de whatsapp
# def whatsapp_datetime_to_datetime_string(date_time_string):
    
#     return date_string















if __name__ == '__main__':

    # Target Contacts, keep them in double colons
    # Not tested on Broadcast
    targets = readContacts("contacts.xlsx")

    # can comment out below line
    print(targets)

    # Driver to open a browser
    driver = webdriver.Firefox()

    #link to open a site
    driver.get("https://web.whatsapp.com/")

# 10 sec wait time to load, if good internet connection is not good then increase the time
# units in seconds
# note this time is being used below also
    wait = WebDriverWait(driver, 15)
    wait5 = WebDriverWait(driver, 5)
    input("Scan the QR code and then press Enter")

    while True:
        for i in targets:
            select_contacto(i)
            if (es_reciente(get_last_whatsapp_time())):
                to_whatsapp(to_rasa(get_last_whatsapp_message()),i)
            clear_whatsapp_chat(i)
        time.sleep(15)




            


# Message to send list
# 1st Parameter: Hours in 0-23
# 2nd Parameter: Minutes
# 3rd Parameter: Seconds (Keep it Zero)
# 4th Parameter: Message to send at a particular time
# Put '\n' at the end of the message, it is identified as Enter Key
# Else uncomment Keys.Enter in the last step if you dont want to use '\n'
# Keep a nice gap between successive messages
# Use Keys.SHIFT + Keys.ENTER to give a new line effect in your Message
# msgToSend = [
#                 [12, 32, 0, "Hello! This is test Msg. Please Ignore." + Keys.SHIFT + Keys.ENTER + "http://bit.ly/mogjm05"]
#             ]

# Count variable to identify the number of messages to be sent
# count = 0
# while count<len(msgToSend):

#     # Identify time
#     curTime = datetime.datetime.now()
#     curHour = curTime.time().hour
#     curMin = curTime.time().minute
#     curSec = curTime.time().second

#     # if time matches then move further
#     if msgToSend[count][0]==curHour and msgToSend[count][1]==curMin and msgToSend[count][2]==curSec:
#         # utility variables to tract count of success and fails
#         success = 0
#         sNo = 1
#         failList = []

#         # Iterate over selected contacts
#         for target in targets:
#             print(sNo, ". Target is: " + target)
#             sNo+=1
#             try:
#                 # Select the target
#                 x_arg = '//span[contains(@title,' + target + ')]'
#                 try:
#                     wait5.until(EC.presence_of_element_located((
#                         By.XPATH, x_arg
#                     )))
#                 except:
#                     # If contact not found, then search for it
#                     searBoxPath = '//*[@id="input-chatlist-search"]'
#                     wait5.until(EC.presence_of_element_located((
#                         By.ID, "input-chatlist-search"
#                     )))
#                     inputSearchBox = driver.find_element_by_id("input-chatlist-search")
#                     time.sleep(0.5)
#                     # click the search button
#                     driver.find_element_by_xpath('/html/body/div/div/div/div[2]/div/div[2]/div/button').click()
#                     time.sleep(1)
#                     inputSearchBox.clear()
#                     inputSearchBox.send_keys(target[1:len(target) - 1])
#                     print('Target Searched')
#                     # Increase the time if searching a contact is taking a long time
#                     time.sleep(4)

#                 # Select the target
#                 driver.find_element_by_xpath(x_arg).click()
#                 print("Target Successfully Selected")
#                 time.sleep(2)

#                 # Select the Input Box
#                 inp_xpath = "//div[(@contenteditable='true') and (@data-tab='1')]"
#                 input_box = wait.until(EC.presence_of_element_located((
#                     By.XPATH, inp_xpath)))
#                 time.sleep(1)

#                 # Send message
#                 # taeget is your target Name and msgToSend is you message
#                 input_box.send_keys("Hello, " + target + "."+ Keys.SHIFT + Keys.ENTER + msgToSend[count][3] + Keys.SPACE) # + Keys.ENTER (Uncomment it if your msg doesnt contain '\n')
#                 # Link Preview Time, Reduce this time, if internet connection is Good
#                 time.sleep(10)
#                 input_box.send_keys(Keys.ENTER)
#                 print("Successfully Send Message to : "+ target + '\n')
#                 success+=1
#                 time.sleep(0.5)

#             except:
#                 # If target Not found Add it to the failed List
#                 print("Cannot find Target: " + target)
#                 failList.append(target)
#                 pass

#         print("\nSuccessfully Sent to: ", success)
#         print("Failed to Sent to: ", len(failList))
#         print(failList)
#         print('\n\n')
#         count+=1
    driver.quit()