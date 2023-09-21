import time
import wget
import shutil
from xlsx import xlsx,text
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

chrome_options = webdriver.ChromeOptions()
#C:\Users\Glenda\AppData\Local\Google\Chrome\User Data\Profile 2
# chrome_options.add_argument("--user-data-dir=C:\\Users\\_kul2o_\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 2")
chrome_options.add_argument("--user-data-dir=C:\\Users\\Glenda\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 2")
pro_type = {
    "apartment" : "16",
    "bar" : "17",
    "cafe" : "21",
    "condominium" : "101",
    "farm" : "35",
    "hotel/resort" : "102",
    "house" : "42",
    "land" : "103",
    "luxury homes" : "47",
    "office" : "56",
    "single family" : "71",
    "store" : "76",
    "villa" : "85"
}

pro_status = {
    "for rent" : "37",
    "for sale" : "38"
}

def checkNan(data):
    if str(data) == 'nan':
        return ""
    else:
        return data
    
bot = xlsx('Listing-updated.xlsx')
bot.read()

driver = webdriver.Chrome(options=chrome_options)

url = "https://charish.co.th/new-property-2/"
driver.get(url)
#time.sleep(10000)

for i  in range(len(bot.data)):

    driver.get(url)

    data = bot.get_value_row(i)

    code = checkNan(data["code"])
    condotel = checkNan(data["condotel"])
    remark = checkNan(data["remark"])
    postal_code = checkNan(data["postal_code"])
    type_text = data["type_text"].lower()
    status = data["status"].lower()
    price = checkNan(data["price"])
    rent = checkNan(data["rent"])
    size = checkNan(data["size"])
    bed = checkNan(data["bed"])
    bath = checkNan(data["bath"])
    floor = checkNan(data["floor"])
    building = checkNan(data["building"])
    fq = checkNan(data["fq"]) != "" and "✅" or ""
    tq = checkNan(data["tq"]) != "" and "✅" or ""
    contact_name = data["contact_name"]
    contact_email = data["contact_email"]
    contact_tel = data["contact_tel"]
    contact_info = data["contact_info"]

    print(f"--- data [{i}] | row {i+1} ---")

    time.sleep(5)
    try:
        driver.find_element(By.XPATH,"/html/body/div[1]/div/div[2]/main/div/section/form/fieldset[1]/div[2]/button[1]").click()
        print('click Show All')
    except:
        time.sleep(5)
        driver.find_element(By.XPATH,"/html/body/div[1]/div/div[2]/main/div/section/form/fieldset[1]/div[2]/button[1]").click()
        print('click Show All')

    time.sleep(5)

    try:
        driver.find_element(By.ID,"property_title").send_keys(condotel)
        print('send title')
    except:
        time.sleep(5)
        driver.find_element(By.ID,"property_title").send_keys(condotel)
        print('send title')

    time.sleep(5)

    try:
        driver.find_element(By.ID,"property_des-html").click()
        print('click massage')
    except:
        time.sleep(5)
        driver.find_element(By.ID,"property_des-html").click()
        print('click massage')
        
    time.sleep(5)

    try:
        driver.find_element(By.ID,"property_des").send_keys(remark)
        print('send description')
    except:
        time.sleep(5)
        driver.find_element(By.ID,"property_des").send_keys(remark)
        print('send description')
        
    time.sleep(5)

    try:
        driver.find_element(By.ID,"address1").send_keys(condotel)
        print('send address')
    except:
        time.sleep(5)
        driver.find_element(By.ID,"address1").send_keys(condotel)
        print('send address')
        
    time.sleep(5)

    try:
        Select(driver.find_element(By.ID,"administrative_area_level_1")).select_by_value("chonburi")
        print('select chonburi')
    except:
        time.sleep(5)
        Select(driver.find_element(By.ID,"administrative_area_level_1")).select_by_value("chonburi")
        print('select chonburi')
        
    time.sleep(5)

    try:
        Select(driver.find_element(By.ID,"city")).select_by_value("pattaya")
        print('select pattaya')
    except:
        time.sleep(5)
        Select(driver.find_element(By.ID,"city")).select_by_value("pattaya")
        print('select pattaya')
        
    time.sleep(5)

    try:
        driver.find_element(By.ID,"zip").send_keys(postal_code)
        print('send postal code')
    except:
        time.sleep(5)
        driver.find_element(By.ID,"zip").send_keys(postal_code)
        print('send postal code')
        
    time.sleep(5)

    try:
        driver.find_element(By.NAME,"property_search_address").send_keys(condotel)
        driver.find_element(By.NAME,"property_search_address").send_keys(Keys.ENTER)
        print('send map location')
    except:
        time.sleep(5)
        driver.find_element(By.NAME,"property_search_address").send_keys(condotel)
        driver.find_element(By.NAME,"property_search_address").send_keys(Keys.ENTER)
        print('send map location')
        
    time.sleep(5)

    try:
        if type_text != '':
            Select(driver.find_element(By.ID,"property_type")).select_by_value(pro_type[type_text])
            print('select property type')
    except:
        time.sleep(5)
        if type_text != '':
            Select(driver.find_element(By.ID,"property_type")).select_by_value(pro_type[type_text])
            print('select property type')
        
    time.sleep(5)

    try:
        if status != '':
            Select(driver.find_element(By.ID,"property_status")).select_by_value(pro_status[status])
            print('select property status')
    except:
        time.sleep(5)
        if status != '':
            Select(driver.find_element(By.ID,"property_status")).select_by_value(pro_status[status])
            print('select property status')
        
    time.sleep(5)

    try:
        driver.find_element(By.ID,"property_price_short").send_keys(price)
        print('send price')
    except:
        time.sleep(5)
        driver.find_element(By.ID,"property_price_short").send_keys(price)
        print('send price')
        
    time.sleep(5)

    try:
        driver.find_element(By.ID,"property_price_postfix").send_keys(rent)
        print('send after price')
        time.sleep(5)
        if rent != "":
            try:
                Select(driver.find_element(By.ID,"property_status")).select_by_value(pro_status["For Rent"])
                print('select property status')
            except:
                time.sleep(5)
                Select(driver.find_element(By.ID,"property_status")).select_by_value(pro_status["For Rent"])
                print('select property status')
    except:
        time.sleep(5)
        driver.find_element(By.ID,"property_price_postfix").send_keys(rent)
        print('send after price')
        time.sleep(5)
        if rent != "":
            try:
                Select(driver.find_element(By.ID,"property_status")).select_by_value(pro_status["For Rent"])
                print('select property status')
            except:
                time.sleep(5)
                Select(driver.find_element(By.ID,"property_status")).select_by_value(pro_status["For Rent"])
                print('select property status')
        
    time.sleep(5)

    try:
        driver.find_element(By.ID,"property_size").send_keys(size)
        print('send size')
    except:
        time.sleep(5)
        driver.find_element(By.ID,"property_size").send_keys(size)
        print('send size')
        
    time.sleep(5)

    try:
        driver.find_element(By.ID,"property_bedrooms").send_keys(bed)
        print('send bedroom')
    except:
        time.sleep(5)
        driver.find_element(By.ID,"property_bedrooms").send_keys(bed)
        print('send bedroom')
        
    time.sleep(5)

    try:
        driver.find_element(By.ID,"property_bathrooms").send_keys(bath)
        print('send bathroom')
    except:
        time.sleep(5)
        driver.find_element(By.ID,"property_bathrooms").send_keys(bath)
        print('send bathroom')
        
    time.sleep(5)

    try:
        driver.find_element(By.ID,"property_identity").send_keys(code)
        print('send property id')
    except:
        time.sleep(5)
        driver.find_element(By.ID,"property_identity").send_keys(code)
        print('send property id')
        
    time.sleep(5)

    try:
        driver.find_element(By.XPATH,"/html/body/div[1]/div/div[2]/main/div/section/form/fieldset[6]/div[1]/div/div[3]/table/tfoot/tr/td[2]/button").click()
        driver.find_element(By.ID,"additional_feature_title_0").send_keys("Floor")
        driver.find_element(By.ID,"additional_feature_value_0").send_keys(floor)
        print('add floor')
    except:
        time.sleep(5)
        driver.find_element(By.XPATH,"/html/body/div[1]/div/div[2]/main/div/section/form/fieldset[6]/div[1]/div/div[3]/table/tfoot/tr/td[2]/button").click()
        driver.find_element(By.ID,"additional_feature_title_0").send_keys("Floor")
        driver.find_element(By.ID,"additional_feature_value_0").send_keys(floor)
        print('add floor')
        
    time.sleep(5)

    try:
        driver.find_element(By.XPATH,"/html/body/div[1]/div/div[2]/main/div/section/form/fieldset[6]/div[1]/div/div[3]/table/tfoot/tr/td[2]/button").click()
        driver.find_element(By.ID,"additional_feature_title_1").send_keys("Building")
        driver.find_element(By.ID,"additional_feature_value_1").send_keys(building)
        print('add building')
    except:
        time.sleep(5)
        driver.find_element(By.XPATH,"/html/body/div[1]/div/div[2]/main/div/section/form/fieldset[6]/div[1]/div/div[3]/table/tfoot/tr/td[2]/button").click()
        driver.find_element(By.ID,"additional_feature_title_1").send_keys("Building")
        driver.find_element(By.ID,"additional_feature_value_1").send_keys(building)
        print('add building')
        
    time.sleep(5)

    try:
        driver.find_element(By.XPATH,"/html/body/div[1]/div/div[2]/main/div/section/form/fieldset[6]/div[1]/div/div[3]/table/tfoot/tr/td[2]/button").click()
        driver.find_element(By.ID,"additional_feature_title_2").send_keys("FQ")
        driver.find_element(By.ID,"additional_feature_value_2").send_keys(fq)
        print('add FQ')
    except:
        time.sleep(5)
        driver.find_element(By.XPATH,"/html/body/div[1]/div/div[2]/main/div/section/form/fieldset[6]/div[1]/div/div[3]/table/tfoot/tr/td[2]/button").click()
        driver.find_element(By.ID,"additional_feature_title_2").send_keys("FQ")
        driver.find_element(By.ID,"additional_feature_value_2").send_keys(fq)
        print('add FQ')
        
    time.sleep(5)

    try:
        driver.find_element(By.XPATH,"/html/body/div[1]/div/div[2]/main/div/section/form/fieldset[6]/div[1]/div/div[3]/table/tfoot/tr/td[2]/button").click()
        driver.find_element(By.ID,"additional_feature_title_3").send_keys("TQ")
        driver.find_element(By.ID,"additional_feature_value_3").send_keys(tq)
        print('add TQ')
    except:
        time.sleep(5)
        driver.find_element(By.XPATH,"/html/body/div[1]/div/div[2]/main/div/section/form/fieldset[6]/div[1]/div/div[3]/table/tfoot/tr/td[2]/button").click()
        driver.find_element(By.ID,"additional_feature_title_3").send_keys("TQ")
        driver.find_element(By.ID,"additional_feature_value_3").send_keys(tq)
        print('add TQ')
        
    time.sleep(5)

    try:
        driver.find_element(By.XPATH,"/html/body/div[1]/div/div[2]/main/div/section/form/fieldset[9]/div[1]/div/div[2]/div[2]").click()
        print('click other contact')
    except:
        time.sleep(5)
        driver.find_element(By.XPATH,"/html/body/div[1]/div/div[2]/main/div/section/form/fieldset[9]/div[1]/div/div[2]/div[2]").click()
        print('click other contact')
        
    time.sleep(5)

    try:
        driver.find_element(By.ID,"property_other_contact_name").send_keys(contact_name)
        print('send contact name')
    except:
        time.sleep(5)
        driver.find_element(By.ID,"property_other_contact_name").send_keys(contact_name)
        print('send contact name')
        
    time.sleep(5)

    try:
        driver.find_element(By.ID,"property_other_contact_mail").send_keys(contact_email)
        print('send contact email')
    except:
        time.sleep(5)
        driver.find_element(By.ID,"property_other_contact_mail").send_keys(contact_email)
        print('send contact email')
        
    time.sleep(5)

    try:
        driver.find_element(By.ID,"property_other_contact_phone").send_keys(contact_tel)
        print('send contact phone')
    except:
        time.sleep(5)
        driver.find_element(By.ID,"property_other_contact_phone").send_keys(contact_tel)
        print('send contact phone')
        
    time.sleep(5)

    try:
        driver.find_element(By.ID,"property_other_contact_description").send_keys(contact_info)
        print('send contact info')
    except:
        time.sleep(5)
        driver.find_element(By.ID,"property_other_contact_description").send_keys(contact_info)
        print('send contact info')
        
    time.sleep(5)

    try:
        driver.find_element(By.NAME,"submit_property").click()
        print('click submit')
    except:
        time.sleep(5)
        driver.find_element(By.NAME,"submit_property").click()
        print('click submit')
        
    time.sleep(5)

    print("------------------------")


