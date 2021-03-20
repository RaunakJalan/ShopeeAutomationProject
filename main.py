from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import pickle
import time
import os
import json
import pyperclip,pyautogui

url = "https://seller.shopee.ph/account/signin?next=%2F"
chrome_options = Options()
settings = {
       "recentDestinations": [{
            "id": "Save as PDF",
            "origin": "local",
            "account": "",
        }],
        "selectedDestinationId": "Save as PDF",
        "version": 2
    }
prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings)}
chrome_options.add_experimental_option('prefs', prefs)
chrome_options.add_argument('--kiosk-printing')
# self.chrome_options.add_argument("--headless")
# self.chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
browser = webdriver.Chrome(r'chrome_webdriver/chromedriver.exe', options=chrome_options)
browser.maximize_window()
windows_size = browser.get_window_size()
browser.get(url)
time.sleep(5)
if os.path.isfile('cookies.pkl'):
    cookies = pickle.load(open("cookies.pkl", "rb"))
    for cookie in cookies:
        browser.add_cookie(cookie)
    browser.refresh()

else:

    WebDriverWait(browser, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@id="app"]/div[2]/div/div/div/div[3]/div/div/div/div[2]/button')))

    username = "gardenlabphilippines:staff1"
    passw = "IOU123shopee"

    browser.find_element_by_xpath('//*[@id="app"]/div[2]/div/div/div/div[3]/div/div/div/div[2]/button').click()

    WebDriverWait(browser, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div/main/div/div[1]/div/div/div/div/div/div/div[2]/div[2]/div/div/input')))#

    browser.find_element_by_xpath(
        '/html/body/div/main/div/div[1]/div/div/div/div/div/div/div[2]/div[1]/div/div/div/input').send_keys(username)
    browser.find_element_by_xpath(
        '/html/body/div/main/div/div[1]/div/div/div/div/div/div/div[2]/div[2]/div/div/input').send_keys(passw)
    browser.find_element_by_xpath('/html/body/div/main/div/div[1]/div/div/div/div/div/div/button[2]').click()

    WebDriverWait(browser, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div/main/div/div[1]/div/div/div/div/div/div/div[3]/div/div/input')))

    otp = input("Enter OTP:")
    browser.find_element_by_xpath(
        '/html/body/div/main/div/div[1]/div/div/div/div/div/div/div[3]/div/div/input').send_keys(otp)
    browser.find_element_by_xpath('/html/body/div/main/div/div[1]/div/div/div/div/div/div/button').click()

    time.sleep(10)

    cookies = browser.get_cookies()
    pickle.dump(cookies, open('cookies.pkl', 'wb'))

# Mass shipping
WebDriverWait(browser, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@id="app"]/div[2]/div[1]/div/div[2]/ul/li[1]/ul/li[2]/a')))

browser.find_element_by_xpath('//*[@id="app"]/div[2]/div[1]/div/div[2]/ul/li[1]/ul/li[2]/a').click()


def get_orders_generate_pdf(jt=0):
    time.sleep(3)
    orders = browser.find_elements_by_class_name("mass-ship-list-item")
    orders = orders[1:3]

    number_of_orders = len(orders)

    order_ids = []
    already_have = [dir.split('.')[0] for dir in os.listdir(os.getcwd() + '\\waybill_pdf')]
    order_details = []
    order_config = {"OrderId": None, "TrackingId": None, "Products": None}
    product_details = {"Name": None, "Quantity": None, "Variation": None, "UnitPrice": None, "SubTotal": None}

    while number_of_orders > 0:
        time.sleep(3)
        
        orders = browser.find_elements_by_class_name("mass-ship-list-item")
        orders = orders[1:]
        order = orders[0]
        number_of_orders -= 1
        order_id_ele = order.find_element_by_class_name('orderid')
        order_id = order_id_ele.text
        if order_id not in already_have:
            order_ids.append(order_id)
            order_config["OrderId"] = order_id
            order_config["Products"] = list()

            order_id_ele.find_element_by_tag_name('a').click()
            time.sleep(3)
            # Order details page
            browser.switch_to.window(browser.window_handles[1])
            WebDriverWait(browser, 10).until(
                EC.presence_of_element_located(
                    (By.CLASS_NAME, 'product-list-item')))

            products = browser.find_elements_by_class_name('product-list-item')[1:]
            for product in products:
                product_details['Name'] = product.find_element_by_class_name("product-name").text
                product_details['Variation'] = product.find_element_by_class_name('product-meta').text
                product_details['Quantity'] = product.find_element_by_class_name('qty').text
                product_details['UnitPrice'] = product.find_element_by_class_name('price').text
                product_details['SubTotal'] = product.find_element_by_class_name('subtotal').text

                order_config['Products'].append(product_details.copy())
                for key in product_details.keys():
                    product_details[key] = None

            order_details.append(order_config.copy())

            browser.close()
            browser.switch_to.window(browser.window_handles[0])

            # Print waybill page
            order.find_element_by_class_name("shopee-checkbox__indicator").click()
            time.sleep(3)

            # Mass pickup button
            if jt:
                browser.find_element_by_xpath(
                '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/button').click()
            else:
                browser.find_element_by_xpath(
                '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[1]/div/div/div[3]/div/div/div/div/button').click()



            # Confirm Button
            WebDriverWait(browser, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH,
                     '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[1]/div/div/div[3]/div/div/div/div/div[4]/button[2]')))

            browser.find_element_by_xpath(
                '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[1]/div/div/div[3]/div/div/div/div/div[4]/button[2]').click()

            # Generate button
            WebDriverWait(browser, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH,
                     '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[1]/div[2]/div/div/div[1]/div/div[2]/div[3]/div[2]/div/button')))

            # tracking_num = ""
            # while len(tracking_num.strip()) == 0:
            #     tracking_num = browser.find_element_by_xpath(
            #         '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[1]/div[2]/div/div/div[1]/div/div[2]/div[2]/div/div[2]/div[3]').text
            #
            # order_config["TrackingId"] = browser.find_element_by_xpath(
            #         '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[1]/div[2]/div/div/div[1]/div/div[2]/div[2]/div/div[2]/div[3]').text
            time.sleep(5)
            generate_button_ele = browser.find_element_by_xpath(
                '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[1]/div[2]/div/div/div[1]/div/div[2]/div[3]/div[2]/div/button')

            action = ActionChains(browser)
            action.move_to_element(generate_button_ele).perform()

            time.sleep(2)

            generate_waybill_button = browser.find_element_by_xpath('/html/body/div[5]/ul/li[2]/div[2]/div')

            action.move_to_element(generate_waybill_button).perform()
            generate_waybill_button.click()

            time.sleep(3)
            browser.switch_to.window(browser.window_handles[1])
            time.sleep(10)
            pyautogui.rightClick(x=windows_size['width'] // 2, y=windows_size['height'] // 2)
            pyautogui.typewrite(['a'])
            path = os.getcwd() + '\\waybill_pdf\\{0}.pdf'.format(order_id)
            pyperclip.copy(path)
            time.sleep(5)
            pyautogui.hotkey('ctrlleft', 'V')
            pyautogui.press('enter')

            browser.close()
            browser.switch_to.window(browser.window_handles[0])
            time.sleep(3)
            browser.find_element_by_xpath(
                '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[1]/div[2]/div/div/div[1]/div/div[2]/div[1]/div[2]/i').click()
            time.sleep(5)
            browser.refresh()
            WebDriverWait(browser, 10).until(
                EC.presence_of_element_located(
                    (By.CLASS_NAME, 'mass-ship-list-item')))

        else:
            print("already have order id: ", order_id)

    return order_details


# Orders to ship
WebDriverWait(browser, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[1]/div/div[1]/div/div[1]/div[1]/div/i')))

browser.find_element_by_xpath('//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[1]/div/div[1]/div/div[1]/div[1]/div/i').click()

time.sleep(5)

# Standard delivery
WebDriverWait(browser, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[1]/div/div[2]/div[1]/div/div[1]/label[1]')))

browser.find_element_by_xpath('//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[1]/div/div[2]/div[1]/div/div[1]/label[1]').click()

WebDriverWait(browser, 10).until(
        EC.presence_of_element_located(
            (By.CLASS_NAME, 'mass-ship-list-item')))

order_ids = get_orders_generate_pdf()

with open('standard_order.json', 'w') as outfile:
    json.dump(order_ids, outfile)
print(order_ids)

browser.refresh()
time.sleep(3)

# J&T Express delivery
WebDriverWait(browser, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[1]/div/div[2]/div[1]/div/div[1]/label[2]')))

browser.find_element_by_xpath('//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[1]/div/div[2]/div[1]/div/div[1]/label[2]').click()

WebDriverWait(browser, 10).until(
    EC.presence_of_element_located(
        (By.CLASS_NAME, 'mass-ship-list-item')))

order_ids = get_orders_generate_pdf(jt=1)

with open('j&tExpress_order.json', 'w') as outfile:
    json.dump(order_ids, outfile)
print(order_ids)


time.sleep(10)

browser.close()
