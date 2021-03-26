from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import pickle
import time
import os
import json
import pyperclip,pyautogui


class ShopeeAutomation:
    def __init__(self):
        self.url = ""
        self.browser = ""
        self.orderDetails = ""
        self.cookies = ""
        self.windows_size = 0

    def login(self):
        WebDriverWait(self.browser, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="app"]/div[2]/div/div/div/div[3]/div/div/div/div[2]/button')))

        username = "gardenlabphilippines:staff1"
        passw = "IOU123shopee"

        self.browser.find_element_by_xpath(
            '//*[@id="app"]/div[2]/div/div/div/div[3]/div/div/div/div[2]/button').click()

        WebDriverWait(self.browser, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div/main/div/div[1]/div/div/div/div/div/div/div[2]/div[2]/div/div/input')))  #

        self.browser.find_element_by_xpath(
            '/html/body/div/main/div/div[1]/div/div/div/div/div/div/div[2]/div[1]/div/div/div/input').send_keys(
            username)
        self.browser.find_element_by_xpath(
            '/html/body/div/main/div/div[1]/div/div/div/div/div/div/div[2]/div[2]/div/div/input').send_keys(passw)
        self.browser.find_element_by_xpath('/html/body/div/main/div/div[1]/div/div/div/div/div/div/button[2]').click()

        self.browser.find_element_by_xpath(
            '//*[@id="shop-login"]/div[3]/label/span[1]').click()

        WebDriverWait(self.browser, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div/main/div/div[1]/div/div/div/div/div/div/div[3]/div/div/input')))

        otp = input("Enter OTP:")
        self.browser.find_element_by_xpath(
            '/html/body/div/main/div/div[1]/div/div/div/div/div/div/div[3]/div/div/input').send_keys(otp)
        self.browser.find_element_by_xpath('/html/body/div/main/div/div[1]/div/div/div/div/div/div/button').click()

        time.sleep(10)

        self.cookies = self.browser.get_cookies()
        pickle.dump(self.cookies, open('cookies.pkl', 'wb'))

    def setup(self):
        """Take all the inputs needed."""
        #self.read_inputs()
        self.url = "https://seller.shopee.ph/account/signin?next=%2F"
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
        self.browser = webdriver.Chrome(r'chrome_webdriver/chromedriver.exe', options=chrome_options)
        self.browser.maximize_window()
        self.windows_size = self.browser.get_window_size()
        self.browser.get(self.url)
        time.sleep(5)
        if os.path.isfile('cookies.pkl'):
            self.cookies = pickle.load(open("cookies.pkl", "rb"))
            for cookie in self.cookies:
                self.browser.add_cookie(cookie)
            self.browser.refresh()

            if self.browser.current_url == self.url:
                self.login()
            elif 1==2:
                pass

        else:
            self.login()

    def run(self):
        """Main code to run the program"""
        # Mass shipping
        WebDriverWait(self.browser, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="app"]/div[2]/div[1]/div/div[2]/ul/li[1]/ul/li[2]/a')))

        self.browser.find_element_by_xpath('//*[@id="app"]/div[2]/div[1]/div/div[2]/ul/li[1]/ul/li[2]/a').click()

        # Orders to ship
        WebDriverWait(self.browser, 10).until(
            EC.presence_of_element_located(
                (By.XPATH,
                 '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[1]/div/div[1]/div/div[1]/div[1]/div/i')))

        self.browser.find_element_by_xpath(
            '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[1]/div/div[1]/div/div[1]/div[1]/div/i').click()

        time.sleep(5)

        self.standard_delivery()

        self.browser.refresh()
        time.sleep(3)

        self.jandt_express_delivery()

        time.sleep(10)

    def get_orders_generate_pdf(self, jt=0):
        time.sleep(3)
        flag = 0
        orders = self.browser.find_elements_by_class_name("mass-ship-list-item")
        orders = orders[1:3]

        number_of_orders = len(orders)
        if number_of_orders == 0:
            return 0

        order_ids = []
        already_have = [dir.split('.')[0] for dir in os.listdir(os.getcwd() + '\\waybill_pdf')]
        self.orderDetails = []
        order_config = {"OrderId": None, "TrackingId": None, "Products": None}
        product_details = {"Name": None, "Quantity": None, "Variation": None, "UnitPrice": None, "SubTotal": None}

        while number_of_orders > 0:
            time.sleep(3)

            orders = self.browser.find_elements_by_class_name("mass-ship-list-item")
            orders = orders[1:]
            if flag:
                order = orders[1]
            else:
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
                self.browser.switch_to.window(self.browser.window_handles[1])
                WebDriverWait(self.browser, 10).until(
                    EC.presence_of_element_located(
                        (By.CLASS_NAME, 'product-list-item')))

                products = self.browser.find_elements_by_class_name('product-list-item')[1:]
                for product in products:
                    product_details['Name'] = product.find_element_by_class_name("product-name").text
                    product_details['Variation'] = product.find_element_by_class_name('product-meta').text
                    product_details['Quantity'] = product.find_element_by_class_name('qty').text
                    product_details['UnitPrice'] = product.find_element_by_class_name('price').text
                    product_details['SubTotal'] = product.find_element_by_class_name('subtotal').text

                    order_config['Products'].append(product_details.copy())
                    for key in product_details.keys():
                        product_details[key] = None

                self.orderDetails.append(order_config.copy())

                self.browser.close()
                self.browser.switch_to.window(self.browser.window_handles[0])

                # Print waybill page
                order.find_element_by_class_name("shopee-checkbox__indicator").click()
                time.sleep(2)
                if not order.find_element_by_class_name("shopee-checkbox__indicator").is_selected():
                    print("Details saved but cannot generate pdf(Cannot select checkbox) for order id {}.".format(order_id))
                    self.browser.refresh()
                    flag = 1
                    continue
                time.sleep(3)

                # Mass pickup button
                if jt:
                    self.browser.find_element_by_xpath(
                    '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/button').click()
                else:
                    self.browser.find_element_by_xpath(
                    '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[1]/div/div/div[3]/div/div/div/div/button').click()


                # Confirm Button
                WebDriverWait(self.browser, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH,
                         '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[1]/div/div/div[3]/div/div/div/div/div[4]/button[2]')))

                self.browser.find_element_by_xpath(
                    '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[1]/div/div/div[3]/div/div/div/div/div[4]/button[2]').click()

                # Generate button
                WebDriverWait(self.browser, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH,
                         '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[1]/div[2]/div/div/div[1]/div/div[2]/div[3]/div[2]/div/button')))

                # tracking_num = ""
                # while len(tracking_num.strip()) == 0:
                #     tracking_num = self.browser.find_element_by_xpath(
                #         '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[1]/div[2]/div/div/div[1]/div/div[2]/div[2]/div/div[2]/div[3]').text
                #
                # order_config["TrackingId"] = self.browser.find_element_by_xpath(
                #         '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[1]/div[2]/div/div/div[1]/div/div[2]/div[2]/div/div[2]/div[3]').text
                time.sleep(5)
                generate_button_ele = self.browser.find_element_by_xpath(
                    '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[1]/div[2]/div/div/div[1]/div/div[2]/div[3]/div[2]/div/button')

                action = ActionChains(self.browser)
                action.move_to_element(generate_button_ele).perform()

                time.sleep(2)

                generate_waybill_button = self.browser.find_element_by_xpath('/html/body/div[5]/ul/li[2]/div[2]/div')

                action.move_to_element(generate_waybill_button).perform()
                generate_waybill_button.click()

                time.sleep(3)

                self.save_pdf(order_id)

                self.browser.switch_to.window(self.browser.window_handles[0])
                time.sleep(3)
                self.browser.find_element_by_xpath(
                    '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[1]/div[2]/div/div/div[1]/div/div[2]/div[1]/div[2]/i').click()
                time.sleep(5)
                self.browser.refresh()
                WebDriverWait(self.browser, 10).until(
                    EC.presence_of_element_located(
                        (By.CLASS_NAME, 'mass-ship-list-item')))

            else:
                print("already have order id: ", order_id)

    def save_pdf(self, order_id):
        # Saving PDF
        self.browser.switch_to.window(self.browser.window_handles[1])
        time.sleep(10)
        pyautogui.rightClick(x=self.windows_size['width'] // 2, y=self.windows_size['height'] // 2)
        pyautogui.typewrite(['a'])
        path = os.getcwd() + '\\waybill_pdf\\{0}.pdf'.format(order_id)
        pyperclip.copy(path)
        time.sleep(5)
        pyautogui.hotkey('ctrlleft', 'V')
        pyautogui.press('enter')

        self.browser.close()

    def standard_delivery(self):
        # Standard delivery
        WebDriverWait(self.browser, 10).until(
            EC.presence_of_element_located(
                (By.XPATH,
                 '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[1]/div/div[2]/div[1]/div/div[1]/label[1]')))

        self.browser.find_element_by_xpath(
            '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[1]/div/div[2]/div[1]/div/div[1]/label[1]').click()

        WebDriverWait(self.browser, 10).until(
            EC.presence_of_element_located(
                (By.CLASS_NAME, 'mass-ship-list-item')))

        self.get_orders_generate_pdf()
        if len(self.orderDetails) == 0:
            print("No orders in Standard delivery.")
        #
        else:
            with open('standard_order.json', 'w') as outfile:
                json.dump(self.orderDetails, outfile)
        # print(self.orderDetails)

    def jandt_express_delivery(self):
        # J&T Express delivery
        WebDriverWait(self.browser, 10).until(
            EC.presence_of_element_located(
                (By.XPATH,
                 '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[1]/div/div[2]/div[1]/div/div[1]/label[2]')))

        self.browser.find_element_by_xpath(
            '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[1]/div/div[2]/div[1]/div/div[1]/label[2]').click()

        WebDriverWait(self.browser, 10).until(
            EC.presence_of_element_located(
                (By.CLASS_NAME, 'mass-ship-list-item')))

        self.get_orders_generate_pdf(jt=1)
        if self.orderDetails == 0:
            print("No orders in J&T Express delivery.")
        else:
            with open('j&tExpress_order.json', 'w') as outfile:
                json.dump(self.orderDetails, outfile)
        # print(self.orderDetails)

    def teardown(self):
        """Close browser"""
        self.browser.close()


if __name__ == "__main__":
    shopeeAutomate = ShopeeAutomation()
    shopeeAutomate.setup()
    shopeeAutomate.run()
    shopeeAutomate.teardown()
