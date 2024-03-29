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
import pyperclip, pyautogui
import glob
import traceback
from datetime import datetime


class ShopeeAutomation:
    def __init__(self):
        self.browser = ""
        self.orderDetails = []
        self.cookies = ""
        self.windows_size = 0

        paths = ""
        with open("config\\config.json", "r") as file:
            paths = json.loads(file.read())[0]
        self.username = paths.get("username")
        self.passw = paths.get("password")
        self.homeDirectory = paths.get("homeDirectory")
        self.jsonFolder = paths.get("jsonFolder")
        self.pdfFolder = paths.get("pdfFolder")
        self.url = paths.get("url")

    def login(self):
        WebDriverWait(self.browser, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="app"]/div[2]/div/div/div/div[3]/div/div/div/div[2]/button')))

        self.browser.find_element_by_xpath(
            '//*[@id="app"]/div[2]/div/div/div/div[3]/div/div/div/div[2]/button').click()

        WebDriverWait(self.browser, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div/main/div/div[1]/div/div/div/div/div/div/div[2]/div[2]/div/div/input')))  #

        self.browser.find_element_by_xpath(
            '/html/body/div/main/div/div[1]/div/div/div/div/div/div/div[2]/div[1]/div/div/div/input').send_keys(
            self.username)
        self.browser.find_element_by_xpath(
            '/html/body/div/main/div/div[1]/div/div/div/div/div/div/div[2]/div[2]/div/div/input').send_keys(self.passw)
        self.browser.find_element_by_xpath('/html/body/div/main/div/div[1]/div/div/div/div/div/div/button[2]').click()

        # self.browser.find_element_by_xpath(
        #    '//*[@id="shop-login"]/div[3]/label/span[1]').click()

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
            time.sleep(5)
            if self.browser.current_url == self.url:
                self.login()
                print(self.browser.current_url)
                print(self.url)

        else:
            self.login()

    def run(self):
        """Main code to run the program"""
        # Mass shipping
        WebDriverWait(self.browser, 10).until(
            EC.element_to_be_clickable(
                (By.LINK_TEXT, 'Mass Ship')))

        self.browser.find_element_by_link_text('Mass Ship').click()

        # Orders to ship
        WebDriverWait(self.browser, 10).until(
            EC.presence_of_element_located(
                (By.XPATH,
                 '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[1]/div/div[1]/div/div[1]/div[1]/div/i')))

        self.browser.find_element_by_xpath(
            '//*[@id="app"]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div[1]/div/div[1]/div/div[1]/div[1]/div/i').click()

        time.sleep(5)

        WebDriverWait(self.browser, 10).until(
            EC.element_to_be_clickable(
                (By.CLASS_NAME,
                 'shopee-radio-button__label')))

        delivery_types = self.browser.find_elements_by_class_name('shopee-radio-button__label')
        num_types = 0
        while num_types < len(delivery_types):
            print(delivery_types[num_types].text.lower())
            if "other" in delivery_types[num_types].text.lower():
                num_types += 1
                continue
            delivery_types[num_types].click()
            self.delivery(delivery_types[num_types].text)
            time.sleep(5)
            num_types += 1
            delivery_types = self.browser.find_elements_by_class_name('shopee-radio-button__label')

        time.sleep(10)

    def get_orders_generate_pdf(self):

        flag = 1
        time.sleep(5)
        orders = self.browser.find_elements_by_class_name("mass-ship-list-item")
        number_of_orders = len(orders)
        if number_of_orders == 1:
            return 0

        orders = orders[1:]
        # number_of_orders = len(orders)
        number_of_orders = 2 if len(orders) > 2 else 1
        order_ids = []
        order_config = {"OrderId": None, "Products": None}
        product_details = {"Name": None, "Quantity": None, "Variation": None, "UnitPrice": None, "SubTotal": None}

        while number_of_orders > 0:
            time.sleep(5)
            orders = self.browser.find_elements_by_class_name("mass-ship-list-item")
            order = orders[flag]
            order.find_element_by_class_name("shopee-checkbox__indicator").click()

            order_id_ele = order.find_element_by_class_name('orderid')
            order_id = order_id_ele.text
            time.sleep(3)
            if not order.find_element_by_class_name("shopee-checkbox__input").is_selected() and flag < number_of_orders:
                flag += 1
                print("Cannot generate PDF for order id {}. Skipping it".format(order_id))

            number_of_orders = number_of_orders - 1
            # number_of_orders = len(orders) - 1 - flag
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

            time.sleep(3)

            # Mass pickup button
            self.browser.find_element_by_xpath('//button[normalize-space()="Mass Arrange Pickup"]').click()

            # Confirm Button
            WebDriverWait(self.browser, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH,
                     '//button[normalize-space()="Confirm"]')))

            element = self.browser.find_element_by_xpath('//button[normalize-space()="Confirm"]')
            action = ActionChains(self.browser)
            action.move_to_element(element).perform()
            self.browser.execute_script("arguments[0].click();", element)

            # Generate button
            WebDriverWait(self.browser, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH,
                     '//button[normalize-space()="Generate"]')))

            time.sleep(5)
            generate_button_ele = self.browser.find_element_by_xpath(
                '//button[normalize-space()="Generate"]')
            action = ActionChains(self.browser)
            action.move_to_element(generate_button_ele).perform()

            time.sleep(3)

            generate_waybill_button = self.browser.find_element_by_xpath('/html/body/div[5]/ul/li[2]/div[2]/div')

            action.move_to_element(generate_waybill_button).perform()
            generate_waybill_button.click()

            time.sleep(3)

            self.save_pdf(order_id)
            print("Downloaded pdf for order id: {}".format(order_id))

            self.browser.switch_to.window(self.browser.window_handles[0])
            time.sleep(3)
            self.browser.find_element_by_xpath('//div[normalize-space()="Collapse"]').click()

            time.sleep(5)
            self.browser.refresh()
            WebDriverWait(self.browser, 10).until(
                EC.presence_of_element_located(
                    (By.CLASS_NAME, 'mass-ship-list-item')))
        return 1

    def save_pdf(self, order_id):
        # Saving PDF
        self.browser.switch_to.window(self.browser.window_handles[1])
        time.sleep(10)
        pyautogui.rightClick(x=self.windows_size['width'] // 2, y=self.windows_size['height'] // 2)
        pyautogui.typewrite(['a'])
        path = self.format_path([self.homeDirectory, self.pdfFolder, "{0}.pdf".format(order_id)])
        # path = os.getcwd() + '\\waybill_pdf\\{0}.pdf'.format()
        pyperclip.copy(path)
        time.sleep(5)
        pyautogui.hotkey('ctrlleft', 'V')
        pyautogui.press('enter')

        self.browser.close()

    def delivery(self, type_delivery):
        WebDriverWait(self.browser, 10).until(
            EC.presence_of_element_located(
                (By.CLASS_NAME, 'mass-ship-list-item')))

        ret = self.get_orders_generate_pdf()

        if ret == 0:
            del_string = type_delivery.strip().split(' ')
            i = 0
            final_str = ""
            while i < len(del_string) and del_string[i] != '(':
                final_str += del_string[i] + " "
                i += 1
            print("No orders in {0}.".format(final_str))
        # print(self.orderDetails)

    @property
    def get_order_details(self):
        return self.orderDetails

    def format_path(self, pathList):
        path = ""
        for x in pathList:
            path += x
        return os.path.normpath(path)

    def teardown(self):
        """Close browser"""
        try:
            if self.browser is not None:
                self.browser.quit()
        except Exception:
            traceback.print_exc()


class Pdf:
    def __init__(self):
        paths = json.loads(self.read_file("config\\config.json"))[0]
        self.homeDirectory = paths.get("homeDirectory")
        self.jsonPattern = paths.get("jsonPattern")
        self.jsonFolder = paths.get("jsonFolder")
        self.pdfFolder = paths.get("pdfFolder")
        self.readPDFPattern = paths.get("readPDFPattern")
        self.writePDFPattern = paths.get("writePDFPattern")
        self.cpuPDFFile = paths.get("cpuPDFFile")
        self.dic = {}

    def read_file(self, filePath):
        with open(filePath, "r") as file:
            return file.read()

    def get_file(self, file):
        return glob.glob(file)

    def format_path(self, pathList):
        path = ""
        for x in pathList:
            path += x
        return os.path.normpath(path)

    def format_data(self, data):
        orderID = data.get("OrderId", "")
        productStr = ""
        for product in data["Products"]:
            productStr += "{0:>2} x {1:<23}\\n".format(int(product['Quantity']),
                                                       product['Variation'].replace("Variation: ", ""))
        productStr = productStr[:-2]
        return [orderID, productStr]

    def format_pdf(self, pdfCPU, sourcePDFFile, productDetails):
        white_box = 'stamp add -mode text -- "Plain" "font:Courier, rot:0.0, pos:bl, off: 3 5,bo:1 #FFFFFF,fillc:#FFFFFF,bgcol:#FFFFFF, sc: 0.435 rel"'
        cmd = "cmd /c {0} {1} {2}".format(pdfCPU, white_box, sourcePDFFile)
        os.system(cmd)

        write_box_cmd = 'stamp add -mode text --'
        # write_box_param = '"font:Courier, rot:0.0, pos:bl, off: 2 5,bo:1 #FFFFFF,fillc:#000000,bgcol:#FFFFFF, sc: 0.435 rel"'
        write_box_param = '"font:Courier-Bold, rot:0.0, pos:bl, off: 4 4,fillc:#000000, sc: 0.44 rel"'
        cmd = "cmd /c {0} {1} \"{2}\" {3} {4}".format(pdfCPU, write_box_cmd, productDetails, write_box_param, sourcePDFFile)
        os.system(cmd)

    def merger_pdf(self, cpuPDFPath, destinationPDFFile):
        all_pdf = self.get_file(self.format_path([self.homeDirectory, self.pdfFolder, self.readPDFPattern]))
        pdf_str = ""
        for pdf in all_pdf:
            pdf_str = pdf_str + pdf + " "
        cmd = "cmd /c {0} merge {1} {2}".format(cpuPDFPath, destinationPDFFile, pdf_str)
        os.system(cmd)

        for pdf in all_pdf:
            os.remove(pdf)

    def write(self):
        if (self.get_file(self.format_path([self.homeDirectory, self.pdfFolder, self.readPDFPattern]))):
            for file in self.get_file(self.format_path([self.homeDirectory, self.jsonFolder, self.jsonPattern])):
                fileData = json.loads(self.read_file(file))
                for data in fileData:
                    orderID, productStr = self.format_data(data)
                    self.dic[orderID] = productStr

            cpuPDFPath = self.format_path([self.homeDirectory, self.cpuPDFFile])
            for pdfFile in self.get_file(
                    self.format_path([self.homeDirectory, self.pdfFolder, self.readPDFPattern])):
                pdfName = pdfFile.split("\\")[-1]
                pdf = pdfName.split(".")[0]
                productDetails = self.dic.get(pdf, "")
                if productDetails:
                    self.format_pdf(cpuPDFPath, pdfFile, productDetails)

            now = datetime.now()
            dt_string = "{0}{1}{2}".format("merge(", now.strftime("%d-%B-%Y_%H.%M.%S"), ").pdf")
            destinationPDFFile = self.format_path([self.homeDirectory, self.writePDFPattern, dt_string])
            self.merger_pdf(cpuPDFPath, destinationPDFFile)


class Setup:
    def __init__(self):
        paths = ""
        with open("config\\config.json", "r") as file:
            paths = json.loads(file.read())[0]
        self.homeDirectory = paths.get("homeDirectory")
        self.jsonFolder = paths.get("jsonFolder")
        self.writePDFPattern = paths.get("writePDFPattern")
        self.pdfFolder = paths.get("pdfFolder")

    def checkDir(self, file):
        if not os.path.isdir(os.path.normpath(self.homeDirectory + file)):
            os.mkdir(file.replace("/", ""))

    def run(self):
        self.checkDir(self.jsonFolder)
        self.checkDir(self.writePDFPattern)
        self.checkDir(self.pdfFolder)


if __name__ == "__main__":
    s = Setup()
    s.run()

    shopeeAutomate = ShopeeAutomation()
    pdf = Pdf()
    shopeeAutomate.setup()
    file = ""
    try:
        file = pdf.format_path([s.homeDirectory, s.jsonFolder, "order_details.json"])
        shopeeAutomate.run()
        shopeeAutomate.teardown()
    except Exception:
        traceback.print_exc()
    finally:
        with open(file, 'w') as outfile:
            json.dump(shopeeAutomate.get_order_details, outfile)
            shopeeAutomate.teardown()
        pdf.write()
