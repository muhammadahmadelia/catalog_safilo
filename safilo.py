# from email.mime import image
import os
import re
import sys
import json
from time import sleep
from urllib.parse import quote
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime
# import chromedriver_autoinstaller

from models.store import Store
from models.brand import Brand
from models.product import Product
from models.variant import Variant
# from models.metafields import Metafields
import glob
import requests
from bs4 import BeautifulSoup
# import pandas as pd
import threading
from urllib import parse
from openpyxl import Workbook
from openpyxl.drawing.image import Image as Imag
# from openpyxl.utils import get_column_letter
from PIL import Image

from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select

import warnings
warnings.filterwarnings("ignore")

class myScrapingThread(threading.Thread):
    def __init__(self, threadID: int, name: str, obj, brand_name: str, product_url: str, product_number: str, headers: dict, glasses_type: str) -> None:
        threading.Thread.__init__(self)
        self.threadID = threadID
        self.name = name
        self.brand_name = brand_name
        self.glasses_type = glasses_type
        self.product_number = product_number
        self.product_url = product_url
        self.headers = headers
        self.obj = obj
        self.status = 'in progress'
        pass

    def run(self):
        self.obj.scrape_product(self.brand_name, self.product_url, self.product_number, self.headers, self.glasses_type)
        self.status = 'completed'

    def active_threads(self):
        return threading.activeCount()

class Safilo_Scraper:
    def __init__(self, DEBUG: bool, result_filename: str, logs_filename: str) -> None:
        self.DEBUG = DEBUG
        self.result_filename = result_filename
        self.logs_filename = logs_filename
        self.thread_list = []
        self.thread_counter = 0
        self.chrome_options = Options()
        self.chrome_options.add_argument('--disable-infobars')
        self.chrome_options.add_argument("--start-maximized")
        self.chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
        self.args = ["hide_console", ]
        # self.browser = webdriver.Chrome(options=self.chrome_options, service_args=self.args)
        self.browser = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=self.chrome_options)
        self.data = []
        self.ref_json_data = None
        pass

    def controller(self, store: Store, brands_with_types: list[dict]):
        try:
            cookies = ''

            self.browser.get(store.link)
            self.wait_until_browsing()
            self.accept_cookies()

            if self.login(store.username, store.password):
                self.wait_for_page_loading()

                self.select_language()
                sleep(0.8)
                self.wait_for_page_loading()

                print('Scraping products for')
                for brand_with_type in brands_with_types:
                    brand_name: str = brand_with_type['brand']
                    brand_code: str = brand_with_type['code']
                    # print(f'Brand: {brand_name}')

                    for glasses_type in brand_with_type['glasses_type']:

                        if self.select_brand(brand_name):
                            self.select_sunglasses_category(str(glasses_type).strip().title())
                            self.load_all_products()

                            print(f'Brand: {brand_name} | Type: {str(glasses_type).strip().title()}')

                            self.wait_until_element_found(40, 'xpath', '//div[@class="productListContent cc_results_list cc_grid_container"]/span[@class="cc_product_container productFlexItem"]')

                            total_products = self.get_total_products()
                            scraped_products = 0

                            print(f'Type: {glasses_type} | Total products: {total_products}')

                            brand_url = str(self.browser.current_url).strip()

                            for product_span in self.browser.find_elements(By.XPATH, '//div[@class="productListContent cc_results_list cc_grid_container"]/span[@class="cc_product_container productFlexItem"]'):
                                scraped_products += 1
                                ActionChains(self.browser).move_to_element(product_span).perform()
                                
                                product_url, product_name, product_number = self.get_product_data(product_span, brand_name, brand_code)

                                if not cookies: cookies = self.get_cookies_from_browser()
                                headers = self.get_headers(cookies, brand_url)
                                # self.scrape_product(brand, product_url, product_number, headers, glasses_type)
                                self.create_thread(brand_name, product_url, product_number, headers, glasses_type)
                                if self.thread_counter >= 10: 
                                    self.wait_for_thread_list_to_complete()
                                    self.save_to_json(self.data)
                                # self.save_to_json(self.data)
                                # self.printProgressBar(scraped_products, total_products, prefix = 'Progress:', suffix = 'Complete', length = 50)
                            
                            self.wait_for_thread_list_to_complete()
                            self.save_to_json(self.data)

                        self.browser.get(brand_url)
                        self.wait_until_browsing()

                    self.wait_for_thread_list_to_complete()
                    self.save_to_json(self.data)
            else: print(f'Failed to login \nURL: {store.link}\nUsername: {str(store.username)}\nPassword: {str(store.password)}')
        except Exception as e:
            if self.DEBUG: print(f'Exception in scraper controller: {e}')
            else: pass
        finally: 
            self.wait_for_thread_list_to_complete()
            self.save_to_json(self.data)
            self.browser.quit()

    def wait_until_browsing(self) -> None:
        while True:
            try:
                state = self.browser.execute_script('return document.readyState; ')
                if 'complete' == state: break
                else: sleep(0.2)
            except: pass

    def wait_until_element_found(self, wait_value: int, type: str, value: str) -> bool:
        flag = False
        try:
            if type == 'id':
                WebDriverWait(self.browser, wait_value).until(EC.presence_of_element_located((By.ID, value)))
                flag = True
            elif type == 'xpath':
                WebDriverWait(self.browser, wait_value).until(EC.presence_of_element_located((By.XPATH, value)))
                flag = True
            elif type == 'css_selector':
                WebDriverWait(self.browser, wait_value).until(EC.presence_of_element_located((By.CSS_SELECTOR, value)))
                flag = True
            elif type == 'class_name':
                WebDriverWait(self.browser, wait_value).until(EC.presence_of_element_located((By.CLASS_NAME, value)))
                flag = True
            elif type == 'tag_name':
                WebDriverWait(self.browser, wait_value).until(EC.presence_of_element_located((By.TAG_NAME, value)))
                flag = True
        except: pass
        finally: return flag

    def accept_cookies(self) -> None:
        try:
            # accept cookies if found
            if self.wait_until_element_found(30, 'xpath', '//button[@id="acceptCookiesPolicy"]'):
                for _ in range(0, 20):
                    try:
                        self.browser.find_element(By.XPATH,'//button[@id="acceptCookiesPolicy"]').click()
                        sleep(0.2)
                        break
                    except: sleep(0.5)
        except Exception as e:
            self.print_logs(f'Exception in accept_cookies: {str(e)}')
            if self.DEBUG: print(f'Exception in accept_cookies: {str(e)}')
            else: pass

    def login(self, email: str, password: str) -> bool:
        login_flag = False
        try:
            if self.wait_until_element_found(20, 'xpath', '//input[@id="emailField"]'):
                self.browser.find_element(By.XPATH, '//input[@id="emailField"]').send_keys(email)
                sleep(0.2)
                if self.wait_until_element_found(20, 'xpath', '//input[@id="passwordField"]'):
                    self.browser.find_element(By.XPATH, '//input[@id="passwordField"]').send_keys(password)
                    sleep(0.2)
                    self.browser.find_element(By.XPATH, '//input[@id="send2Dsk"]').click()

                    if self.wait_until_element_found(20, 'xpath', '//ul[@data-value="Marchi"]'): login_flag = True
                else: print('Password input not found')
            else: print('Email input not found')
        except Exception as e:
            self.print_logs(f'Exception in login: {str(e)}')
            if self.DEBUG: print(f'Exception in login: {str(e)}')
            else: pass
        finally: return login_flag
    
    def wait_for_page_loading(self):
        self.wait_until_browsing()
        for _ in range(0, 100):
            try:
                self.browser.find_element(By.XPATH, '//div[@id="overlay"]')
                sleep(0.5)
            except: break

    def select_language(self):
        try:
            self.browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            self.browser.find_element(By.XPATH, '//a[@class="changeLocale cc_change_locale link"]').click()
            for _ in range(0, 100):
                try:
                    if 'display: none;' not in self.browser.find_element(By.XPATH, '//div[@id="switcherMod"]').get_attribute('style'):
                        select = Select(self.browser.find_element(By.XPATH, '//select[@class="selectpicker localeSelector"]'))
                        select.select_by_value('en_US')
                        self.browser.find_element(By.CSS_SELECTOR, 'input[class="btn btn-primary setLocale cc_set_locale"]').click()
                        sleep(1)
                        break
                    else: sleep(0.2)
                except: sleep(0.2)
        except Exception as e:
            self.print_logs(f'Exception in select_language: {str(e)}')
            if self.DEBUG: print(f'Exception in select_language: {str(e)}')
            else: pass

    def select_brand(self, brand_name: str):
        flag = False
        self.wait_until_element_found(40, 'xpath', '//ul[@class="nav navbar-nav cc_navbar-nav"]/li//a[contains(text(), "Brands")]')
        try:
            brand_li = self.browser.find_element(By.XPATH, '//ul[@class="nav navbar-nav cc_navbar-nav"]/li//a[contains(text(), "Brands")]')
            ActionChains(self.browser).move_to_element(brand_li).perform()
            main_ul = self.browser.find_element(By.XPATH, '//ul[@data-value="Brands"]')
            for ul in main_ul.find_elements(By.TAG_NAME, 'ul'):
                if ul.get_attribute('id') in ['saf_pur', 'exc_pur', 'exc_visible']:
                    for li in ul.find_elements(By.TAG_NAME, 'li'):
                        if str(li.find_element(By.TAG_NAME, 'a').text).strip().lower() == str(brand_name).strip().lower():
                            li.click()
                            self.wait_until_browsing()
                            self.wait_for_page_loading()
                            flag = True
                            break
                if flag: break
        except Exception as e:
            self.print_logs(f'Exception in select_brand: {str(e)}')
            if self.DEBUG: print(f'Exception in select_brand: {str(e)}')
            else: pass
            flag = False
        finally: return flag
        
    def select_sunglasses_category(self, glasses_type: str):
        try:
            for a in self.browser.find_elements(By.XPATH, '//a[@class="cc_collapse_group"]'):
                if str('Product Type').strip().lower() in str(a.text).strip().lower():
                    a.click()
                    sleep(0.3)
                    xpath_glasses_type = ''
                    if glasses_type == 'Sunglasses': xpath_glasses_type = '//input[@data-value="SUN"]'
                    elif glasses_type == 'Eyeglasses': xpath_glasses_type = '//input[@data-value="EYE"]'
                    elif glasses_type == 'Ski & Snowboard Goggles': xpath_glasses_type = '//input[@data-value="SPO"]'
                    checkbox = self.browser.find_element(By.XPATH, xpath_glasses_type)
                    self.browser.execute_script("arguments[0].scrollIntoView();", checkbox)
                    ActionChains(self.browser).move_to_element(checkbox).click().perform()
                    sleep(0.2)
                    self.wait_for_page_loading()
                    break
        except Exception as e:
            self.print_logs(f'Exception in select_sunglasses_category: {str(e)}')
            if self.DEBUG: print(f'Exception in select_sunglasses_category: {str(e)}')
            else: pass

    def load_all_products(self):
        while True:
            try:
                self.browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                self.browser.find_element(By.XPATH, '//div[@class="row cc_list_footer"]/div/button').click()
                self.wait_for_page_loading()
                sleep(0.4)
            except: break

    def get_product_data(self, product_span, brand_name: str, brand_code: str):
        product_url, product_name, product_number = '', '', ''
        try:
            product_url = product_span.find_element(By.XPATH, './/h5[@class="cc_product_link"]/a').get_attribute('href')
            text = str(product_span.find_element(By.XPATH, './/h5[@class="cc_product_link"]/a').text).strip()
            text = str(text).lower().replace(str(brand_name).strip().lower(), '').strip()
            text = str(text).lower().replace(str(brand_code).strip().lower(), '').strip()
            product_name = str(text).strip().title()
            product_number = str(product_span.find_element(By.XPATH, './/h5[@class="cc_product_link"]/a').get_attribute('data-id')).strip()
        except Exception as e:
            self.print_logs(f'Exception in get_product_data: {str(e)}')
            if self.DEBUG: print(f'Exception in get_product_data: {str(e)}')
            else: pass
        finally: return product_url, product_name, product_number

    def get_cookies_from_browser(self) -> str:
        cookies = ''
        try:
            browser_cookies = self.browser.get_cookies()
        
            for browser_cookie in browser_cookies:
                if browser_cookie['name'] == '_hjAbsoluteSessionInProgress': cookies = f'_hjAbsoluteSessionInProgress=0; {browser_cookie["name"]}={browser_cookie["value"]}; {cookies}'
                else: cookies = f'{browser_cookie["name"]}={browser_cookie["value"]}; {cookies}'
            cookies = cookies.strip()[:-1]
        except Exception as e:
            if self.DEBUG: print(f'Exception in get_cookies_from_browser: {e}')
            self.print_logs(f'Exception in get_cookies_from_browser: {e}')
        finally: return cookies

    def get_headers(self, cookies: str, referer: str) -> dict:
        return {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'Accept-Encoding': 'gzip, deflate, br',
            'Accept-Language': 'en-US,en;q=0.9',
            'Cache-Control': 'max-age=0',
            'Connection': 'keep-alive',
            'Cookie': cookies,
            'Host': 'www.youandsafilo.com',
            'Referer': referer,
            'sec-ch-ua': '"Not?A_Brand";v="8", "Chromium";v="108", "Google Chrome";v="108"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'Sec-Fetch-Dest': 'document',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-Site': 'same-origin',
            'Sec-Fetch-User': '?1',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': 'PostmanRuntime/7.29.2'
        }
    
    def scrape_product(self, brand_name: str, product_url: str, product_number: str, headers: dict, glasses_type: str):
        try:
            main_json_data = {}
            response = None
            for _ in range(0, 10):
                try:
                    response = requests.get(url=product_url, headers=headers, verify=False)
                    break
                except requests.exceptions.Timeout: sleep(1)
                except requests.exceptions.ConnectionError: sleep(1)
                except Exception as e:
                    self.print_logs(f'Exception in scrape_product request: {e}')
                    break

            if response and response.status_code == 200:
                
                soup = BeautifulSoup(response.text, 'lxml')
                for script in soup.select('script'):
                    if 'CCRZ.detailData.jsonProductData =' in str(script.text).strip():
                        value = str(script.text).split('CCRZ.detailData.jsonProductData =')[1]
                        value = str(value).split('CCRZ.detailData.pageSections = [];')[0].strip()[:-1]
                        main_json_data = json.loads(value)

                if main_json_data:
                    id, main_category, product_type = self.get_variables_for_new_request(main_json_data)

                    if id and main_category and product_type:
                        required_json = self.get_required_json(soup, product_url, id, main_category, product_type)
                        payload = self.get_payload_for_request(required_json)
                        json_dataz = self.get_other_json(payload, headers['Cookie'], required_json['cartId'], required_json['effectiveAccount'], product_number)
                        if not self.ref_json_data: self.get_ref_data(headers)

                        product_json = main_json_data['product']
                        prodBean = product_json['prodBean']
                        
                        product_name = str(prodBean['name']).replace(prodBean['b2BBrandLabel'], '').strip()
                        # product_name = self.clean_product_name(product_name)

                        for_who, frame_material, frame_shape = self.get_metafields(prodBean)
                        
                        b2BRetailPrice = ''
                        try: b2BRetailPrice = float(prodBean['b2BRetailPriceItemS'][0]['b2BRetailPrice'])
                        except: pass

                        whole_price = ''
                        try: whole_price = float(product_json['price'])
                        except: pass

                        products: list[Product] = []

                        for value in prodBean['compositeProductsSByColor']:
                            frame_codes = []
                            for somevalue in prodBean['compositeProductsSByColor'][value]['compositeProductsS']:
                                frame_code = somevalue['b2BColorCode']

                                if frame_code not in frame_codes:
                                    try:
                                        frame_codes.append(frame_code)
                                        price = None
                                        # get new product data
                                        product = self.get_product(brand_name, product_number, product_name, frame_code, somevalue, glasses_type)
                                        
                                        product.metafields.for_who = for_who
                                        product.metafields.lens_material = self.get_lens_material(somevalue)
                                        product.metafields.frame_material = frame_material
                                        product.metafields.frame_shape = frame_shape

                                        # get frame color and price against frame code
                                        images = []
                                        if product.frame_code: 
                                            product.metafields.frame_color, price, price2, images = self.get_frame_color(json_dataz, product.frame_code)
                                        
                                        if images: product.images_360 = images
                                        
                                        self.get_product_images(product)

                                        bridge, template = self.get_bridge_template(somevalue)
                                        if not product.bridge: product.bridge = bridge
                                        if not product.template: product.template = template
                                        
                                        variant = self.get_variant_data(somevalue)                                    
                                        if price: variant.listing_price = price
                                        else: variant.listing_price = b2BRetailPrice

                                        if price2: variant.wholesale_price = price2
                                        else: variant.wholesale_price = whole_price

                                        product.add_single_variant(variant)

                                        products.append(product)
                                    except Exception as e:
                                        if self.DEBUG: print(f'Exception in new_product adding: {e}')
                                        self.print_logs(f'Exception in new_product adding: {e}')
                                else:
                                    try:
                                        variant = self.get_variant_data(somevalue)

                                        for product in products:
                                            if product.frame_code == frame_code:
                                                flag = None
                                                if 'eyeglasses' in str(glasses_type).strip().lower():
                                                    flag = bool(str(product.lens_code).strip().upper() == str(variant.sku)[-2:].strip())
                                                else:
                                                    flag = bool(str(product.lens_code).strip() and str(product.lens_code).strip().upper() == str(variant.sku)[-2:].strip())

                                                if flag:
                                                    variant.listing_price = product.variants[0].listing_price
                                                    variant.wholesale_price = product.variants[0].wholesale_price
                                                    product.add_single_variant(variant)
                                                else:
                                                    flag = True
                                                    for product_variant in product.variants:
                                                        if str(variant.title).strip() == str(product_variant.title).strip():
                                                            flag = False
                                                            break
                                                    if flag: 
                                                        product.add_single_variant(variant)
                                                        if not product.bridge or not product.template:
                                                            bridge, template = self.get_bridge_template(somevalue)
                                                            if not product.bridge: product.bridge = bridge
                                                            if not product.template: product.template = template
                                    except Exception as e:
                                        if self.DEBUG: print(f'Exception in new_variant adding: {e}')
                                        self.print_logs(f'Exception in new_variant adding: {e}')

                        try:
                            for product in products:
                                gtins, product_sizes = [], []
                                for variant in product.variants:
                                    if variant.barcode_or_gtin: gtins.append(variant.barcode_or_gtin)
                                    if variant.title: product_sizes.append(variant.size)

                                if product_sizes: product.metafields.size_bridge_template = ', '.join(product_sizes)
                                if gtins: product.metafields.gtin1 = ', '.join(gtins)
                        except Exception as e:
                            if self.DEBUG: print(f'Exception in adding gtin and size: {e}')
                            self.print_logs(f'Exception in adding gtin and size: {e}')

                        for product in products: self.data.append(product)

        except Exception as e:
            if self.DEBUG: print(f'Exception in scrape_product: {e}')
            self.print_logs(f'Exception in scrape_product: {e}')

    def get_variables_for_new_request(self, json_data: dict) -> list[str]:
        id, main_category, product_type = '', '', ''
        try:
            product_json = json_data['product']
            prodBean = product_json['prodBean']            
            id = str(prodBean['id']).strip()
            main_category = str(prodBean['mainCategory']).strip()
            product_type = str(prodBean['ProductType']).strip()
        except Exception as e:
            if self.DEBUG: print(f'Exception in get_variables_for_new_request: {e}')
            self.print_logs(f'Exception in get_variables_for_new_request: {e}')
        finally: return [id, main_category, product_type]

    def get_required_json(self, soup: BeautifulSoup, url: str, id: str, main_category: str, product_type: str) -> dict:
        json_required_for_payload = {}
        try:
            number, cartId, store, effectiveAccount, cclcl, role = self.get_values_from_url(url)
            for script in soup.select('script'):
                if 'Visualforce.remoting.Manager.add' in str(script.text).strip():
                    value = str(script.text).strip().split('Visualforce.remoting.Manager.add(new $VFRM.RemotingProviderImpl(')[-1]
                    json_data = json.loads(value[:-3])

                    MenuBar_method, MenuBar_ns, MenuBar_csrf, MenuBar_authorization, MenuBar_ver =  '', '', '', '', 0
                    
                    for value in json_data['actions']['ccrz.cc_ctrl_MenuBar']["ms"]:
                        if value['name'] == 'getMenuJson':
                            MenuBar_method = str(value['name']).strip()
                            MenuBar_ns = str(value['ns']).strip()
                            MenuBar_ver = int(value['ver'])
                            MenuBar_csrf = str(value['csrf']).strip()
                            MenuBar_authorization = str(value['authorization']).strip()

                    ProductDetailRD_method, ProductDetailRD_ns, ProductDetailRD_csrf, ProductDetailRD_authorization, ProductDetailRD_ver = '', '', '', '', 0
                    for value in json_data['actions']['ccrz.cc_ctrl_ProductDetailRD']["ms"]:
                        if value['name'] == 'fetchCompositeProducts':
                            ProductDetailRD_method = str(value['name']).strip()
                            ProductDetailRD_ns = str(value['ns']).strip()
                            ProductDetailRD_ver = int(value['ver'])
                            ProductDetailRD_csrf = str(value['csrf']).strip()
                            ProductDetailRD_authorization = str(value['authorization']).strip()

                    vid = str(json_data['vf']['vid']).strip()

                    json_required_for_payload = {
                        'number': number,
                        'cartId': cartId,
                        'store': store, 
                        'effectiveAccount': effectiveAccount,
                        'cclcl': cclcl,
                        'role': role,
                        'vid': vid,
                        'cc_ctrl_MenuBar': {
                            'MenuBar_method': MenuBar_method,
                            'MenuBar_ns': MenuBar_ns,
                            'MenuBar_ver': MenuBar_ver,
                            'MenuBar_csrf': MenuBar_csrf,
                            'MenuBar_authorization': MenuBar_authorization
                        },
                        'cc_ctrl_ProductDetailRD': {
                            'ProductDetailRD_method': ProductDetailRD_method,
                            'ProductDetailRD_ns': ProductDetailRD_ns,
                            'ProductDetailRD_ver': ProductDetailRD_ver,
                            'ProductDetailRD_csrf': ProductDetailRD_csrf,
                            'ProductDetailRD_authorization': ProductDetailRD_authorization
                        },
                        'id': id,
                        'main_category': main_category,
                        'product_type': product_type
                        
                    }
        except Exception as e:
            if self.DEBUG: print(f'Exception in get_required_json: {e}')
            self.print_logs(f'Exception in get_required_json: {e}')
        finally: return json_required_for_payload

    def get_values_from_url(self, url: str) -> list[str]:
        number, cartId, store, effectiveAccount, cclcl, role = '', '', '', '', '', ''
        try:
            number = parse.parse_qs(parse.urlparse(url).query)['sku'][0]
            cartId = parse.parse_qs(parse.urlparse(url).query)['cartId'][0]
            store = parse.parse_qs(parse.urlparse(url).query)['store'][0]
            effectiveAccount = parse.parse_qs(parse.urlparse(url).query)['effectiveAccount'][0]
            cclcl = parse.parse_qs(parse.urlparse(url).query)['cclcl'][0]
            role = parse.parse_qs(parse.urlparse(url).query)['role'][0]
        except Exception as e:
            if self.DEBUG: print(f'Exception in get_values_from_url: {e}')
            self.print_logs(f'Exception in get_values_from_url: {e}')
        finally: return [number, cartId, store, effectiveAccount, cclcl, role]

    def get_payload_for_request(self, json_data: dict) -> list[dict]:
        return [
            {
                "action":"ccrz.cc_ctrl_MenuBar",
                "method":json_data['cc_ctrl_MenuBar']['MenuBar_method'],
                "data": [
                    {
                        "storefront":json_data['store'],
                        "portalUserId":"",
                        "effAccountId":json_data['effectiveAccount'],
                        "priceGroupId":"",
                        "currentCartId":json_data['cartId'],
                        "userIsoCode":"EUR",
                        "userLocale":json_data['cclcl'],
                        "currentPageName":"ccrz__ProductDetails",
                        "currentPageURL":f"https://www.youandsafilo.com/ccrz__ProductDetails?cartId={json_data['cartId']}&cclcl={json_data['cclcl']}&effectiveAccount={json_data['effectiveAccount']}&role={json_data['role']}&sku={json_data['number']}&store={json_data['store']}&refURL=https%3A%2F%2Fwww.youandsafilo.com%2Fccrz__ProductList%3FcategoryId%3D{json_data['main_category']}%26portalUser%3D%26store%3D{json_data['store']}%26effectiveAccount%3D{json_data['effectiveAccount']}%26cclcl%3Den_US%26role%3DS2",
                        "queryParams":{
                            "sku":json_data['number'],
                            "cartId":json_data['cartId'],
                            "store":json_data['store'],
                            "effectiveAccount":json_data['effectiveAccount'],
                            "cclcl":json_data['cclcl'],
                            "role":json_data['role'],
                        }
                    }
                ],
                "type":"rpc",
                "tid":6,
                "ctx":{
                    "csrf":json_data['cc_ctrl_MenuBar']['MenuBar_csrf'],
                    "vid":json_data['vid'],
                    "ns":json_data['cc_ctrl_MenuBar']['MenuBar_ns'],
                    "ver":json_data['cc_ctrl_MenuBar']['MenuBar_ver'],
                    "authorization":json_data['cc_ctrl_MenuBar']['MenuBar_authorization']
                }
            },
            {
                "action":"ccrz.cc_ctrl_ProductDetailRD",
                "method":json_data['cc_ctrl_ProductDetailRD']['ProductDetailRD_method'],
                "data":[
                    {
                        "storefront":json_data['store'],
                        "portalUserId":"",
                        "effAccountId":json_data['effectiveAccount'],
                        "priceGroupId":"",
                        "currentCartId":json_data['cartId'],
                        "userIsoCode":"EUR",
                        "userLocale":json_data['cclcl'],
                        "currentPageName":"ccrz__ProductDetails",
                        "currentPageURL":f"https://www.youandsafilo.com/ccrz__ProductDetails?cartId={json_data['cartId']}&cclcl={json_data['cclcl']}&effectiveAccount={json_data['effectiveAccount']}&role={json_data['role']}&sku={json_data['number']}&store={json_data['store']}&refURL=https%3A%2F%2Fwww.youandsafilo.com%2Fccrz__ProductList%3FcategoryId%3D{json_data['main_category']}%26portalUser%3D%26store%3D{json_data['store']}%26effectiveAccount%3D{json_data['effectiveAccount']}%26cclcl%3D{json_data['cclcl']}%26role%3D{json_data['role']}",
                        "queryParams":{
                            "sku":json_data['number'],
                            "cartId":json_data['cartId'],
                            "store":json_data['store'],
                            "effectiveAccount":json_data['effectiveAccount'],
                            "cclcl":json_data['cclcl'],
                            "role":json_data['role']
                        }
                    },
                    json_data['id'], json_data['product_type']
                ],
                "type":"rpc",
                "tid":9,
                "ctx":{
                    "csrf":json_data['cc_ctrl_ProductDetailRD']['ProductDetailRD_csrf'],
                    "vid":json_data['vid'],
                    "ns":json_data['cc_ctrl_ProductDetailRD']['ProductDetailRD_ns'],
                    "ver":json_data['cc_ctrl_ProductDetailRD']['ProductDetailRD_ver'],
                    "authorization":json_data['cc_ctrl_ProductDetailRD']['ProductDetailRD_authorization']
                }
            }
        ]

    def  get_other_json(self, payload: list[dict], cookies: str, cartId: str, effectiveAccount: str, product_number: str) -> dict:
        json_data = {}
        try:
            headers = {
                'Accept': '*/*',
                'Accept-Encoding': 'gzip, deflate, br',
                'Accept-Language': 'en-US,en;q=0.9',
                'Connection': 'keep-alive',
                'Content-Length': '3169',
                'Content-Type': 'application/json',
                'Cookie': cookies,
                'Host': 'www.youandsafilo.com',
                'Origin': 'https://www.youandsafilo.com',
                'Referer': f'https://www.youandsafilo.com/ccrz__ProductDetails?cartId={cartId}&store=Safilo&effectiveAccount={effectiveAccount}&cclcl=en_US&role=S2&sku={product_number}',
                'sec-ch-ua': '"Not?A_Brand";v="8", "Chromium";v="108", "Google Chrome";v="108"',
                'sec-ch-ua-mobile': '?0',
                'sec-ch-ua-platform': '"Windows"',
                'Sec-Fetch-Dest': 'empty',
                'Sec-Fetch-Mode': 'cors',
                'Sec-Fetch-Site': 'same-origin',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36',
                'X-Requested-With': 'XMLHttpRequest',
                'X-User-Agent': 'Visualforce-Remoting'
            }
            for _ in range(0, 10):
                try:
                    response = requests.post(url='https://www.youandsafilo.com/apexremote', json=payload, headers=headers, verify=False)
                    if response.status_code == 200:
                        json_data = json.loads(response.text)
                    break
                except requests.exceptions.Timeout: sleep(1)
                except requests.exceptions.ConnectionError: sleep(1)
                except Exception as e:
                    self.print_logs(f'Exception in get_other_json request: {e}')
                    break
        except Exception as e:
            if self.DEBUG: print(f'Exception in get_other_json: {e}')
            self.print_logs(f'Exception in get_other_json: {e}')
        finally: return json_data

    def get_ref_data(self, headers: dict) -> None:
        try:
            url = 'https://www.youandsafilo.com/ccrz__PageLabels?storefront=Safilo&pageName=ProductDetails&userLocale=en_US&pageKey=&trg='
            for _ in range(0, 10):
                try:
                    response = requests.get(url=url, headers=headers, verify=False)
                    if response.status_code == 200:
                        text = str(response.text).strip().replace('var CCRZ=CCRZ||{};CCRZ.pagevars=CCRZ.pagevars||{};CCRZ.pagevars.pageLabels=', '')
                        self.ref_json_data = json.loads(text)
                    break
                except requests.exceptions.Timeout: sleep(1)
                except requests.exceptions.ConnectionError: sleep(1)
                except Exception as e:
                    self.print_logs(f'Exception in get_ref_data request: {e}')
                    break
        except Exception as e:
            if self.DEBUG: print(f'Exception in get_ref_data: {e}')
            self.print_logs(f'Exception in get_ref_data: {e}')

    def get_frame_color(self, json_data: dict, frame_code: str):
        listing_price, wholesale_price, frame_color = None, None, ''
        images = []
        try:
            for json_d in json_data:
                if json_d['method'] == 'fetchCompositeProducts':
                    for v in json_d['result']['data']['v']:
                        product = v['v']['product']
                        prodBean = product['v']['prodBean']['v']

                        if frame_code == str(prodBean['b2BColorCode']).strip():
                            frame_color = str(prodBean['shortDesc']).strip()
                            if 'b2BRetailPriceItemS' in prodBean:
                                value = str(prodBean['b2BRetailPriceItemS']['v'][0]['v']['b2BRetailPrice']).strip()
                                value = re.search('\d+', value)
                                if value: listing_price = float(int(value.group()))

                            if 'price' in product['v']:
                                wholesale_price = product['v']['price']

                            try:
                                for value in prodBean['EProductMediasS']['v']:
                                    if str(value['v']['mediaType']) == 'Product Image':
                                        images.append(str(value['v']['URI']))
                            except: pass
        
        except:
            try:
                for json_d_str in str(json_data).strip().split(','):
                    if 'shortDesc' in str(json_d_str).strip():
                        frame_color = str(json_d_str).replace("'shortDesc': ", "").replace("'", "").strip()
                    if 'b2BRetailPrice' in str(json_d_str).strip():
                        value = str(json_d_str).strip()
                        value = re.search('\d+', value)
                        if value: listing_price = float(int(value.group()))
                    if 'price' in str(json_d_str).strip():
                        value = str(json_d_str).strip()
                        value = re.search('\d+', value)
                        if value: wholesale_price = float(int(value.group()))
                        # price = str(json_d_str).replace("'b2BRetailPrice': ", "").replace("'", "").strip()
            except Exception as e:
                if self.DEBUG: print(f'Exception in get_frame_color: {e}')
                self.print_logs(f'Exception in get_frame_color: {e}')
        finally: return frame_color, listing_price, wholesale_price, images

    def get_variant_data(self, somevalue: dict) -> Variant:
        variant = Variant()
        try:
            try: variant.title = str(somevalue["b2BLensWidthSize"]).strip()
            except: pass
            try: variant.sku = str(somevalue['SKU']).strip() if 'SKU' in somevalue else ''
            except: pass
            try: variant.barcode_or_gtin = str(somevalue['b2BEANCode']).strip() if 'b2BEANCode' in somevalue else ''
            except: pass
            try: variant.size = str(f'{somevalue["b2BLensWidthSize"]}-{int(somevalue["b2BBridgeLengthSize"])}-{int(somevalue["b2BTempleLengthSize"])}').strip().replace(' ', '')
            except: pass
            
            try: 
                variant.inventory_quantity = 0
                if int(somevalue['b2BStockValue']) > 0: variant.inventory_quantity = 5
            except: pass
        except Exception as e:
            self.print_logs(f'Exception in get_variant_data: {e}')
            if self.DEBUG: print(f'Exception in get_variant_data: {e}')
        finally: return variant

    def create_thread(self, brand_name: str, product_url: str, product_number: str, headers: dict, glasses_type: str):
        thread_name = "Thread-"+str(self.thread_counter)
        self.thread_list.append(myScrapingThread(self.thread_counter, thread_name, self, brand_name, product_url, product_number, headers, glasses_type))
        self.thread_list[self.thread_counter].start()
        self.thread_counter += 1

    def is_thread_list_complted(self):
        for obj in self.thread_list:
            if obj.status == "in progress":
                return False
        return True

    def wait_for_thread_list_to_complete(self):
        while True:
            result = self.is_thread_list_complted()
            if result: 
                self.thread_counter = 0
                self.thread_list.clear()
                break
            else: sleep(1)

    def get_total_products(self) -> int:
        total_products = 0
        try:
            total_products = len(self.browser.find_elements(By.XPATH, '//div[@class="productListContent cc_results_list cc_grid_container"]/span[@class="cc_product_container productFlexItem"]'))
        except Exception as e:
            self.print_logs(f'Exception in get_total_products: {str(e)}')
            if self.DEBUG: print(f'Exception in get_total_products: {str(e)}')
            else: pass
        finally: return total_products

    def clean_product_name(self, product_name: str) -> None:
        try:
            if 'CA' in product_name.split(' '): product_name = product_name.replace('CA ', '').strip()
            elif 'CARDUC' in product_name.split(' '): product_name = product_name.replace('CARDUC ', '').strip()
            elif 'CF' in product_name.split(' '): product_name = product_name.replace('CF ', '').strip()
            elif 'DB' in product_name.split(' '): product_name = product_name.replace('DB ', '').strip()
            elif 'PLD' in product_name.split(' '): product_name = product_name.replace('PLD ', '').strip()
            elif 'MARC' in product_name.split(' '): product_name = product_name.replace('MARC ', '').strip()
            elif 'MJ' in product_name.split(' '): product_name = product_name.replace('MJ ', '').strip()
        except Exception as e:
            if self.DEBUG: print(f'Exception in clean_product_name: {e}')
            self.print_logs(f'Exception in clean_product_name: {e}')
        finally: return product_name

    def get_metafields(self, prodBean: dict) -> list[str]:
        for_who, frame_material, frame_shape = '', '', ''
        try:
            try:
                b2BTargetGroupCode = prodBean['b2BTargetGroupCode'] # B2B_Segment_3
                for_who = self.ref_json_data[f'B2B_Segment_{b2BTargetGroupCode}']
            except: pass
            try:
                b2BFrameMaterial = prodBean['b2BFrameMaterial'] # B2B_FrameMaterial_EP
                frame_material = self.ref_json_data[f'B2B_FrameMaterial_{b2BFrameMaterial}']
            except: pass
            try:
                b2BFrameShape = prodBean['b2BFrameShape'] # B2B_Shape_RO
                frame_shape = self.ref_json_data[f'B2B_Shape_{b2BFrameShape}']
            except: pass
        except Exception as e:
            if self.DEBUG: print(f'Exception in get_metafields: {e}')
            self.print_logs(f'Exception in get_metafields: {e}')
        finally: return [for_who, frame_material, frame_shape]

    def get_lens_material(self, somevalue: dict) -> str:
        lens_material = ''
        try:
            if 'b2BLensesMaterial' in somevalue:
                b2BLensesMaterial = somevalue['b2BLensesMaterial']
                lens_material = self.ref_json_data[f'B2B_LensesMaterial_{b2BLensesMaterial}']
        except Exception as e:
            if self.DEBUG: print(f'Exception in get_lens_material: {e}')
            self.print_logs(f'Exception in get_lens_material: {e}')
        finally: return lens_material

    def get_product(self, brand_name: str, number: str, name: str, frame_code: str, somevalue: dict, glasses_type: str) -> Product:
        product = Product()
        try:
            product.brand = brand_name
            product.number = number
            product.name = str(name).strip().upper()
            product.frame_code = str(frame_code).strip().upper()
            try: product.lens_code = str(somevalue['b2BLensCode']).strip().upper()
            except: pass
            product.type = glasses_type
        except Exception as e:
            if self.DEBUG: print(f'Exception in get_product: {e}')
            self.print_logs(f'Exception in get_product: {e}')
        finally: return product

    def get_product_images(self, product: Product) -> None:
        try:
            product.image = f'https://safilo-spa-pd-cde002.azureedge.net/damapi/damimage/public/sfcc.getimagenofb?modelCode={product.number}&colorCode={product.frame_code}&lensCode={product.lens_code}&detail=00&imagesize=medium'
            if not product.images_360:
                product.images_360 = [
                    f'https://safilo-spa-pd-cde002.azureedge.net/damapi/damimage/public/sfcc.getimagenofb?modelCode={product.number}&colorCode={product.frame_code}&lensCode={product.lens_code}&detail=02&imagesize=big',
                    f'https://safilo-spa-pd-cde002.azureedge.net/damapi/damimage/public/sfcc.getimagenofb?modelCode={product.number}&colorCode={product.frame_code}&lensCode={product.lens_code}&detail=03&imagesize=big',
                    f'https://safilo-spa-pd-cde002.azureedge.net/damapi/damimage/public/sfcc.getimagenofb?modelCode={product.number}&colorCode={product.frame_code}&lensCode={product.lens_code}&detail=00&imagesize=big',
                    f'https://safilo-spa-pd-cde002.azureedge.net/damapi/damimage/public/sfcc.getimagenofb?modelCode={product.number}&colorCode={product.frame_code}&lensCode={product.lens_code}&detail=01&imagesize=big',
                    f'https://safilo-spa-pd-cde002.azureedge.net/damapi/damimage/public/sfcc.getimagenofb?modelCode={product.number}&colorCode={product.frame_code}&lensCode={product.lens_code}&detail=07&imagesize=big'
                ]
            
        except Exception as e:
            if self.DEBUG: print(f'Exception in get_product_images: {e}')
            self.print_logs(f'Exception in get_product_images: {e}')

    def get_bridge_template(self, somevalue: dict) -> list[str]:
        bridge, template = '', ''
        try:
            try: bridge = str(int(somevalue["b2BBridgeLengthSize"]))
            except: pass
            try: template = str(int(somevalue["b2BTempleLengthSize"]))
            except: pass
        except Exception as e:
            if self.DEBUG: print(f'Exception in get_bridge_template: {e}')
            self.print_logs(f'Exception in get_bridge_template: {e}')
        finally: return [bridge, template]

    def save_to_json(self, products: list[Product]):
        try:
            json_products = []
            for product in products:
                json_varinats = []
                for index, variant in enumerate(product.variants):
                    json_varinat = {
                        'position': (index + 1), 
                        'title': variant.title, 
                        'sku': variant.sku, 
                        'inventory_quantity': variant.inventory_quantity,
                        'found_status': variant.found_status,
                        'wholesale_price': variant.wholesale_price,
                        'listing_price': variant.listing_price, 
                        'barcode_or_gtin': variant.barcode_or_gtin,
                        'size': variant.size,
                    }
                    json_varinats.append(json_varinat)
                json_product = {
                    'brand': product.brand, 
                    'number': product.number, 
                    'name': product.name, 
                    'frame_code': product.frame_code,  
                    'lens_code': product.lens_code, 
                    'lens_color': product.metafields.lens_color,
                    'frame_color': product.metafields.frame_color,
                    # 'status': product.status, 
                    # 'type': product.type, 
                    # 'url': product.url, 
                    'metafields': [
                        { 'key': 'for_who', 'value': product.metafields.for_who },
                        # { 'key': 'product_size', 'value': product.metafields.product_size }, 
                        { 'key': 'lens_material', 'value': product.metafields.lens_material }, 
                        { 'key': 'lens_technology', 'value': product.metafields.lens_technology }, 
                        { 'key': 'frame_material', 'value': product.metafields.frame_material }, 
                        { 'key': 'frame_shape', 'value': product.metafields.frame_shape },
                        { 'key': 'gtin1', 'value': product.metafields.gtin1 }, 
                        { 'key': 'img_url', 'value': product.image }
                    ],
                    'variants': json_varinats
                }
                json_products.append(json_product)
            
           
            with open(self.result_filename, 'w') as f: json.dump(json_products, f)
            
        except Exception as e:
            if self.DEBUG: print(f'Exception in save_to_json: {e}')
            else: pass

    # print logs to the log file
    def print_logs(self, log: str) -> None:
        try:
            with open(self.logs_filename, 'a') as f:
                f.write(f'\n{log}')
        except: pass

def read_data_from_json_file(DEBUG, result_filename: str):
    data = []
    try:
        files = glob.glob(result_filename)
        if files:
            f = open(files[-1])
            json_data = json.loads(f.read())
            products = []

            for json_d in json_data:
                number, frame_code, brand, img_url, frame_color, lens_color = '', '', '', '', '', ''
                # product = Product()
                brand = json_d['brand']
                number = str(json_d['name']).strip().upper()
                if '-' in number: number = number.replace('-', '/').strip()
                # product.name = str(json_d['name']).strip().upper()
                frame_code = str(json_d['frame_code']).strip().upper()
                if '-' in frame_code: frame_code = frame_code.replace('-', '/').strip()
                frame_color = str(json_d['frame_color']).strip().title()
                # product.lens_code = str(json_d['lens_code']).strip().upper()
                lens_color = str(json_d['lens_color']).strip().title()
                # product.status = str(json_d['status']).strip().lower()
                # product.type = str(json_d['type']).strip().title()
                # product.url = str(json_d['url']).strip()
                # metafields = Metafields()
                
                for json_metafiels in json_d['metafields']:
                    # if json_metafiels['key'] == 'for_who':metafields.for_who = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'product_size':metafields.product_size = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'activity':metafields.activity = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'lens_material':metafields.lens_material = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'graduabile':metafields.graduabile = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'interest':metafields.interest = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'lens_technology':metafields.lens_technology = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'frame_material':metafields.frame_material = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'frame_shape':metafields.frame_shape = str(json_metafiels['value']).strip().title()
                    # elif json_metafiels['key'] == 'gtin1':metafields.gtin1 = str(json_metafiels['value']).strip().title()
                    if json_metafiels['key'] == 'img_url':img_url = str(json_metafiels['value']).strip()
                    # elif json_metafiels['key'] == 'img_360_urls':
                    #     value = str(json_metafiels['value']).strip()
                    #     if '[' in value: value = str(value).replace('[', '').strip()
                    #     if ']' in value: value = str(value).replace(']', '').strip()
                    #     if "'" in value: value = str(value).replace("'", '').strip()
                    #     for v in value.split(','):
                    #         metafields.img_360_urls = str(v).strip()
                # product.metafields = metafields
                for json_variant in json_d['variants']:
                    sku, price = '', ''
                    # variant = Variant()
                    # variant.position = json_variant['position']
                    # variant.title = str(json_variant['title']).strip()
                    sku = str(json_variant['sku']).strip().upper()
                    if '/' in sku: sku = sku.replace('/', '-').strip()
                    # variant.inventory_quantity = json_variant['inventory_quantity']
                    # variant.found_status = json_variant['found_status']
                    wholesale_price = str(json_variant['wholesale_price']).strip()
                    listing_price = str(json_variant['listing_price']).strip()
                    # variant.barcode_or_gtin = str(json_variant['barcode_or_gtin']).strip()
                    # variant.size = str(json_variant['size']).strip()
                    # variant.weight = str(json_variant['weight']).strip()
                    # product.variants = variant

                    image_attachment = download_image(img_url)
                    if image_attachment:
                        with open(f'Images/{sku}.jpg', 'wb') as f: f.write(image_attachment)
                        crop_downloaded_image(f'Images/{sku}.jpg')

                    data.append([brand, number, frame_code, frame_color, lens_color,  sku, wholesale_price, listing_price])
    except Exception as e:
        if DEBUG: print(f'Exception in read_data_from_json_file: {e}')
        else: pass
    finally: return data

def download_image(url):
    image_attachment = ''
    try:
        headers = {
            'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'accept-Encoding': 'gzip, deflate, br',
            'accept-Language': 'en-US,en;q=0.9',
            'cache-Control': 'max-age=0',
            'sec-ch-ua': '"Google Chrome";v="95", "Chromium";v="95", ";Not A Brand";v="99"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'document',
            'sec-fetch-mode': 'navigate',
            'sec-fetch-site': 'none',
            'Sec-Fetch-User': '?1',
            'upgrade-insecure-requests': '1',
        }
        counter = 0
        while True:
            try:
                response = requests.get(url=url, headers=headers, timeout=20)
                # print(response.status_code)
                if response.status_code == 200:
                    # image_attachment = base64.b64encode(response.content)
                    image_attachment = response.content
                    break
                else: print(f'{response.status_code} found for downloading image')
            except: sleep(0.3)
            counter += 1
            if counter == 10: break
    except Exception as e: print(f'Exception in download_image: {str(e)}')
    finally: return image_attachment

def crop_downloaded_image(filename):
    try:
        im = Image.open(filename)
        width, height = im.size   # Get dimensions
        new_width = 1120
        new_height = 600
        if width > new_width and height > new_height:
            left = (width - new_width)/2
            top = (height - new_height)/2
            right = (width + new_width)/2
            bottom = (height + new_height)/2
            im = im.crop((left, top, right, bottom))
            im.save(filename)
        elif height > new_height:
            left = (width - new_width)/2
            top = (height - new_height)/2
            right = (width + new_width)/2
            bottom = (height + new_height)/2
            im = im.crop((left, top, right, bottom))
            im.save(filename)
    except Exception as e: print(f'Exception in crop_downloaded_image: {e}')

def saving_picture_in_excel(data: list):
    workbook = Workbook()
    worksheet = workbook.active

    worksheet.cell(row=1, column=1, value='Brand')
    worksheet.cell(row=1, column=2, value='Model Code')
    worksheet.cell(row=1, column=3, value='Lens Code')
    worksheet.cell(row=1, column=4, value='Color Frame')
    worksheet.cell(row=1, column=5, value='Color Lens')
    worksheet.cell(row=1, column=6, value='SKU')
    worksheet.cell(row=1, column=7, value='Wholesale Price')
    worksheet.cell(row=1, column=8, value='Listing Price')
    worksheet.cell(row=1, column=9, value="Image")

    for index, d in enumerate(data):
        new_index = index + 2

        worksheet.cell(row=new_index, column=1, value=d[0])
        worksheet.cell(row=new_index, column=2, value=d[1])
        worksheet.cell(row=new_index, column=3, value=d[2])
        worksheet.cell(row=new_index, column=4, value=d[3])
        worksheet.cell(row=new_index, column=5, value=d[4])
        worksheet.cell(row=new_index, column=6, value=d[5])
        worksheet.cell(row=new_index, column=7, value=d[6])
        worksheet.cell(row=new_index, column=8, value=d[7])

        image = f'Images/{d[-3]}.jpg'
        if os.path.exists(image):
            im = Image.open(image)
            width, height = im.size
            worksheet.row_dimensions[new_index].height = height
            worksheet.add_image(Imag(image), anchor='I'+str(new_index))
            # col_letter = get_column_letter(7)
            # worksheet.column_dimensions[col_letter].width = width

    workbook.save('Safilo Results.xlsx')

DEBUG = True
try:
    pathofpyfolder = os.path.realpath(sys.argv[0])
    # get path of Exe folder
    path = pathofpyfolder.replace(pathofpyfolder.split('\\')[-1], '')
    # download chromedriver.exe with same version and get its path
    # if os.path.exists('chromedriver.exe'): os.remove('chromedriver.exe')
    if os.path.exists('Safilo Results.xlsx'): os.remove('Safilo Results.xlsx')

    # chromedriver_autoinstaller.install(path)
    if '.exe' in pathofpyfolder.split('\\')[-1]: DEBUG = False
    
    f = open('Safilo start.json')
    json_data = json.loads(f.read())
    f.close()

    brands = json_data['brands']

    
    f = open('requirements/safilo.json')
    data = json.loads(f.read())
    f.close()

    store = Store()
    store.link = data['url']
    store.username = data['username']
    store.password = data['password']
    store.login_flag = True

    result_filename = 'requirements/Safilo Results.json'

    if not os.path.exists('Logs'): os.makedirs('Logs')

    log_files = glob.glob('Logs/*.txt')
    if len(log_files) > 5:
        oldest_file = min(log_files, key=os.path.getctime)
        os.remove(oldest_file)
        log_files = glob.glob('Logs/*.txt')

    scrape_time = datetime.now().strftime('%d-%m-%Y %H-%M-%S')
    logs_filename = f'Logs/Logs {scrape_time}.txt'
    
    Safilo_Scraper(DEBUG, result_filename, logs_filename).controller(store, brands)
    
    for filename in glob.glob('Images/*'): os.remove(filename)
    data = read_data_from_json_file(DEBUG, result_filename)
    os.remove(result_filename)

    saving_picture_in_excel(data)
except Exception as e:
    if DEBUG: print('Exception: '+str(e))
    else: pass
