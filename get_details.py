from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, \
    ElementNotVisibleException, ElementNotSelectableException
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
import json
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import xlsxwriter


class Amazon:
    def __init__(self):
        chrome_options = Options()
        chrome_options.add_argument("--disable-extensions")
        # chrome_options.add_argument('--headless')
        # chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--disable-setuid-sandbox')
        chrome_obj = ChromeDriverManager()
        self.driver = webdriver.Chrome(executable_path=chrome_obj.install(),
                                       chrome_options=chrome_options)

        self.main_window = self.driver.current_window_handle
        self.order_details = {}

        self.wait = WebDriverWait(self.driver, 90, poll_frequency=0.1,
                                  ignored_exceptions=[NoSuchElementException, ElementNotVisibleException,
                                                      ElementNotSelectableException])

    def get_all_order_details(self):
        order_table = self.driver.find_element_by_id("orders-table")
        order_table_body = order_table.find_element_by_tag_name("tbody")
        order_table_trs = order_table_body.find_elements_by_tag_name("tr")
        tr: WebElement
        current_order = {}
        current_order_track = ""
        for tr in order_table_trs:
            tds = tr.find_elements_by_tag_name("td")

            order_date = tds[1].text
            print(order_date)
            if "ASIN" not in order_date:
                order_details = tds[2].find_elements_by_tag_name("a")
                order_id = order_details[0].text
                order_name = order_details[1].text
                print("Executed")
                print(order_id)

                product_name = tr.find_element_by_class_name("myo-list-orders-product-name-cell").text
                current_order_track = order_id
                current_order[order_id] = {}
                current_order[order_id]["product_order"] = [product_name]
                current_order[order_id]["name"] = order_name
            else:
                product_order = tr.find_element_by_class_name("myo-list-orders-product-name-cell").text
                current_order[current_order_track]["product_order"].append(product_order)

        return current_order

    def fetch_all_information(self):
        self.order_details = self.get_all_order_details()
        for order_id in self.order_details:
            try:
                current_customer_info: dict = self.order_details[order_id]
                order_link = self.driver.find_element_by_link_text(order_id)
                # open a link in new tab
                order_link.location_once_scrolled_into_view
                # order_link.send_keys(Keys.COMMAND + 't')
                from selenium.webdriver import ActionChains
                actions = ActionChains(self.driver)
                about = self.driver.find_element_by_link_text(order_id)
                about = self.driver.find_element_by_link_text(order_id)
                actions.key_down(Keys.CONTROL).click(about).key_up(Keys.CONTROL).perform()

                self.wait.until(EC.number_of_windows_to_be(2))

                individual_information: dict = self.fetch_individual_information()
                current_customer_info.update(individual_information)
            except Exception as e:
                print(e)
                if len(self.driver.window_handles) == 2:
                    self.driver.switch_to.window(self=self.main_window)

        with open("demo.json", mode="w") as fd:
            json.dump(self.order_details, fd, indent=3)

        self.write_to_excel(order_dicts=self.order_details)

    def fetch_individual_information(self):

        self.driver.switch_to.window(window_name=self.driver.window_handles[1])
        result = {}
        try:
            ele = self.wait.until(EC.presence_of_element_located((By.XPATH,
                                                                  "//div[@data-test-id"
                                                                  "='shipping-section-buyer-address']")))

            self.wait.until(EC.visibility_of(ele))

            address_element = self.driver.find_element_by_xpath("//div[@data-test-id='shipping-section-buyer-address']")
            address_text = address_element.text

            # shipping-section-phone
            phone_element = self.driver.find_element_by_xpath("//span[@data-test-id='shipping-section-phone']")
            phone_element = phone_element.text

            self.driver.close()
            self.driver.switch_to.window(window_name=self.main_window)
            return {"phone": phone_element, "address": address_text}
        except Exception as e:
            self.driver.switch_to.window(window_name=self.main_window)
            return {"phone": "", "address": ""}

    @staticmethod
    def write_to_excel(order_dicts: dict):
        """It help to write the file in excel sheet

        :param order_dicts:
        :return:
        """
        products = ["Multani", "Orange", "Sandalwood"]

        # Generate Excel sheet with different time zone
        current_date = datetime.now()
        workbook = xlsxwriter.Workbook(f"customer_list_{current_date.strftime('%Y%b%d_%H%M%S')}.xlsx")
        worksheet = workbook.add_worksheet()

        worksheet.write('A1', 'Order ID')
        worksheet.write('B1', 'Name')
        worksheet.write('C1', 'Phone Number')
        worksheet.write('D1', 'Address')

        for index, value in enumerate(products):
            worksheet.write(0, 4 + index, value)

        for index, order_key in enumerate(order_dicts):
            worksheet.write(index + 1, 0, order_key)
            worksheet.write(index + 1, 1, order_dicts[order_key].get("name", "name_not_found"))
            worksheet.write(index + 1, 2, order_dicts[order_key]["phone"])
            worksheet.write(index + 1, 3, order_dicts[order_key]["address"])

            # Write the product quantity
            for or_pd in order_dicts[order_key]["product_order"]:
                for pd_index, pd in enumerate(products):
                    if pd in or_pd:
                        worksheet.write(index + 1, 4 + pd_index, "yes")

        workbook.close()


if __name__ == '__main__':
    a = Amazon()
    input("Please login the screen and go to manage order ")
    # a.get_all_order_details()
    a.fetch_all_information()

    for iteration in range(3):
        second_time = str(input("Do you wish to execute script for next page , (yes/no)"))

        if second_time.lower() == "yes":
            # Empty the first page list
            a.order_details = {}
            a.fetch_all_information()
        else:
            break







    print("This is break point")
