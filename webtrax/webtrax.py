import os, shutil
from time import sleep
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
from webtrax import constants as const


class Webtrax(webdriver.Edge):
    """webtrax class"""
    def __init__(self, driver_path = const.DRIVER_PATH , teardown = False):
        """init method"""
        self.driver_path = driver_path
        self.teardown = teardown
        options = webdriver.EdgeOptions()
        options.add_experimental_option('detach', True)
        # options.add_experimental_option('--headless', True)
        super(Webtrax, self).__init__(options=options)
        self.maximize_window()
        self.implicitly_wait(15)

    def __exit__(self, exc_type, exc, traceback):
        """exit method"""
        if self.teardown:
            self.quit()

    def land_page(self):
        """land to the escalation webpage"""
        self.get(const.BASE_URL)
        self.find_element(By.ID, "ctl00_ContentPlaceHolder1_Login1_UserName").send_keys(const.WEBTRAX_USERNAME)
        self.find_element(By.ID, "ctl00_ContentPlaceHolder1_Login1_Password").send_keys(const.WEBTRAX_PASSWORD)
        self.find_element(By.ID, "ctl00_ContentPlaceHolder1_Login1_LoginButton").click()
        self.get(const.ESCALATION_URL)

    def navigate_to_page(self, url):
        """navigate"""
        self.get(url)

    def get_datatable(self):
        """hello"""
        data_list = []
        header_list = ['Age Group', 'Age', 'NoClaims', 'ID', 'Name', 'Description', 'Submitter_Name', 'PlatDesc', 'Submitted Date', 'Due_Date', 'Assigned To', 'Opened By', 'DIV', 'ICN']
        self.find_element(By.ID, 'ctl00_ContentPlaceHolder1_ccHold').click()
        self.find_element(By.ID, 'ctl00_ContentPlaceHolder1_CCInclude').click()
        table = self.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_GridView1"]/tbody')
        # for row in table.find_elements(By.XPATH, ".//tr"):
        #     print(row.text)
        #     for td in row.find_elements(By.XPATH, ".//td"):
        #         print(td.text)
        for row in table.find_elements(By.XPATH, ".//tr"):
            data_list.append([td.text for td in row.find_elements(By.XPATH, ".//td") if not td.text.startswith('No')])
            data_list = list(filter(lambda x: x != [], data_list))
        return pd.DataFrame(data_list, columns=header_list)

    def get_links(self):
        "Loop over tags"
        link_list = []
        self.find_element(By.ID, 'ctl00_ContentPlaceHolder1_ccHold').click()
        self.find_element(By.ID, 'ctl00_ContentPlaceHolder1_CCInclude').click()
        sleep(15)
        url = r"https://optumhealthopscontrol.uhc.com/Escalation/User/Queue-Details.aspx?DeptID=1&Type="
        queues = self.find_elements(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_GridView2"]/tbody/tr/td[2]')
        locations = self.find_elements(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_GridView2"]/tbody/tr/td[3]')
        for i in range(len(queues)):
            temp_data = url + locations[i].text + "&Que=" + queues[i].text
            link_list.append(temp_data)
        return link_list

    @classmethod
    def get_file_path(cls, file_name):
        """get the file path"""
        folder_path = os.path.join(os.path.expanduser(r'~\Documents'), 'Templates')
        source_path = os.path.join(r"\\nas00913pn\hbs\optum_team_reports\01_Team\03_UBH Unet Team\Sameer_Khan\Templates", file_name)
        destination_path = os.path.join(folder_path, file_name)
        if not os.path.isdir(folder_path):
            os.makedirs(folder_path)
        if not os.path.exists(destination_path):
            shutil.copy(source_path, destination_path)
        return destination_path

    def get_file_path2(self):
        """get path of file"""
        return str(Path.joinpath(Path().absolute(), 'Webtrax Escalation Inventory.xlsm'))
        
        