from time import sleep
import os
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files


class ItDashboard:
    agencies_data = {}
    def __init__(self, url):
        self.browser = Selenium()
        self.files = Files()
        self.browser.open_available_browser(url)
        self.browser.wait_until_page_contains_element('//*[@id="node-23"]/div/div/div/div/div/div/div/a')
        self.browser.find_element('//*[@id="node-23"]/div/div/div/div/div/div/div/a').click()
        self.agencies = self.browser.find_elements\
            ('//div[@id="agency-tiles-widget"]//div[@class="col-sm-4 text-center noUnderline"]')
        self.browser.set_download_directory(os.path.join(os.getcwd(), "output/"))
        sleep(10)

    def get_agencies(self):
        for item in self.agencies:
            data = item.text.split
            wb = self.files.create_workbook("output/Agencies.xlsx")
            wb.append_worksheet("Sheet", [data[0], data[2]])
            wb.save()


if __name__ == "__main__":
    it_dashboard = ItDashboard("https://itdashboard.gov/")
    it_dashboard.get_agencies()
