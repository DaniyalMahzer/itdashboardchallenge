import os
from time import sleep

from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files


class ItDashboard:
    agencies = []

    def __init__(self):
        self.browser = Selenium()
        self.files = Files()
        self.browser.open_available_browser("https://itdashboard.gov/")
        self.browser.set_download_directory(os.path.join(os.getcwd(), "output/"))

    def get_agencies(self):
        self.browser.wait_until_page_contains_element('//*[@id="node-23"]/div/div/div/div/div/div/div/a')
        self.browser.find_element('//*[@id="node-23"]/div/div/div/div/div/div/div/a').click()
        self.agencies = self.browser.find_elements\
            ('//div[@id="agency-tiles-widget"]//div[@class="col-sm-4 text-center noUnderline"]')

    def write_agencies(self):
        companies = ['companies', ]
        investments = ['investments', ]
        for item in self.agencies:
            agency_data = item.text.split('\n')
            companies.append(agency_data[0])
            investments.append(agency_data[2])
        entries = {"companies": companies, "investments": investments}
        wb = self.files.create_workbook("output/Agencies.xlsx")
        wb.append_worksheet("Sheet", entries)
        wb.save()

    def make_agency_excel(self):
        self.get_agencies()
        self.write_agencies()

    def scrap_agency(self, agency):
        agency = self.agencies[agency]
        self.browser.wait_until_page_contains_element(agency)
        self.browser.find_element(agency).click()
        sleep(10)


if __name__ == "__main__":
    obj = ItDashboard()
    obj.make_agency_excel()
    obj.scrap_agency(0)
