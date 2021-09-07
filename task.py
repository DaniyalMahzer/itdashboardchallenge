import os
from time import sleep
from datetime import timedelta
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files


class ItDashboard:
    agencies = []
    uii_ids = []
    uii_links = []
    investment_table_data = []

    def __init__(self):
        self.browser = Selenium()
        self.files = Files()
        self.browser.open_available_browser("https://itdashboard.gov/")
        self.browser.set_download_directory(os.path.join(os.getcwd(), "output/"))

    def get_agencies(self):
        self.browser.wait_until_page_contains_element('//*[@id="node-23"]',)
        self.browser.find_element('//*[@id="node-23"]').click()
        self.agencies = self.browser.find_elements(
            '//div[@id="agency-tiles-widget"]//div[@class="col-sm-4 text-center noUnderline"]')

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

    def scrap_agency(self, agency_open):
        agency = self.agencies[agency_open]
        self.browser.wait_until_page_contains_element(agency)
        self.browser.find_element(agency).click()
        self.browser.wait_until_page_contains_element('//*[@id="investments-table-object_info"]',
                                                      timeout=timedelta(seconds=50))
        raw_total = self.browser.find_element('//*[@id="investments-table-object_info"]')
        data = raw_total.text.split(" ")
        total_entries = int(data[-2])
        self.browser.find_element('//*[@id="investments-table-object_length"]/label/select').click()
        self.browser.find_element('//*[@id="investments-table-object_length"]/label/select/option[4]').click()
        self.browser.wait_until_page_contains_element(
            f'//*[@id="investments-table-object"]/tbody/tr[{total_entries}]/td[1]', timeout=timedelta(seconds=30))
        for i in range(1, total_entries + 1):
            item = self.browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[1]')
            link = self.browser.find_element(
                f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[1]').find_element_by_tag_name(
                "a").get_attribute("href")
            self.uii_ids.append(item.text)
            self.uii_links.append(link)
            data = {"uii": self.uii_links, "links": self.uii_links}
        wb = self.files.create_workbook("output/uii.xlsx")
        wb.append_worksheet("Sheet", data)
        wb.save()

    def make_agency_excel(self):
        self.get_agencies()
        self.write_agencies()


if __name__ == "__main__":
    obj = ItDashboard()
    obj.make_agency_excel()
    obj.scrap_agency(0)
