from time import sleep
import os
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Application import Application


class ItDashboard:
    agencies = []

    def __init__(self, url):
        self.browser = Selenium()
        self.app = Application()
        self.browser.open_available_browser(url)
        self.browser.wait_until_page_contains_element('//*[@id="node-23"]/div/div/div/div/div/div/div/a')
        self.browser.find_element('//*[@id="node-23"]/div/div/div/div/div/div/div/a').click()
        self.agencies = self.browser.find_elements\
            ('//div[@id="agency-tiles-widget"]//div[@class="col-sm-4 text-center noUnderline"]')
        self.browser.set_download_directory(os.path.join(os.getcwd(), "output/"))
        sleep(10)

    def agencies_excel(self):
        app = Application()

        app.open_application()
        app.open_workbook('workbook.xlsx')
        app.set_active_worksheet(sheetname='new stuff')
        app.write_to_cells(row=1, column=1, value='new data')
        app.save_excel()
        app.quit_application()
        # self.app.open_application()
        # self.app.open_workbook('Agencies.xlsx')
        # self.app.set_active_worksheet(sheetname='Agencies')
        # for i, item in enumerate(self.agencies):
        #         agency_data = item.text.split("\n")
        #         self.app.write_to_cells(row=i, column=i, value=agency_data[0])
        #         self.app.write_to_cells(row=i, column=i+1, value=agency_data[2])
        # self.app.save_excel()
        # self.app.quit_application()


if __name__ == "__main__":
    it_dashboard = ItDashboard("https://itdashboard.gov/")
    it_dashboard.agencies_excel()
