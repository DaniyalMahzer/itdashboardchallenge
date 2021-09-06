from time import sleep
import os
from RPA.Browser.Selenium import Selenium


class ItDashboard:
    agencies = []

    def __init__(self, url):
        self.browser = Selenium()
        self.browser.open_available_browser(url)
        self.browser.wait_until_page_contains_element('//*[@id="node-23"]/div/div/div/div/div/div/div/a')
        self.browser.find_element('//*[@id="node-23"]/div/div/div/div/div/div/div/a').click()
        self.agencies = self.browser.find_elements\
            ('//div[@id="agency-tiles-widget"]//div[@class="col-sm-4 text-center noUnderline"]')
        self.browser.set_download_directory(os.path.join(os.getcwd(), "output/"))
        sleep(10)
        print(self.agencies)
        for item in self.agencies:
            data = item.text.split("\n")
            print(data[0], data[2])


if __name__ == "__main__":
    it_dashboard = ItDashboard("https://itdashboard.gov/")
