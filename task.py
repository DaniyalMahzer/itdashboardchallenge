from time import sleep

from RPA.Browser.Selenium import Selenium


class ItDashboard:

    def __init__(self, url):
        self.browser = Selenium()
        self.browser.open_available_browser(url)
        self.browser.wait_until_page_contains_element('//*[@id="node-23"]/div/div/div/div/div/div/div/a')
        sleep(10)
        self.browser.find_element('//*[@id="node-23"]/div/div/div/div/div/div/div/a').click()
        


if __name__ == "__main__":
    it_dashboard = ItDashboard("https://itdashboard.gov/")

