import os
from time import sleep
from datetime import timedelta
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files


class ItDashboard:
    agencies = []
    headers = []
    def __init__(self):
        self.browser = Selenium()
        self.files = Files()
        self.browser.open_available_browser("https://itdashboard.gov/")

    def get_agencies(self):
        self.browser.wait_until_page_contains_element('//*[@id="node-23"]')
        self.browser.find_element('//*[@id="node-23"]').click()
        self.browser.wait_until_page_contains_element('//div[@id="agency-tiles-widget"]', timeout=timedelta(seconds=10))
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

    def get_headers(self):
        for i in range(1, 7):
            head = self.browser.find_element(
                f'//*[@id="investments-table-object_wrapper"]/div[3]/div[1]/div/table/thead/tr[2]/th[{i}]')
            self.headers.append(head)

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
        self.get_headers()
        uii_ids = [self.headers[0], ]
        bureau = [self.headers[1], ]
        investment_title = [self.headers[2], ]
        total_FY2021 = [self.headers[3], ]
        type_agency = [self.headers[4], ]
        CIO_rating = [self.headers[5], ]
        num_of_project =[self.headers[6], ]
        for i in range(1, total_entries + 1):
            item = self.browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[1]')
            try:
                bureau_current = self.browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[2]').text
                investment_title_current = self.browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[3]').text
                total_FY2021_current = self.browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[4]').text
                type_agency_current = self.browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[5]').text
                CIO_rating_current = self.browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[6]').text
                num_of_project_current = self.browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[7]').text
            except:
                bureau_current = ''
                investment_title_current = ''
                total_FY2021_current = ''
                type_agency_current = ''
                CIO_rating_current = ''
                num_of_project_current = ''
            try:
                link = self.browser.find_element(
                    f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[1]').find_element_by_tag_name(
                    "a").get_attribute("href")
            except:
                link = ''
            if link:
                downloader = Selenium()
                downloader.open_available_browser(link)
                downloader.find_element('//div[@id="business-case-pdf"]').click()
                downloader.set_download_directory("outputs/")
                while True:
                    try:
                        sleep(2)
                        if downloader.find_element('//div[@id="business-case-pdf"]').find_element_by_tag_name("span"):
                            sleep(1)
                        else:
                            break
                    except:
                        if downloader.find_element('//*[contains(@id,"business-case-pdf")]//a[@aria-busy="false"]'):
                            sleep(1)
                            break
                sleep(2)
                downloader.close_browser()
            bureau.append(bureau_current)
            investment_title.append(investment_title_current)
            total_FY2021.append(total_FY2021_current)
            type_agency.append(type_agency_current)
            CIO_rating.append(CIO_rating_current)
            num_of_project.append(num_of_project_current)
            uii_ids.append(item.text)
        data = {"uii": uii_ids,
                "bureau": bureau,
                "company": investment_title,
                "FY2021": total_FY2021,
                "agency_type": type_agency,
                "CIO rating": CIO_rating,
                "# of project": num_of_project,
                }
        wb = self.files.create_workbook("output/uii.xlsx")
        wb.append_worksheet("Sheet", data)
        wb.save()

    def make_agency_excel(self):
        self.get_agencies()
        self.write_agencies()


if __name__ == "__main__":
    obj = ItDashboard()
    obj.make_agency_excel()
    obj.scrap_agency(-3)
