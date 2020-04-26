from datetime import date
from xlrd import open_workbook
from selenium import webdriver, common
from selenium.webdriver.chrome.options import Options
import time
from bs4 import BeautifulSoup as bs
import datetime
import xlwt
from tqdm import tqdm


class crm_robot():

    def __init__(self,com_len):
        self.driver = webdriver.Chrome()
        self.search_len = com_len
        self.built_in_name = 'backyard'
        self.u = input("please input your account: \n")
        self.p = input("please input your password: \n")
        self.sleep_time = 3
        self.company_name_index = 1
        self.url_box = {}
        self.detail_box = {}
        self.account_prefix= input("please enter the url when you do search: \n(this can be got if you do a search with a word that doesn't exist, you copy the url without the searched word)")
        self.crm_login_url = input("please input your benefit-point login url: \n")


    def login(self):

        self.driver.get(self.crm_login_url)
        time.sleep(2)
        self.driver.find_element_by_id('i0116').send_keys(self.u)
        self.driver.find_element_by_id('idSIButton9').click()
        time.sleep(4)
        self.driver.find_element_by_id("passwordInput").send_keys(self.p)
        self.driver.find_element_by_id('submitButton').click()
        time.sleep(2)
        self.driver.find_element_by_id('idSIButton9').click()

    def search(self, name):
        url_search_head = self.account_prefix
        if name is not None:
            name = name.replace(',', '')
            text = name.split()
        else:
            text = self.built_in_name.split()
        
        if len(text) >= self.search_len:
            text = text[0:self.search_len]

        appendix = ''
        for t in text:
            appendix = appendix+t+" "
        appendix = appendix[0:-1]+'&searchType=1'
        search = url_search_head + appendix
        time.sleep(3)
        self.driver.get(search)
        time.sleep(3)
        page_source = self.driver.page_source
        soup = bs(page_source, 'html.parser')
        soup.prettify()
        company_names = soup.findAll(
            "ul", {"id": "MscrmControls.Grid.GridControl-account-MscrmControls.Grid.GridControl.account-GridList"})
         # this is the box containing all company_names and url of them
        if len(company_names) == 0:
            self.detail_box[name] = {}
        else:
            self.detail_box[name] = {}
            for item in company_names[0]:
                # print(item)
                if(not item.get('id').__contains__('contact')):
                    url = self.account_prefix + item['data-id'][-36:]
                    self.url_box[item["aria-label"][5:]] = url

    def account_detail(self, name):
        counter = 0
        for item in self.url_box.items():
            counter += 1
            self.driver.get(item[1])
            time.sleep(4)
            page_source = self.driver.page_source
            soup = bs(page_source, 'html.parser')


            zip_code = soup.findAll("input", {"aria-label": "ZIP/Postal Code"})
            self.detail_box[name][item[0]] = set()
            for zz in zip_code:
                self.detail_box[name][item[0]].add(zz['value'])


            city = soup.findAll("input", {"aria-label": "Address 1: City"})
            for zz in city:
                self.detail_box[name][item[0]].add(zz['value'])


            country = soup.findAll("input", {"aria-label": "Country"})
            for zz in country:
               self.detail_box[name][item[0]].add(zz['value'])

            state = soup.findAll("input", {"aria-label": "State/Province"})
            for zz in state:
               self.detail_box[name][item[0]].add(zz['value'])
            # print(str(counter/size*100)+"%")


if __name__ == "__main__":
    start = datetime.datetime.now()

    path = input("please input you file path: \n")
    # Below is the information could be changed
    company_name_index = int(input("please input the colum number of comany name\n"))-1
    zip_code_index = int(input("please input the colum number of zip_code\n"))-1
    com_len = int(input("How many words do you want to use to search the comanpy?\n"))
    city_index = int(input("please input the colum number of city\n"))-1
    state_index = int(input("please input the colum number of state\n"))-1
    country_index = int(input("please input the colum number of country\n"))-1
    crm_index = int(input("please input the colum number of databse category\n"))-1
    GET_INDEX = int(input("please input the colum number of get\n"))-1
    sheet_name = input("please input the name the sheet you want to scrub: \n")
    start_row = 1



    robot = crm_robot(com_len)
    robot.login()

    from xlutils.copy import copy
    with open_workbook(path) as workbook:
        new_workbook = copy(workbook)
        new_worksheet = new_workbook.get_sheet(0)
        worksheet = workbook.sheet_by_name(sheet_name)
        sheet_len = worksheet.nrows
        for row in tqdm(range(start_row,sheet_len)):
            c = worksheet.row(row)[int(crm_index)].value.lower() == "crm" or worksheet.row(row)[int(crm_index)].value.lower() == ""
            if c:
                # this line extracted the company name
                company_name = worksheet.row(row)[int(company_name_index)].value
                company_name = company_name.replace(",","")
                company_name = company_name.replace("INC","")
                company_name = company_name.replace(".","")
                company_name = company_name.replace("LLC","")
                robot.search(company_name)
                
                robot.account_detail(company_name)
                # This line should record all the information in the two boxes


                if len(robot.detail_box[company_name]) is not 0:
                    all_apply = False

                    com = robot.detail_box[company_name]
                    
                        # company.items() contains
                    for company in com.items():
                        company = company[1]
                        table_data = set()
                        table_data.add(worksheet.row(row)[int(zip_code_index)].value)
                        table_data.add(worksheet.row(row)[int(city_index)].value)
                        table_data.add(worksheet.row(row)[int(country_index)].value)
                        table_data.add(worksheet.row(row)[int(state_index)].value)
                        c = company.intersection(table_data)
                        # print(table_data)

                        if len(c) is not 0:
                            all_apply = True
                            continue

                    if all_apply:
                        # print(company_name)
                        new_worksheet.write(row,GET_INDEX,"crm-get")

                robot.url_box = {}
                robot.detail_box = {}


        robot.driver.close()
        new_workbook.save(path[0:-5] + "-after-scrubing-CORE.xls")


    endtime = datetime.datetime.now()
    print("time: ",end="")
    print((endtime.minute-start.minute),end=" mins")

        

