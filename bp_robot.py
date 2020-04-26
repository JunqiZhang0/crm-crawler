from datetime import date
from xlrd import open_workbook
from selenium import webdriver, common
from selenium.webdriver.chrome.options import Options
import time
from bs4 import BeautifulSoup as bs
import datetime
import xlwt
from tqdm import tqdm


class bp_robot():

    def __init__(self,com_len):
        self.driver = webdriver.Chrome()
        self.search_len = com_len
        self.u = input("please input your account: \n")
        self.p = input("please input your password: \n")
        self.sleep_time = 3
        self.company_name_index = 1
        self.url_box = {}
        self.detail_box = {}
        self.benifit_point_url = input("please input your benefit-point login url: \n")
        self.search_url= input("please enter the url when you do search: \n(this can be got if you do a search with a word that doesn't exist, you copy the url without the searched word)")

    def login(self):

        self.driver.get(self.benifit_point_url)
        time.sleep(0.5)
        self.driver.find_element_by_name('username').send_keys(self.u)
        self.driver.find_element_by_name("password").send_keys(self.p)
        time.sleep(0.5)
        self.driver.find_element_by_xpath("//input[@value='Login']").click()


    def search(self, name):
        url_search_head = self.search_url
        if name is not None:
            name = name.replace(',', '')
            text = name.split()
        else:
            text = "test".split()
        
        if len(text) >= self.search_len:
            text = text[0:self.search_len]

        appendix = ''
        for t in text:
            appendix = appendix+t+"+"
        appendix = appendix[0:-1]
        search = url_search_head + appendix

        time.sleep(5)
        self.driver.get(search)
        time.sleep(5)
        page_source = self.driver.page_source
        soup = bs(page_source, 'html.parser')
        self.detail_box[name] = {}
        if not str(soup.prettify()).__contains__("There are no Account Search Results."):
            # street1 = str(soup.findAll("div",{"id":"addres_street1"})[0].string)
            # print("found")
            if len(soup.findAll("div",{"id":"addres_street1"})) != 0:
                street1 = str(soup.findAll("div",{"id":"addres_street1"})[0].string)
            else:
                street1 = ""

            if len(soup.findAll("div",{"id":"addres_street2"})) != 0:
                street2 = str(soup.findAll("div",{"id":"addres_street2"})[0].string)
            else:
                street2 = ""
            if len(soup.findAll("div",{"id":"address_city"})) is not 0:
                city_state_zip = str(soup.findAll("div",{"id":"address_city"})[0].string).replace(",","").replace("-","").split()
            else:
                city_state_zip = ""

            if len(soup.findAll("td",{"id":"saleslead_value"})) is not 0:
                producer = str(soup.findAll("td", {"id":"saleslead_value"})[0].string).replace(","," ").split()
            else:
                producer = ""
            producer_name = ""
            for item in producer:
                producer_name = " " + item + producer_name
            producer_name = producer_name[1:]
            self.detail_box[name]['address'] = street1+""+street2
            if len(city_state_zip) is not 0:
                self.detail_box[name]['city'] = city_state_zip[0]
                self.detail_box[name]['state'] = city_state_zip[1]
                self.detail_box[name]['zip_code'] = city_state_zip[2]
            else:
                self.detail_box[name]['city'] = ""
                self.detail_box[name]['state'] = ""
                self.detail_box[name]['zip_code'] = ""
            self.detail_box[name]["producer"] = producer_name

        



if __name__ == "__main__":
    start = datetime.datetime.now()



    # Below is the information could be changed
    path = input("please input you file path: \n")
    company_name_index = int(input("please input the colum number of comany name\n"))-1
    zip_code_index = int(input("please input the colum number of zip_code\n"))-1
    city_index = int(input("please input the colum number of city\n"))-1
    state_index = int(input("please input the colum number of state\n"))-1
    country_index = int(input("please input the colum number of country\n"))-1
    street_index = int(input("please input the colum number of street\n"))-1
    crm_index = int(input("please input the colum number of databse category\n"))-1
    producer_index =  int(input("please input the colum number of producer name\n"))-1
    GET_INDEX = int(input("please input the colum number of get\n"))-1
    com_len = int(input("How many words do you want to use to search the comanpy?\n"))
    sheet_name = input("please input the name the sheet you want to scrub: \n")

    start_row = 1



    robot = bp_robot(com_len)
    robot.login()

    from xlutils.copy import copy
    with open_workbook(path) as workbook:
        new_workbook = copy(workbook)
        new_worksheet = new_workbook.get_sheet(0)
        worksheet = workbook.sheet_by_name(sheet_name)
        sheet_len = worksheet.nrows
        for row in tqdm(range(start_row,sheet_len)):
            c = worksheet.row(row)[int(crm_index)].value.lower() == "bp" or worksheet.row(row)[int(crm_index)].value.lower() == ""
            if c:
                # this line extracted the company name
                company_name = worksheet.row(row)[int(company_name_index)].value
                company_name = company_name.replace(",","")
                company_name = company_name.replace("INC","")
                company_name = company_name.replace(".","")
                company_name = company_name.replace("LLC","")
                robot.search(company_name)
                # This line should record all the information in the two boxes


                if len(robot.detail_box[company_name]) is not 0:
                    all_apply = False

                    com = robot.detail_box[company_name]
                    
                        # company.items() contains
                    coming_data = set()
                    for company in com.items():
                        company = company[1]
                        coming_data.add(company)
                    

                    table_data = set()
                    table_data.add(worksheet.row(row)[int(street_index)].value)
                    table_data.add(worksheet.row(row)[int(zip_code_index)].value)
                    table_data.add(worksheet.row(row)[int(city_index)].value)
                    table_data.add(worksheet.row(row)[int(country_index)].value)
                    table_data.add(worksheet.row(row)[int(state_index)].value)
                    c = coming_data.intersection(table_data)

                    if len(c) is not 0:
                        all_apply = True

                    if all_apply:
                        new_worksheet.write(row,GET_INDEX,"bp_get")
                        new_worksheet.write(row, producer_index,robot.detail_box[company_name]['producer'])

                robot.url_box = {}
                robot.detail_box = {}




        robot.driver.close()
        new_workbook.save(path[0:-5] + "-after-scrubing-BP.xls")


    endtime = datetime.datetime.now()
    print("time: ",end="")
    print((endtime.minute-start.minute),end=" mins")
