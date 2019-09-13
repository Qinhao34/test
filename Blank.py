from bs4 import BeautifulSoup
from xlwt import *
import re
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options


def exportExcel123(urlList):

    excelFile = Workbook(encoding='utf-8')
    excelTable = excelFile.add_sheet("Doosan")

    optionChrome = Options()
    optionChrome.add_argument('--headless')
    optionChrome.add_argument('--disable-gpu')
    optionChrome.add_argument('disable-plugins')
    optionChrome.add_argument('disable-extensions')

    urlIterator = iter(urlList)
    row = 0

    for i in urlIterator:
        driverChrome = webdriver.Chrome(options = optionChrome)
        driverChrome.get(i)
        time.sleep(2)
        htmlResult = driverChrome.page_source
        driverChrome.quit()


        soupMachine = BeautifulSoup(htmlResult, 'html5lib')


        # Acquire name of the machine
        temp = str(soupMachine.find_all('div', {"class": "productList"}))
        MachineName = str(re.findall(r'<p class="name".*</p>', temp)).replace('<p class="name">', '').replace('</p>', '').split(',')
        # specName = []
        # nameIterator = iter(MachineName)
        #print(MachineName)
        for i in range(len(MachineName)):
             #specName.append(str(i))
            if i != len(MachineName)-1:
                Name = MachineName[i]
                excelTable.write(row, 0, Name[2:][:-1])
                row += 1
            else:
                Name = MachineName[i]
                excelTable.write(row, 0, Name[2:][:-2])
                row += 1

        # Acquire label data
        # temp_Labels = soupMachine.find_all('p', {"class": re.compile('specOrd\d{1,2}')})
        # print(temp_Labels)
        # specLabel = []
        # legendIterator = iter(temp_Labels)

        # for i in legendIterator:
        #     if '"display: none;"' not in str(i):
        #         strLabel = str(re.findall(r'>.*<', str(i), re.DOTALL)).replace("['>", "").replace("<']", "")
        #         specLabel.append(str(strLabel))

        # Acquire number data
        temp_numbers = soupMachine.find_all('p', {"class": re.compile('specValOrd\d{1,2}')})
        print(temp_numbers)
        # specList = []
        # listIterator = iter(temp_numbers)
        #
        # for i in listIterator:
        #     strI = str(re.findall(r'>.*<', str(i), re.DOTALL)).replace('\\n', "").replace('\\t', "").replace('\\r',
        #                                                                                                      "").replace(
        #         ' ', "").replace("'>", "").replace("<'", "").replace('\\xa0', "")
        #     if str(strI) != "[]":
        #         specList.append(str(strI)[1:][:-1])







    excelFile.save('Doosan_MachineData.xls')


urlList = ["https://www.doosanmachinetools.com/en/product/series/D201_48/view.do"]
exportExcel123(urlList)
