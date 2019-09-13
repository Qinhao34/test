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
        time.sleep(1)
        htmlResult = driverChrome.page_source
        driverChrome.quit()


        soupMachine = BeautifulSoup(htmlResult, 'html5lib')


        # Acquire name of the machine
        temp = str(soupMachine.find_all('div', {"class": "productList"}))
        MachineName = str(re.findall(r'<p class="name".*</p>', temp)).replace('<p class="name">', '').replace('</p>', '')
        specName = []
        nameIterator = iter(str(MachineName))

        for i in nameIterator:
            specName.append(str(i))
            excelTable.write(row, 0, specName)


        # Acquire label data
        temp_Labels = soupMachine.find_all('p', {"class": re.compile('specOrd\d{1,2}')})
        specLabel = []
        legendIterator = iter(temp_Labels)

        for i in legendIterator:
            if '"display: none;"' not in str(i):
                strLabel = str(re.findall(r'>.*<', str(i), re.DOTALL)).replace("['>", "").replace("<']", "")
                specLabel.append(str(strLabel))

        # Acquire number data
        temp_numbers = soupMachine.find_all('p', {"class": re.compile('specValOrd\d{1,2}')})
        specList = []
        listIterator = iter(temp_numbers)

        for i in listIterator:
            strI = str(re.findall(r'>.*<', str(i), re.DOTALL)).replace('\\n', "").replace('\\t', "").replace('\\r', "").replace(' ', "").replace("'>", "").replace("<'", "").replace('\\xa0', "")
            if str(strI) != "[]":
                specList.append(str(strI)[1:][:-1])

        print(len(specLabel))
        print(len(specList))

        # Export data to Excel
        # colNumber = 0
        # for i in range(len(specList)):
        #     print(i)
        #     excelTable.write(row, 2*colNumber+1, specLabel[i])
        #     excelTable.write(row, 2*colNumber+2, specList[i])
        #     colNumber += 1
        #

        row += 1

    excelFile.save('Doosan_MachineData.xls')


urlList = ["https://www.doosanmachinetools.com/en/product/series/D201_101/view.do"]
exportExcel123(urlList)

