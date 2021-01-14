from selenium import webdriver
from selenium import webdriver
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from bs4 import BeautifulSoup
import xlsxwriter
from openpyxl import Workbook
import lxml
import html5lib
import pandas as pd



def attach_to_session(executor_url, session_id):
    original_execute = WebDriver.execute
    def new_command_execute(self, command, params=None):
        if command == "newSession":

            return {'success': 0, 'value': None, 'sessionId': session_id}
        else:
            return original_execute(self, command, params)

    WebDriver.execute = new_command_execute
    driver = webdriver.Remote(command_executor=executor_url, desired_capabilities={})
    driver.session_id = session_id

    WebDriver.execute = original_execute
    return driver


class PopboxData:
    def __init__(self,projectNumber,ExpShipDate,ProductMix,ProjectCreator,MatDesc,LocalMatDesc,SalesOrg,ReplacingSKU,BaseSKU,UOM,EANStrategy,Type,PPG,AsVendor):
        self.ProjectNumber=projectNumber
        self.ExpShipDate= ExpShipDate
        self.ProductMix= ProductMix
        self.ProjectCreator= ProjectCreator
        self.MatDesc= MatDesc
        self.LocalMatDesc= LocalMatDesc
        self.SalesOrg= SalesOrg
        self.ReplacingSKU = ReplacingSKU
        self.BaseSKU= BaseSKU
        self.UOM= UOM
        self.EANStrategy = EANStrategy
        self.Type= Type
        self.PPG= PPG
        self.AsVendor= AsVendor

    def getProjectDescription(self):
        return str(self.MatDesc+ '\n' + "ExpectedShipDate:" + self.ExpShipDate.strip() + '\n'+"ProductMix:" + self.ProductMix.strip() + '\n'+"BaseSKU:" + self.BaseSKU.strip() + '\n'+ "ProjectCreator:" + self.ProjectCreator.strip())


    def __str__(self):
        return '%s; %s; %s; %s; %s; %s; %s; %s; %s; %s; %s; %s; %s; %s' % (self.ProjectNumber,self.ExpShipDate,self.ProductMix,
                                                                       self.ProjectCreator,self.MatDesc,
                                                                       self.LocalMatDesc,self.SalesOrg,self.ReplacingSKU,
                                                                       self.BaseSKU,self.UOM,self.EANStrategy,self.Type,self.PPG,
                                                                       self.AsVendor)

class PopboxScrapping:

    def __init__(self, page_source):
        self.data= []
        self.page_source = page_source.replace("<br>", ' ')
        self.details_html = BeautifulSoup(self.page_source, 'html.parser')

    def getPopboxData(self):

        allValues = self.details_html.find_all('label', class_="control-value")
        projectNumber = self.getProjectNumber()
        productMix = self.getProductMix()
        expShipDate = self.getExpShipDate()
        projectCreator = self.getProjectCreator()
        matDesc = self.getMatDesc()
        localMatDesc = self.getLocalMatDesc()
        salesOrg = self.getSalesOrg()

        replacingSKU = self.getReplacingSKU()
        baseSKU = self.getBaseSKU()
        uom = self.getUOM()
        eanstrategy= self.getEANStrategy()
        type = self.getType()
        ppg = self.getPPG()
        plant = self.getAsVendor()

        projectData = PopboxData(projectNumber,expShipDate, productMix,projectCreator,
                                 matDesc,localMatDesc,salesOrg,replacingSKU,
                                 baseSKU, uom, eanstrategy,type,ppg, plant)

        print(projectData)
        self.data.append(projectData)
        return projectData



    def getValue(self, entry):
        tags = self.details_html.find_all(lambda tag: tag.name == "label" and entry == str(tag.text).strip() and 'col-3' in tag.get('class'))
        value = ''
        for m in tags:
            divs = m.find_next_siblings(lambda tag: tag.name == "div" and 'col-9' in tag.get('class') and tag.find('span'))
            if (len(divs) > 0):
                value = divs[0].span.text
        return value

    def getProjectNumber(self):
        number = self.details_html.find_all('li', class_="breadcrumb-item")
        str = number[1].a.text
        projectNumber = str[-8:]
        return projectNumber
    def getExpShipDate(self):
        return self.getValue('Expected Ship Date')
    def getProductMix(self):
        return self.getValue('Product Mix / SKU inside')
    def getProjectCreator(self):
        return self.getValue('Project Creator (COPACK)')
    def getMatDesc(self):
        return self.getValue('SKU Material Description (EN)')
    def getLocalMatDesc(self):
        return self.getValue('SKU Local Material Description')
    def getSalesOrg(self):
        return self.getValue('Sales Org')
    def getReplacingSKU(self):
        return  self.getValue('Replacing SKU')
    def getBaseSKU(self):
        return self.getValue('Base SKU Nr')
    def getUOM(self):
        return self.getValue('UOM')
    def getEANStrategy(self):
        return  self.getValue('EAN Strategy')
    def getType(self):
        return self.getValue('MDM SKU Type')
    def getPPG(self):
        return self.getValue('PPG to be used')
    def getAsVendor(self):
        return self.getValue('Assembly Vendor')

class ExportController(PopboxScrapping):

    def exportToXlsx(self):
        tableTitle = [
            "ProjectNumber",
            "ExpShipDate",
            "ProductMix",
            "ProjectCreator",
            "MatDesc",
            "LocalMatDesc",
            "SalesOrg",
            "ReplacingSku",
            "BaseSKU",
            "UOM",
            "EANStrategy",
            "Type",
            "PPG",
            "Plant"
            ]
        workbook = xlsxwriter.Workbook('ProjectData.xlsx')
        worksheet = workbook.add_worksheet()
        tableFormat = workbook.add_format({'border': 1})
        headerFormat = workbook.add_format({'bold': True, 'bg_color': 'yellow', 'border': 1})
        worksheet.set_column(0, 0, 100)
        worksheet.set_column(1, 1, 20)
        worksheet.set_column(2, 2, 30)
        row = 0
        column = 0
        while (column < len(tableTitle)):
            worksheet.write(row, column, tableTitle[column], headerFormat)
            column += 1
        row = 1
        while (row - 1 < len(self.data)):
            column = 0
            while (column < len(tableTitle)):
                worksheet.write(row, column, str(self.data[row - 1]).split("; ")[column], tableFormat)
                column += 1
            row += 1
        workbook.close()

def run():
    popboxSessionFile = open("sessionIdAndProduct.txt", "r")
    sessionAndExecutorUrl = popboxSessionFile.read().splitlines()
    session_id = sessionAndExecutorUrl[0]
    executor_url = sessionAndExecutorUrl[1]
    print(session_id)
    print(executor_url)
    popboxProductUrl = sessionAndExecutorUrl[2]

    driver = attach_to_session(executor_url=executor_url, session_id=session_id)
    driver.get(popboxProductUrl)

    cs = ExportController(driver.page_source)

    projectData = cs.getPopboxData()

    cs.exportToXlsx()
    return projectData



















