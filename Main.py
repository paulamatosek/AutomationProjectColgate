
from selenium import webdriver
from selenium.webdriver.remote.webdriver import WebDriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from bs4 import BeautifulSoup



popBoxLoginUrl = 'z'
ourColgateUrl = 'y'
popBoxProductUrl = 'x'


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


def getSession():
    capabilities = {'chromeOptions': {'useAutomationExtension': False}}

    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.get(ourColgateUrl)
    print(driver.session_id, driver.command_executor._url)
    file = open("sessionIdAndProduct.txt", "w")
    file.write(driver.session_id)
    file.write('\n')
    file.write(driver.command_executor._url)
    file.close()
    return driver


driver = getSession()

