import time
import ReadWriteDataFromExcel as RWDE
import BrowserElementProperties as BEP
import WebElementReusability as WER
import openpyxl
import xlsxwriter
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from pathlib import Path



#1. Registering New HCP in LunarHCPPortalV1

FilePath = str(Path().resolve()) + r'\Excel Files\UrlsForProject.xlsx'
Sheet = 'Portal Urls'
Url = str(RWDE.ReadData(FilePath, Sheet, 3, 3))

#driver = webdriver.Chrome(executable_path=r'C:\Program Files\Application\Browser\chromedriver_win32\chromedriver')
driver = webdriver.Chrome(executable_path = str(Path().resolve()) + '\Browser\chromedriver_win32\chromedriver')
driver.maximize_window()
driver.get(Url)


FilePath = str(Path().resolve()) + '\Excel Files\HCPRegistrationFromPortal.xlsx'

Seconds = 300 / 1000
Sheet = 'HCP Registration Page Data'
RowCount = RWDE.RowCount(FilePath, Sheet)

#time.sleep(Seconds)
#LoginButton = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[. = "Log in"]', Seconds)
#LoginButton.click()

for RowIndex in range(2, RowCount + 1):
    time.sleep(Seconds)
    Notamemberlink = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//a[. = "Not a member?"]', 60)
    Notamemberlink.click()
    #This is for filling data in the fields
    time.sleep(Seconds)
    FirstNameTextBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[2]/lightning-input//input', 60)
    if(str(RWDE.ReadData(FilePath, Sheet, RowIndex, 2)) != 'None'):
        FirstNameTextBox.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex, 2)))
    else:
        FirstNameTextBox.click()

    time.sleep(Seconds)
    LastNameTextBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[3]/lightning-input//input', 60)
    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 3)) != 'None'):
        LastNameTextBox.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 3))
    else:
        LastNameTextBox.click()

    time.sleep(Seconds)
    EmailTextBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[9]//input', 60)
    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 4)) != 'None'):
        EmailTextBox.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 4))
    else:
        EmailTextBox.click()

    time.sleep(Seconds)
    StreetTextBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[15]/lightning-input', 60)
    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 5)) != 'None'):
        StreetTextBox.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 5))
    else:
        StreetTextBox.click()

    time.sleep(Seconds)
    UserTypeDropDownList = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[22]//select', 60)
    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 6)) != 'None'):
        SelectUserTypeDropDownList = Select(UserTypeDropDownList)
        UserTypeDropDownListOptions = SelectUserTypeDropDownList.options
        SelectUserTypeDropDownList.select_by_visible_text(RWDE.ReadData(FilePath, Sheet, RowIndex, 6))
    else:
        UserTypeDropDownList.click()


    time.sleep(Seconds)
    if RWDE.ReadData(FilePath, Sheet, RowIndex, 7) == 'S':
        SiteAdminCheckBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[23]/label/span/span[1]', 60)
        SiteAdminCheckBox.click()

    time.sleep(Seconds)
    NPITextBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[29]//input', 60)
    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 8)) != 'None'):
        NPITextBox.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 8))
    else:
        NPITextBox.click()

    time.sleep(Seconds)
    ClinicTextBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[5]//input', 60)
    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 9)) != 'None'):
        ClinicTextBox.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 9))
    else:
        ClinicTextBox.click()

    time.sleep(Seconds)
    PhoneTextBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[11]//input', 60)
    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 10)) != 'None'):
        PhoneTextBox.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 10))
    else:
        PhoneTextBox.click()

    time.sleep(Seconds)
    CityTextBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[17]//input', 60)
    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 11)) != 'None'):
        CityTextBox.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 11))
    else:
        CityTextBox.click()

    time.sleep(Seconds)
    StateDropDownList = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[18]//select', 60)
    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 12)) != 'None'):
        SelectStateDropDownList = Select(StateDropDownList)
        StateDropDownListOptions = SelectStateDropDownList.options
        SelectStateDropDownList.select_by_visible_text(RWDE.ReadData(FilePath, Sheet, RowIndex, 12))
    else:
        StateDropDownList.click()

    time.sleep(Seconds)
    ZipCodeTextBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[19]//input', 60)
    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 13)) != 'None'):
        ZipCodeTextBox.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 13))
    else:
        ZipCodeTextBox.click()

    time.sleep(Seconds)
    FaxTextBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[25]//input', 60)
    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 14)) != 'None'):
        FaxTextBox.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 14))
    else:
        FaxTextBox.click()

    #time.sleep(Seconds)
    #if RWDE.ReadData(FilePath, Sheet, RowIndex, 15) == 'S':
    #    PhlebotomistCheckBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[31]/label/span/span[1]', 60)
    #    PhlebotomistCheckBox.click()

    # This is for catching validation errors
    ErrorMsg = ''
    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 2)) == 'None'):
        FirstNameErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[. = "First name is mandatory!"]', 60)
        print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, 1)) + ' : ' + FirstNameErrorMsg.text)
        ErrorMsg = FirstNameErrorMsg.text

    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 3)) == 'None'):
        LastNameErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[. = "Last name is mandatory!"]', 60)
        print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, 1)) + ' : ' + LastNameErrorMsg.text)
        ErrorMsg += ' ' + LastNameErrorMsg.text

    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 4)) == 'None'):
        EmailErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[. = "Email is mandatory!"]', 60)
        print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, 1)) + ' : ' + EmailErrorMsg.text)
        ErrorMsg += ' ' + EmailErrorMsg.text

    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 5)) == 'None'):
        StreetErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[. = "Street is mandatory!"]', 60)
        print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, 1)) + ' : ' + StreetErrorMsg.text)
        ErrorMsg += ' ' + StreetErrorMsg.text

    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 8)) == 'None'):
        NPIErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[. = "NPI is mandatory!"]', 60)
        print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, 1)) + ' : ' + NPIErrorMsg.text)
        ErrorMsg += ' ' + NPIErrorMsg.text

    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 9)) == 'None'):
        ClinicErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[. = "Hospital/Clinic is mandatory!"]', 60)
        print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, 1)) + ' : ' + ClinicErrorMsg.text)
        ErrorMsg += ' ' + ClinicErrorMsg.text

    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 10)) == 'None'):
        PhoneErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[. = "Phone is mandatory!"]', 60)
        print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, 1)) + ' : ' + PhoneErrorMsg.text)
        ErrorMsg += ' ' + PhoneErrorMsg.text

    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 11)) == 'None'):
        CityErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[. = "City is mandatory!"]', 60)
        print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, 1)) + ' : ' + CityErrorMsg.text)
        ErrorMsg += ' ' + CityErrorMsg.text

    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 12)) == 'None'):
        StateErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[. = "State is mandatory!"]', 60)
        print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, 1)) + ' : ' + StateErrorMsg.text)
        ErrorMsg += ' ' + StateErrorMsg.text

    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 13)) == 'None'):
        ZipCodeErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[. = "Zip Code is mandatory!"]', 60)
        print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, 1)) + ' : ' + ZipCodeErrorMsg.text)
        ErrorMsg += ' ' + ZipCodeErrorMsg.text

    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 14)) == 'None'):
        FaxErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[. = "Fax is mandatory!"]', 60)
        print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, 1)) + ' : ' + FaxErrorMsg.text)
        ErrorMsg += ' ' + FaxErrorMsg.text

    time.sleep(Seconds)
    SubmitButton = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[. = "Submit"]', 60)
    SubmitButton.click()

    time.sleep(Seconds)
    ErrorElementAvailable = WER.check_exists_by_xpath(driver, '//div[3]/div/div/div[3]/div[1]/div')
    if (ErrorElementAvailable == True):
        RequiredErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[3]/div/div/div[3]/div[1]/div', 60)
        ErrorMsg += ' ' + RequiredErrorMsg.text

    time.sleep(2)
    ErrorElementAvailable = WER.check_exists_by_xpath(driver, '//div[3]/div/div/div[3]/div[1]/div')
    if (ErrorElementAvailable == True):
        RequiredErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[3]/div/div/div[3]/div[1]/div', 60)
        ErrorMsg += ' ' + RequiredErrorMsg.text

    if (ErrorMsg != ' Loading Loading'):
        if (ErrorMsg != ''):
            time.sleep(Seconds)
            ErrorElementAvailable = WER.check_exists_by_xpath(driver, '//div[3]/div/div/div[3]/div[1]/div')
            if (ErrorElementAvailable == True):
                RequiredErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[3]/div/div/div[3]/div[1]/div', 60)
                ErrorMsg += ' ' + RequiredErrorMsg.text
                RWDE.WriteData(FilePath, Sheet, RowIndex, 16, 'Hold')
                RWDE.WriteData(FilePath, Sheet, RowIndex, 17, ErrorMsg)
    else:
        RWDE.WriteData(FilePath, Sheet, RowIndex, 16, 'Passed')

    time.sleep(5)
    LoginButton = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[. = "Log in"]', 60)
    LoginButton.click()

driver.quit()