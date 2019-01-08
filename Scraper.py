from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from tabulate import tabulate
import pandas as pd
from functools import reduce

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--incognito")
driver = webdriver.Chrome(options = chrome_options)

driver.get("https://my.mskcc.org/login")
WebDriverWait(driver, 10).until(EC.visibility_of_all_elements_located((By.XPATH, "//*[@id='UserName']")))
driver.find_element_by_id('UserName').send_keys('#INSERTUSERHERE')
driver.find_element_by_id('Password').send_keys('#INSERTPASSWORDHERE')
driver.find_element_by_id('btnSubmitLogin').click()
WebDriverWait(driver, 10).until(EC.visibility_of_all_elements_located((By.XPATH, "//*[@id='nav_medical-info']/a/span")))



driver.find_element_by_xpath("//*[@id='ctl00_ctl00_MAIN_CONTENT_divLabResults']/a").click()
WebDriverWait(driver, 10).until(EC.visibility_of_all_elements_located((By.ID,
            "ctl00_ctl00_MAIN_CONTENT_cmbLabSuggestSearchDate_Input")))

driver.find_element_by_xpath("//*[@id='ctl00_ctl00_MAIN_CONTENT_cmbLabSuggestSearchDate']/span/button").click()
WebDriverWait(driver, 10).until(EC.visibility_of_all_elements_located((By.XPATH,
            "//*[@id='ctl00_ctl00_MAIN_CONTENT_cmbLabSuggestSearchDate_DropDown']/div/ul/li[1]")))

#hover = ActionChains(driver).move_to_element(driver.find_element_by_id(
# 'ctl00_ctl00_MAIN_CONTENT_cmbLabSuggestSearchDate_DropDown'))
#hover.perform()

ds = driver.find_elements_by_class_name("rcbList")

dateslist = [str(d.text).split() for d in ds][0]
dateslist.reverse()
print(dateslist)

soups = []
dfs = []

for i,x in enumerate(dateslist):
    driver.find_element_by_id('ctl00_ctl00_MAIN_CONTENT_cmbLabSuggestSearchDate_Input').send_keys(dateslist[i] + Keys.RETURN)
    WebDriverWait(driver, 10).until(EC.visibility_of_all_elements_located((By.XPATH,
            "//*[@id='ctl00_ctl00_MAIN_CONTENT_tblLatestOrderPanel']")))

    soups.append(BeautifulSoup(driver.page_source, 'lxml'))
    dfs.append(pd.read_html(str(soups[i]))[0])

    driver.find_element_by_id("ctl00_ctl00_MAIN_CONTENT_hlClearSearch").click()
    WebDriverWait(driver, 10).until(EC.visibility_of_all_elements_located((By.XPATH,
            "//*[@id='ctl00_ctl00_MAIN_CONTENT_tblLatestOrderPanel']")))

driver.close()

for i,x in enumerate(dateslist):
    dfs[i] = dfs[i].loc[:,1:2].transpose()
    dfs[i].insert(loc = 0, column = 'Date', value = ['Date', x])
    dfs[i] = dfs[i].transpose()
    # colnames = dfs[i].iloc[0]
    # dfs[i] = dfs[i][1:]
    # dfs[i].columns = colnames

merged = reduce(lambda ldf, rdf: pd.merge(ldf, rdf, on = 1, how = 'outer'), dfs)


# write each dataframe to a sheet of an excel file
writer = pd.ExcelWriter('Scraped_Data.xlsx')
for n, df in enumerate(dfs):
    df.to_excel(writer,'sheet%s' % n)
writer.save()
#print(tabulate(df[0], headers='keys', tablefmt='psql')) # print pretty df to console