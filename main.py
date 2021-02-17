import time
import logging
import os
import glob
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from pathlib import Path

#from tkinter import *
#from functools import partial

#gblPassword = ""
#gblUsername = ""

def validateLogin(username, password):
	print("username entered :", username.get())
	print("password entered :", password.get())
	return

def initialize_logger():
    global logger
    logger = logging.getLogger(__name__)

    # set log minimum level
    logger.setLevel(logging.DEBUG)

    # define file handler and set formatter
    file_handler = logging.FileHandler('Python.log', mode='w')
    formatter = logging.Formatter('%(asctime)s : %(levelname)s : %(name)s : %(message)s')
    file_handler.setFormatter(formatter)

    # add file handler to logger
    logger.addHandler(file_handler)

# END initialize_logger()

def login_to_SL2015(driver):
    logger.info('ENTER login_to_SL2015()')
    logger.info('Logging in to SL2015')
    driver.get("https://kendall.logikos.com/")

#    gblUsername = username.get()
#    gblPassword = password.get()

    # log in
    select = Select(driver.find_element_by_id('DSLCpnyID'))
    select.select_by_visible_text('Logikos, Inc.')

    elem = driver.find_element_by_id("DSLUserID")
    elem.clear()
    elem.send_keys("LOGFWA\dreser")
#    elem.send_keys("LOGFWA\\" + gblUsername)

    elem = driver.find_element_by_id("DSLPassword")
    elem.clear()
    elem.send_keys("Dogmeat12312020")
#    elem.send_keys(gblPassword)

    elem.send_keys(Keys.RETURN)
    logger.info('LEAVE login_to_SL2015()')

# END login_to_SL2015()

def go_to_active_projects_list(driver):

    logger.info('ENTER go_to_active_projects_list()')
    logger.info('Going to list of active projects.')
    # Click on Project to expand sub-options.
    elem = driver.find_element_by_class_name('ui-collapsible-heading-toggle').click()

    # Click on Project Analyst
    elem = driver.find_element_by_partial_link_text('Project Analyst').click()
    time.sleep(10)

    # Select the Project Status combo box.
    elem = driver.find_element_by_id("RSReportViewer_ctl04_ctl19_txtValue")
    elem.click()

    # Disable the checkbox for all project statuses.
    elem = driver.find_element_by_id("RSReportViewer_ctl04_ctl19_divDropDown_ctl00")
    elem.click()

    # Enable the checkbox for the Active projects.
    elem = driver.find_element_by_id("RSReportViewer_ctl04_ctl19_divDropDown_ctl02")
    elem.click()

    # Click on View Report to see the subset of active projects.
    elem = driver.find_element_by_id("RSReportViewer_ctl04_ctl00")
    elem.click()
    time.sleep(3)

    logger.info('LEAVE go_to_active_projects_list()')

# END go_to_active_projects_list()

def rename_csv_file(project):
    logger.info('ENTER rename_csv_file()')
    try:
        directory = Path('C:/Users/dreser/Downloads/')
        old_pathname = directory / 'BillingDetailInquiry.csv'
        new_filename = project + '_BillingDetailInquiry.csv'
        new_pathname = directory / new_filename
        os.rename(old_pathname, new_pathname)
    except:
        logger.error('Rename of csv file failed')
    logger.info('LEAVE rename_csv_file()')
# END rename_csv_file


def process_project(driver,xpath,project):
    logger.info('ENTER process_project()')

    logger.info('Process project')
    logger.info('Click on Unbilled Summary - Total')
    # Click on Unbilled Summary - Total
    elem = driver.find_element_by_xpath("//*[contains(text(), 'Unbilled Summary')]/parent::td/parent::tr/parent::tbody")
    elem = elem.find_element_by_xpath(".//*[contains(text(), 'Total')]/parent::td/following-sibling::td/child::div/child::a")
    elem.send_keys(Keys.NULL) #give element focus
    elem.click()
    time.sleep(5)

    # Select "All" for Billing Status and click on View Report button.
    logger.info('Select ALL for Billing Status')
    select = Select(driver.find_element_by_id('RSReportViewer_ctl04_ctl11_ddValue'))
    select.select_by_visible_text('All')
    time.sleep(1)
    elem = driver.find_element_by_id("RSReportViewer_ctl04_ctl00")
    elem.click()
    time.sleep(5)

    # Select to save report as CVS.
    logger.info('Save CVS file.')
    elem = driver.find_element_by_id("RSReportViewer_ctl05_ctl04_ctl00_ButtonImg")
    elem.click()
    time.sleep(5)
    elem = driver.find_element_by_xpath("//*[contains(text(),'CSV (comma delimited)')]")
    elem.click()
    time.sleep(5)

    rename_csv_file(project)

    # Go back to list of active projects.
    logger.info('Return to list of active projects.')
    elem = driver.find_element_by_id("RSReportViewer_ctl05_ctl01_ctl00_ctl00_ctl00")
    elem.click()
    time.sleep(3)

    logger.info('Make sure project is visible on screen before collapsing project details')
    elem = driver.find_element_by_xpath(xpath)
    desired_y = (elem.size['height'] / 2) + elem.location['y']
    window_h = driver.execute_script('return window.innerHeight')
    window_y = driver.execute_script('return window.pageYOffset')
    current_y = (window_h / 2) + window_y
    scroll_y_by = desired_y - current_y
    driver.execute_script("window.scrollBy(0, arguments[0]);", scroll_y_by)
    elem.click()
    time.sleep(3)

    logger.info('LEAVE process_project()')
# END process_project()

def move_to_next_page_of_active_projects(driver):
    logger.info('ENTER move_to_next_page_of_active_projects()')
    # Go to next page of active projects.
    logger.info('Moving to next page of active projects.')
    elem = driver.find_element_by_id("RSReportViewer_ctl05_ctl00_Next_ctl00_ctl00")
    elem.click()
    time.sleep(3)

    logger.info('LEAVE move_to_next_page_of_active_projects()')
# END move_to_next_page_of_active_projects()

def delete_all_billingdetailinquery_csv_files():
    logger.info('ENTER delete_all_csv_files()')

    # Get a list of all the BillingDetailInquiry files in the Download folder.
    file_expression = 'C:\\Users\\dreser\\Downloads\\*BillingDetailInquiry*.csv'
    fileList = glob.glob(file_expression)

    # Iterate over the list of filepaths & remove each file.
    for filePath in fileList:
        logger.info('Deleting {}'.format(filePath))
        try:
            os.remove(filePath)
        except:
            logger.error('Error while deleting file : {}'.format(filePath))

    logger.info('LEAVE delete_all_csv_files()')
# END delete_all_billingdetailinquery_csv_files()

# ===========================================================================================
# Main
# ===========================================================================================

initialize_logger()

logger.info('ENTER main()')
logger.info('********** Starting Active Projects processing **********')

# Get list of active projects.
logger.info('Creating list of active projects.')
activeProjects = []
#
# TODO - Update to ignore blank lines.
#
with open("ActiveProjects.txt", "r") as f:
    for line in f:
        if(line.startswith('#')):  # skip comment lines
            continue
        line.rstrip("\n")
#        fields = line.split(", ")
#        dim = fields.__slots__
        logger.info(line.rstrip("\n"))
        activeProjects.append(line.rstrip("\n"))
#        activeProjects.append(line.rstrip("\n"))
logger.info('End list of active projects')

logger.info('Starting WebDriver')

# Delete any old BillingDetailInquiry csv files from Downloads folder
delete_all_billingdetailinquery_csv_files()

# Create a web criver instance and log in to SL2015
driver = webdriver.Chrome()
#getLoginCredentials()
login_to_SL2015(driver)

# In SL2015 filter on just active projects.
go_to_active_projects_list(driver)

# Loop through all active projects.
logger.info('Process each active project')
for project in activeProjects:
    # Click on project to see details.
    found = False
    while not found:
        try:
            logger.info('****** Processing {} ******'.format(project))
            xpath = "//*[contains(text(),  " + "'" + project + "'" + ")]//child::img"
            elem = driver.find_element_by_xpath(xpath)
            found = True
            driver.execute_script("arguments[0].scrollIntoView();", elem)
            elem.click()
            time.sleep(3)
        except NoSuchElementException:
            # Look on next page for the project
            logger.info('Project not on current page, moving to next page.')
            move_to_next_page_of_active_projects(driver)
        # END try
    # END while not found

    # Get the actuals for the project and save as a csv file.
    process_project(driver,xpath,project)

#end for project in activeProjects

driver.close()

logger.info('LEAVE main()')
# END Main()