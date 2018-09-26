#Magpie Pricer: Takes ISBNs and compares prices at 3 book buying websites, returning the best price
#Copyright StannoSoft 2018 <stevetasticsteve@gmail.com>

##    This program is free software: you can redistribute it and/or modify
##    it under the terms of the GNU General Public License as published by
##    the Free Software Foundation, either version 3 of the License, or
##    (at your option) any later version.
##
##    This program is distributed in the hope that it will be useful,
##    but WITHOUT ANY WARRANTY; without even the implied warranty of
##    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
##    GNU General Public License for more details.
##
##    You should have received a copy of the GNU General Public License
##    along with this program.  If not, see <https://www.gnu.org/licenses/>.

#Dependencies: Selenium, Google Chrome, Google API key, google-api-python-client, openpyxl 
#Specific module imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, ElementNotVisibleException, TimeoutException
from selenium.webdriver.common.keys import Keys
from apiclient.discovery import build
from musicMagpiePricer import searchMomox, waitForID, timeOut
import os, sys

if __name__ == '__main__':

    scriptFolder = os.path.dirname(os.path.realpath(__file__))                          #Folder program is contained in
    rootFolder = scriptFolder.rstrip('\\Scripts')
    resultsFolder = os.path.join(rootFolder, 'Results')
    momoxPath = os.path.join(resultsFolder, 'momoxBooks.txt')
    try:
        momoxFile = open(momoxPath)
    except FileNotFoundError:
        print('No ISBNS to process')
        input()
        sys.exit()
        
    data = momoxFile.readlines()
    ISBNs = []
    for book in data:
        ISBNs.append(book.split(',')[0])

    ISBNs = set(ISBNs)

    chromeDriverPath = os.path.join(scriptFolder, 'chromedriver.exe')
    browser = webdriver.Chrome(chromeDriverPath)
    browser.get("https://www.momox.co.uk")

    for ISBN in ISBNs:
        
        searchMomox(ISBN, timeOut, browser)
        addToBasketID = 'buttonAddToCart'
        waitForID('Momox', addToBasketID, timeOut, browser)
        element = browser.find_element_by_id(addToBasketID)
        browser.execute_script("arguments[0].click();", element)
    
    momoxFile.close()

    print(str(len(ISBNs)) + ' ISBNs added to basket, is this correct, y or n?')
    ans = input()

    if ans == 'y':
           os.remove(momoxPath)

    sys.exit()
