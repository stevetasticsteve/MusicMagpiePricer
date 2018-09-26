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
from musicMagpiePricer import searchMagpie, waitForID, timeOut
import os, sys

if __name__ == '__main__':

    scriptFolder = os.path.dirname(os.path.realpath(__file__))                          #Folder program is contained in
    rootFolder = scriptFolder.rstrip('\\Scripts')
    resultsFolder = os.path.join(rootFolder, 'Results')
    magpiePath = os.path.join(resultsFolder, 'musicmagpieBooks.txt')
    try:
        magpieFile = open(magpiePath)
    except FileNotFoundError:
        print('No ISBNS to process')
        input()
        sys.exit()
        
    data = magpieFile.readlines()
    ISBNs = []
    for book in data:
        ISBNs.append(book.split(',')[0])

    ISBNs = set(ISBNs)

    chromeDriverPath = os.path.join(scriptFolder, 'chromedriver.exe')
    browser = webdriver.Chrome(chromeDriverPath)
    browser.execute_script("window.open('https://www.musicmagpie.co.uk/sell-books', 'tab2');")

    for ISBN in ISBNs:
        
        searchMagpie(ISBN, timeOut, browser)
    
    magpieFile.close()

    print(str(len(ISBNs)) + ' ISBNs added to basket, is this correct, y or n?')
    ans = input()

    if ans == 'y':
           os.remove(magpiePath)

    sys.exit()
