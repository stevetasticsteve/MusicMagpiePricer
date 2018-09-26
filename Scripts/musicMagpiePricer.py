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

#Dependencies: Selenium, Google Chrome, openpyxl 
#Specific module imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, ElementNotVisibleException, TimeoutException
from selenium.webdriver.common.keys import Keys
#Standard moudle imports
import time, sys, os, pprint, logging, datetime, openpyxl
startTime = time.time()
logging.basicConfig(filename='logfile.log',level=logging.INFO)
logging.info('\n\nProgram started: ' + datetime.datetime.now().strftime('%H'':''%M''.''%S'' ''%d''/''%m''/''%Y') +
             '============================================================================')

#Options
##debugMode = 'off'
timeOut = 5                                                                         #Website timeout wait
ISBNFileName = 'ISBNToProcess.txt'                                                  #Name of file containing ISBN numbers (for batch mode)
headless = 'on'                                                                    #'on' or 'off' does the chrome window show or not?
                                                                                    #to check the result before adding to sale lists. 'no' skips the save entirely and 'save' saves but skips the question
###
#Functions
###
def errorHandling():
    return 'Error: {}. {}, line: {}'.format(sys.exc_info()[0],
                                         sys.exc_info()[1],
                                         sys.exc_info()[2].tb_lineno)

def waitForCSS(websiteName, selector, timeOut, browser):
    try:
        WebDriverWait(browser, timeOut).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
    except NoSuchElementException as e:
        logging.error(websiteName + ' seem to have changed their website.')
        logging.error("Try replacing the CSS code for the affected element")
        logging.error(e, exc_info=True)
        browser.quit()
        sys.exit()
    except ElementNotVisibleException as e:
        logging.error(e, exc_info=True)
        logging.error(websiteName + ' Element website not responding')
        logging.error("Try replacing the CSS code for the affected element")
        browser.quit()
        sys.exit()
    except TimeoutException as e:
        logging.error(e, exc_info=True)
        logging.error(websiteName + ' timed out and is not responding')
        logging.error(errorHandling())
##        browser.quit()
##        sys.exit()

def waitForID(websiteName, selector, timeOut, browser):
    try:
        WebDriverWait(browser, timeOut).until(
            EC.presence_of_element_located((By.ID, selector)))
    except NoSuchElementException as e:
        logging.error(websiteName + ' seem to have changed their website.')
        logging.error("Try replacing the ID for the affected element")
        logging.error(e, exc_info=True)
        browser.quit()
        sys.exit()
    except ElementNotVisibleException as e:
        logging.error(websiteName + ' Element website not responding')
        logging.error("Try replacing the ID for the affected element")
        logging.error(e, exc_info=True)
        browser.quit()
        sys.exit()
    except TimeoutException as e:
        logging.error(websiteName + ' timed out and is not responding')
        logging.error("Try replacing the ID for the affected element")
        logging.error(errorHandling())
##        browser.quit()
##        sys.exit()

def searchMomox(ISBN, timeOut, browser):
    browser.switch_to.window(browser.window_handles[0])
    momoxSearchCSS = ".media-search-input.manually .product-input"

    waitForCSS("Momox", momoxSearchCSS, timeOut, browser)
    momoxSearch = browser.find_element_by_css_selector(momoxSearchCSS)
    momoxSearch.clear()
    momoxSearch.send_keys(ISBN)
    momoxSearch.send_keys(Keys.RETURN)

def readMomox(timeOut):
    browser.switch_to.window(browser.window_handles[0])
    momoxResultCSS = ".searchresult .searchresult-price-block .searchresult-price"
    waitForCSS("Momox", momoxResultCSS, timeOut, browser)
    if browser.find_element_by_css_selector('h1, .h1').get_attribute('innerHTML') == ('Item not found'):
        logging.info('Timeout exception handled, invalid or unknown ISBN, book skipped')
        return 0
    else:
        title = (browser.find_element_by_css_selector('h1, .h1').get_attribute('innerHTML'))
    momoxResult = browser.find_element_by_css_selector(momoxResultCSS)
    momoxString = momoxResult.get_attribute("innerHTML")
    momoxFloat = float(momoxString.lstrip('£'))
    return momoxFloat, title

def searchMagpie(ISBN, timeOut, browser):
    magpieSearchCSS = "#value .container .valuationEngineMediaTabContent.active input, #value .container .valuationEngineTechTabContent.active input"
    browser.switch_to.window("tab2")
    
    waitForCSS("Music Magpie", magpieSearchCSS, timeOut, browser)
    magpieSearch = browser.find_element_by_css_selector(magpieSearchCSS)
    magpieSearch.clear()
    magpieSearch.send_keys(ISBN)
    magpieSearch.send_keys(Keys.RETURN)

def readMagpie(timeOut):
    magpieResultCSS = "#basketareaWrapper .basketareaContent .left .row .col_Price"
    browser.switch_to.window("tab2")    
    
    waitForCSS("Music Magpie", magpieResultCSS, timeOut, browser)
    magpieResult = browser.find_element_by_css_selector(magpieResultCSS)
    magpieString = magpieResult.get_attribute("innerHTML")
    magpieFloat = float(magpieString)
    return magpieFloat

def searchWebuy(ISBN, timeOut, browser):  #Webuybooks provides a basket total, not price per book
    webuySearchCSS = "section.home-isbn-input .input-container .inner-input-container form .search-container input, section.home-isbn-input-slider .input-container .inner-input-container form .search-container input, section.innerIsbnInput .input-container .inner-input-container form .search-container input, .input-area .input-container .inner-input-container form .search-container input"
    browser.switch_to.window("tab3")
    
    waitForCSS("We buy books", webuySearchCSS, timeOut, browser)
    webuySearch = browser.find_element_by_css_selector(webuySearchCSS)
    webuySearch.clear()
    webuySearch.send_keys(ISBN)
    webuySearch.send_keys(Keys.RETURN)

def readWebuy(timeOut):
    webuyResultID = "headerPrice"
    browser.switch_to.window("tab3")

    waitForID("We buy books", webuyResultID, timeOut, browser)
    webuyResult = browser.find_element_by_id(webuyResultID)
    webuyString = webuyResult.get_attribute("innerHTML").lstrip('(').rstrip(')')
    webuyFloat = float(webuyString.lstrip('(£'))
    return webuyFloat

def getBookPrices(ISBN, timeout):
    global webuyTotal, allResults
#Send search
    searchMomox(ISBN, timeOut, browser)
    searchMagpie(ISBN, timeOut, browser)
    searchWebuy(ISBN, timeOut, browser)
# Read results
    momoxResult = readMomox(timeOut)
    if momoxResult == 0:
        return 0
    title = momoxResult[1]
    magpieResult = readMagpie(timeOut)
#Webuy is slow to update. This loop checks it 3 times if price is unchanged
    counter = 0
    while counter <=3:
        webuyNewTotal = readWebuy(timeOut)
        if webuyNewTotal == webuyTotal:
            if counter < 3:
                time.sleep(1)
                counter += 1
            elif counter == 3:
                break
        elif webuyNewTotal > webuyTotal:
            break      
    webuyResult = webuyNewTotal - webuyTotal
    webuyTotal = webuyNewTotal
    
#Store results in a list of dictionaries
    prices = {'momox': momoxResult[0],'magpie': magpieResult,
              'webuy': webuyResult}
    bestPrice = max(prices, key=prices.get)
    
    dictTitle = {'ISBN' : ISBN, 'title': title, 'momox': momoxResult[0],
                                'magpie': magpieResult, 'webuy': webuyResult, 'bestPrice': bestPrice}
    allResults.append(dictTitle)
    return dictTitle

def saveResults():
    try:
        wb = openpyxl.load_workbook(os.path.join(resultsFolder, 'Results.xlsx'))
        logging.debug('Excel spreadsheet openened')
        newSpreadsheet = False
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        newSpreadsheet = True
        logging.debug('No Excel spreadsheet, one created')

    sheet = wb.active
    sheet.title = 'Best prices'
    sheet.column_dimensions['A'].width = 60
    sheet.column_dimensions['D'].width = 15
    if newSpreadsheet:
        lastRow = 0
    else:
        lastRow = sheet.max_row

    for rowNum, book in enumerate(allResults, start = lastRow):
        rowNum += 1
        rowNum = str(rowNum)

        if book['bestPrice'] == 'momox':            
            colour = openpyxl.styles.colors.Color(rgb='FF3333')  #red
        elif book['bestPrice'] == 'webuy':
            colour = openpyxl.styles.colors.Color(rgb='99FF33')  #Green
        elif book['bestPrice'] == 'magpie':
            colour = openpyxl.styles.colors.Color(rgb='3399FF')  #Blue

        if book['ISBN'] in duplicates:
            colour = openpyxl.styles.colors.Color(rgb='FFFF00')  #yellow

        fill = openpyxl.styles.fills.PatternFill(patternType='solid', fgColor = colour)
 
        sheet['A' + rowNum] = book['title']
        sheet['A' + rowNum].fill = fill
        sheet['B' + rowNum] = book['bestPrice']
        sheet['B' + rowNum].fill = fill
        sheet['C' + rowNum] = book[book['bestPrice']]
        sheet['C' + rowNum].fill = fill
        sheet['D' + rowNum] = book['ISBN']
        sheet['D' + rowNum].fill = fill
    while True:
        try:
            wb.save(os.path.join(resultsFolder, 'Results.xlsx'))
            break
        except PermissionError:
            print('\n!!!Permission error: Close the Excel workbook and press return!!!')
            print('hit Ctrl + c to exit if this freezes')
            logging.error('permission error')
            input()
    
    momoxFilePath = os.path.join(resultsFolder, 'momoxBooks.txt')
    webuyFilePath = os.path.join(resultsFolder, 'webuybooksBooks.txt')
    magpieFilePath = os.path.join(resultsFolder, 'musicmagpieBooks.txt')
    momoxFile = open(momoxFilePath, 'a')
    webuyFile = open(webuyFilePath, 'a')
    magpieFile = open(magpieFilePath, 'a')
    for thing in allResults:
        if thing['bestPrice'] == 'momox':
            momoxFile.write(thing['ISBN'] + ', ' + thing['title'] + '\n')
        elif thing['bestPrice'] == 'magpie':
            magpieFile.write(thing['ISBN'] + ', ' + thing['title'] + '\n')
        elif thing['bestPrice'] == 'webuy':
            webuyFile.write(thing['ISBN'] + ', ' + thing['title'] + '\n')
        else:
            pass
    momoxFile.close()
    webuyFile.close()
    magpieFile.close()

def ISBNDuplicatesCheck():
    duplicates = []
    for ISBN in ISBNList:
        if ISBNList.count(ISBN) > 1:
            duplicates.append(ISBN)
    return set(duplicates)
    
###
#
###

#Blank variables for looping and storing data
if __name__ == "__main__":

    #Load ISBNS
    print('Unleashing the python\n')
    scriptFolder = os.path.dirname(os.path.realpath(__file__))                          #Folder program is contained in
    rootFolder = scriptFolder.rstrip('\\Scripts')
    resultsFolder = os.path.join(rootFolder, 'Results')
    ISBNPath = os.path.join(rootFolder, ISBNFileName)    
    ISBNFile = open(ISBNPath)
    ISBNList = ISBNFile.readlines()
    logging.info(str(len(ISBNList)) + ' ISBNs read from file')
    print(str(len(ISBNList)) + ' ISBNs loaded')
    for i, ISBN in enumerate(ISBNList):
        ISBNList[i] = ISBNList[i].rstrip('\n')
    duplicates = ISBNDuplicatesCheck()        

    # create a new Chrome session
    try:
        chromeDriverPath = os.path.join(scriptFolder, 'chromedriver.exe')
    except Exception:
        logging.critical('Chromdriver.exe is missing, download it and place it into the Scripts folder')
        logging.shutdown()
        sys.exit()
    chromeOptions = webdriver.ChromeOptions()
    chromeOptions.add_argument("--log-level=3")
    if headless == 'on':
        chromeOptions.add_argument("--headless")
    prefs = {'profile.managed_default_content_settings.images':2}
    chromeOptions.add_experimental_option("prefs", prefs)
    browser = webdriver.Chrome(chromeDriverPath, chrome_options=chromeOptions)
    # Navigate to web pages
    browser.get("https://www.momox.co.uk")
    print('Python is pythoning... be patient')
    browser.execute_script("window.open('https://www.musicmagpie.co.uk/sell-books', 'tab2');")
    browser.execute_script("window.open('https://www.webuybooks.co.uk/', 'tab3');")

    webuyTotal = 0.0
    allResults = []

    #Loop through ISBNS
    for number, ISBN in enumerate(ISBNList):
        if getBookPrices(ISBN, timeOut) != 0:
            print('Book ' + str(number+1) + ' processed: ' + allResults[number]['title']
                 + ', £' + str(allResults[number][allResults[number]['bestPrice']]) + ', ' + allResults[number]['bestPrice'] )

        
    #House keeping
    browser.quit()
    endTime = time.time()
    runTime = endTime - startTime
    cycleTime = runTime/len(ISBNList)
    logging.info('Run time = ' + str(runTime) + 's, '+ str(cycleTime) +'s per cycle')

    #Show the results
    print('\nResults:\n')
    for book in allResults:
        spaces = 60
        bookTitle = book['title'].ljust(spaces)
        print(bookTitle + ': ' + book['bestPrice'] + ' £' +str(book[book['bestPrice']]))


    #Saving batch output
    saveResults()
    summaryMessage = ('\nProgram finished, ' + str(len(allResults)) + '/' + str(len(ISBNList))
          + ' ISBNs processed.')
    print(summaryMessage)
    logging.info(summaryMessage)
    print('Press any key to continue')
    input()
    logging.info('End of Program ============================================================================\n')
    logging.shutdown()
    os.startfile(os.path.join(resultsFolder, 'Results.xlsx'))
    sys.exit()
