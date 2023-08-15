import gspread, keyboard, time
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# Keep the window open
options = Options()
options.add_experimental_option("detach", True)
# disable the banner "Chrome is being controlled by automated test software"
options.add_experimental_option("useAutomationExtension", False)
options.add_experimental_option("excludeSwitches", ["enable-automation"])
driver = webdriver.Chrome(options=options)

# Enable google sheets API
scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

creds = ServiceAccountCredentials.from_json_keyfile_name(
    "INSERT CREDS FILE PATH", scopes=scopes
)
file = gspread.authorize(creds)
workbook = file.open("INSERT SPREADSHEET NAME")
sheet = workbook.sheet1

driver.get("INSERT LINK")

try:
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Load More"))
    )
except:
    driver.quit()

loadMoreButton = driver.find_element(By.LINK_TEXT, "Load More")
for i in range(7):
    loadMoreButton.click()
time.sleep(1)
driver.execute_script("arguments[0].scrollIntoView(true);", loadMoreButton)

table = driver.find_element(
    By.XPATH,
    "//table[@class='table table-striped table-bordered table-hover table-condensed ng-scope']",
)
body = table.find_element(By.TAG_NAME, "tbody")
rows = body.find_elements(By.TAG_NAME, "tr")

action = ActionChains(driver)
tabs = 0
tabsToBeOpened = "INSERT NUMBER OF TABS TO BE OPENED"

# Open only a certain number of tabs
for row in reversed(rows):
    tabs += 1
    if tabs > tabsToBeOpened:
        break
    else:
        element = row.find_element(By.TAG_NAME, "a")
        ActionChains(driver).move_to_element(element).perform()
        action.key_down(Keys.CONTROL).click(element).key_up(Keys.CONTROL).perform()

# Gather the other tab handles before closing the current tab
otherTabs = driver.window_handles
driver.close()


def updateSheet(row, applicant, result, reason=None):
    sheet.update_cell(i, 1, applicant)
    sheet.update_cell(i, 2, result)
    print(i, applicant)
    if reason != None:
        print(reason)
        sheet.update_cell(i, 3, reason)


keyList = ["1", "2", "3", "4", "5", "6", "q", "e", "r", "a", "s", "d", "f", "x", "c"]

for i in range(1, tabs):
    driver.switch_to.window(otherTabs[i])
    try:
        test = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located(
                (By.XPATH, "//button[@class='btn btn-danger']")
            )
        )
    except:
        print("Page too long to load. Reset script")
    applicant = driver.current_url[81:]
    key = keyboard.read_key()
    while key not in keyList:
        print("Rechecking for valid option")
        key = keyboard.read_key()
    if key == "1":
        updateSheet(i, applicant, "APPROVE")
    elif key == "2":
        updateSheet(i, applicant, "REJECT")
    elif key == "3":
        updateSheet(i, applicant, "RESET", "VIDEO: Frozen recording")
    elif key == "4":
        updateSheet(i, applicant, "RESET", "VIDEO: Sideways recording")
    elif key == "5":
        updateSheet(i, applicant, "RESET", "VIDEO: Black screen")
    elif key == "6":
        updateSheet(i, applicant, "RESET", "VIDEO: Wrong screen")
    elif key == "q":
        updateSheet(i, applicant, "RESET", "PERFORMANCE: NoArticle")
    elif key == "e":
        updateSheet(i, applicant, "RESET", "PERFORMANCE: Bkgrd Noise")
    elif key == "r":
        updateSheet(i, applicant, "RESET", "PERFORMANCE: MC")
    elif key == "a":
        updateSheet(i, applicant, "RESET", "AUDIO: NoVerbal")
    elif key == "s":
        updateSheet(i, applicant, "RESET", "AUDIO: PoorAudio")
    elif key == "d":
        updateSheet(i, applicant, "REJECT", "Already reset once; no article")
    elif key == "f":
        updateSheet(i, applicant, "REJECT", "Already reset once; background noise")
    elif key == "x":
        updateSheet(i, applicant, "REJECT", "Already reset once; wrong mc")
    elif key == "c":
        updateSheet(i, applicant, "REJECT", "Already reset once; poor audio")
    key = keyboard.read_key()
    while key != "w":
        print("Rechecking for closed window")
        key = keyboard.read_key()
    driver.close()
