from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Keep the window open
options = Options()
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(options=options)

driver.get("INSERT LINK HERE")

# Wait for the element I'm looking for to exist on the page
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
i = 0

# Open 200 tabs
for row in reversed(rows):
    i += 1
    if i > 200:
        break
    element = row.find_element(By.TAG_NAME, "a")
    ActionChains(driver).move_to_element(element).perform()
    action.key_down(Keys.CONTROL).click(element).key_up(Keys.CONTROL).perform()
