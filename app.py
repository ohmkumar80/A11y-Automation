from selenium import webdriver
from selenium.webdriver.common.by import By
from axe_selenium_python import Axe
import pandas as pd

options = webdriver.ChromeOptions()
# options.add_argument('headless')
options.add_argument('window-size=1024x768')
driver = webdriver.Chrome(options=options)

# Read the URLs from the Excel file
urls_df = pd.read_excel('urls.xlsx')
urls = urls_df['URL'].tolist()

# Login credentials
username = "xxx"
password = "xxx"

try:
    writer = pd.ExcelWriter("a11y.xlsx", engine="xlsxwriter")

    for i, url in enumerate(urls):
        driver.get(url)
        page_title = driver.title

        axe = Axe(driver)
        axe.inject()
        results = axe.run()
        violations_df = pd.DataFrame(results["violations"])

        # Add page URL and page title to the violations_df DataFrame
        violations_df["Page URL"] = url
        violations_df["Page Title"] = page_title

        # Take a screenshot of the page
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        driver.execute_script("window.scrollTo(0, 0);")
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        driver.execute_script("window.scrollTo(0, 0);")
        screenshot_file = f"Sheet {i + 1}.png"
        print(f"Saving screenshot to {screenshot_file}")
        driver.save_screenshot(screenshot_file)

        # Insert the screenshot into the worksheet
        sheet_name = f"Sheet {i + 1}"
        worksheet = writer.book.add_worksheet(sheet_name)
        worksheet.insert_image("K1", screenshot_file)
        violations_df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"{len(results['violations'])} accessibility violation(s) found on {url}")
        print(axe.report(results["violations"]))

        # Enter login credentials and login if it's a login page
        if "login" in driver.current_url.lower():
            driver.find_element(By.NAME, 'myusername').send_keys(username)
            driver.find_element(By.NAME, 'mypassword').send_keys(password)
            driver.find_element(By.NAME, 'Submit').click()

    writer._save()

finally:
    driver.quit()
