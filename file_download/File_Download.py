from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from datetime import datetime
from extract_zip import extract_zip
from selenium.webdriver.firefox.options import Options

current_date = datetime.now()

# Format the date as required (month and filename)
year = current_date.strftime("%Y")
month = current_date.strftime("%b").upper()
day = current_date.strftime("%d")
filename_date = current_date.strftime("%d%b%Y").upper()

# Create the new URL
# url = f"https://archives.nseindia.com/content/historical/EQUITIES/{year}/{month}/cm{day}{month}{year}bhav.csv.zip"
file_name = "fo28JUL2023bhav.csv"
url =f"https://archives.nseindia.com/content/historical/DERIVATIVES/2023/JUL/{file_name}.zip"
print(url)

# Enable headless mode in Selenium
options = Options()
options.add_argument('--headless')


# Create the Firefox webdriver with the specified options and capabilities
browser = webdriver.Firefox()

#browser.implicitly_wait(20)
browser.get(url)
# At this point, the file should be downloaded.

# Close the browser
browser.quit()
print("browser closed")
# Perform file extraction
extract_zip(file_name)
