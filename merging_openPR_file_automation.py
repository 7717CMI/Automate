import os
import pandas as pd
import win32com.client
from docx import Document
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
import time

# Get today's date in the format YYYY/MM/DD
today = datetime.today()
folder_path = rf'C:\Users\akshat\Desktop\RPA\Files\{today.year}\{today.strftime("%m")}\{today.strftime("%d")}'
excel_path = r'C:\Users\akshat\Desktop\RPA\ROB.xlsx'

# Read the Excel file and get the 'Market Name' column
keywords_df = pd.read_excel(excel_path)
market_names = keywords_df['Market Name'].dropna().tolist()

def convert_doc_to_docx(doc_path, output_path=None):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(doc_path)
    if not output_path:
        output_path = os.path.splitext(doc_path)[0] + ".docx"
    doc.SaveAs(output_path, FileFormat=16)
    doc.Close()
    word.Quit()
    return output_path

def text_of_press_release(docx_path, start_index=21, end_index=-8):
    doc = Document(docx_path)
    saved = "\n".join([para.text for para in doc.paragraphs])
    words = saved.split()
    if len(words) > abs(end_index):
        chunk = " ".join(words[start_index:end_index])
        return chunk
    else:
        return "Text not found."

def run_selenium_automation(article_code, article_title, multiline_text):
    chromedriver_path = ChromeDriverManager().install()
    options = Options()
    # options.add_argument("--headless")
    options.add_argument("--start-maximized")
    cService = Service(executable_path=chromedriver_path)
    driver = webdriver.Chrome(service=cService, options=options)
    driver.get('https://www.openpr.com/')
    try:
        reject = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="cmpbntnotxt"]'))
        )
        reject.click()
    except:
        pass
    submit = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="navbarText"]/ul/li[3]/a'))
    )
    submit.click()
    input_box = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="code"]'))
    )
    input_box.clear()
    input_box.send_keys(article_code)
    # Try both possible CSS selectors for the submit button
    try:
        submit2 = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '#main > div > div > div:nth-child(5) > div > form > button'))
        )
        submit2.click()
    except:
        submit2 = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '#main > div > div > div:nth-child(6) > div > form > button'))
        )
        submit2.click()
    name = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="formular"]/div[2]/div[1]/div/input'))
    )
    name.send_keys("Vishwas tiwari")
    email = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="formular"]/div[2]/div[2]/div/input'))
    )
    email.clear()
    email.send_keys("vishwas@coherentmarketinsights.com")
    pr_agency = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="formular"]/div[2]/div[3]/div/input'))
    )
    pr_agency.clear()
    pr_agency.send_keys("Vishwas tiwari")
    number = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="formular"]/div[2]/div[4]/div/input'))
    )
    number.clear()
    number.send_keys("1234567890")
    ComName = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="archivnmfield"]'))
    )
    ComName.clear()
    ComName.send_keys("Coherenet Market Insights")
    s1 = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="popup-archiv"]/div/a[1]'))
    )
    s1.click()
    Category_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="formular"]/div[2]/div[6]/div/select'))
    )
    Select(Category_element).select_by_visible_text("Arts & Culture")
    title = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="formular"]/div[2]/div[7]/div/input'))
    )
    title.clear()
    title.send_keys(article_title)
    text = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="inhalt"]'))
    )
    text.clear()
    text.send_keys(multiline_text)
    # ... (rest of your form filling code here, unchanged) ...
    # Close the driver at the end
    #driver.quit()

    about = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="formular"]/div[2]/div[9]/div/textarea'))
    )
    about.clear()
    multi = """Contact Us:
    Mr. Shah
    Coherent Market Insights Pvt. Ltd,
    U.S.: + 12524771362
    U.K.: +442039578553
    AUS: +61-2-4786-0457
    INDIA: +91-848-285-0837
    ✉ Email: sales@coherentmarketinsights.com
    About Us:
    Coherent Market Insights leads into data and analytics, audience measurement, consumer behaviors, and market trend analysis. From shorter dispatch to in-depth insights, CMI has exceled in offering research, analytics, and consumer-focused shifts for nearly a decade. With cutting-edge syndicated tools and custom-made research services, we empower businesses to move in the direction of growth. We are multifunctional in our work scope and have 450+ seasoned consultants, analysts, and researchers across 26+ industries spread out in 32+ countries.")
    """
    about.send_keys(multi)
    address = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="formular"]/div[2]/div[10]/div/textarea'))
    )
    address.clear()
    address.send_keys("123 Test Street, Test City, Test Country")

    image = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="bild"]'))
    )
    image.clear()
    image.send_keys(r"C:\Users\akshat\Desktop\code\image.jpg")  # Replace with the path to your image

    caption = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="formular"]/div[2]/div[12]/div/input'))
    )
    caption.clear()
    caption.send_keys("This is a test caption for the image.")

    notes = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="formular"]/div[2]/div[13]/div/textarea'))
    )
    notes.clear()
    notes.send_keys("This is a test notes section for the press release submission.")

    # Agree to terms and conditions
    tick1 = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="input-agb"]'))
    )
    tick1.click()

    tick2 = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="input-ds"]'))
    )
    tick2.click()

    final = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="formular"]/div[2]/button'))
    )
    final.click()


# MAIN LOOP: For each article in today's folder, publish it
for market_name in market_names:
    doc_file = f"ROB_{market_name}.doc"
    doc_path = os.path.join(folder_path, doc_file)
    if os.path.exists(doc_path):
        docx_path = convert_doc_to_docx(doc_path)
        multiline_text = text_of_press_release(docx_path)
        # You can customize article_code and article_title as needed, e.g.:
        article_code = "6HA-2025-M6K439"  # Or fetch from Excel if available
        article_title = f"{market_name} Market Insights"
        print(f"Publishing article for: {market_name}")
        run_selenium_automation(article_code, article_title, multiline_text)
        print(f"Published: {market_name}")
        time.sleep(10)  # Optional: wait between articles
    else:
        print(f"File not found: {doc_file}")