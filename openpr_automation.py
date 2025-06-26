import streamlit as st
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time

# Function to setup and run Selenium WebDriver
def run_selenium_automation():
    chromedriver_path = ChromeDriverManager().install()
    options = Options()
    #options.add_argument("--headless")  # Run headlessly for Streamlit deployment
    options.add_argument("--start-maximized")

    # Initialize WebDriver
    cService = Service(executable_path=chromedriver_path)
    driver = webdriver.Chrome(service=cService, options=options)
    
    # Visit OpenPR
    driver.get('https://www.openpr.com/')
    try:
        reject = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="cmpbntnotxt"]'))
        )
        reject.click()  # Reject cookies if the button is found
        st.write("Rejected cookies")
    except:
        st.write("Reject cookies button not found, proceeding to next step")
    
    # Click Submit
    submit = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="navbarText"]/ul/li[3]/a'))
    )
    submit.click()

    # Fill form (simplified for demonstration)
    input_box = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="code"]'))
    )
    input_box.clear()
    input_box.send_keys("6HA-2025-M6K439") 
    
    submit2 = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="main"]/div/div/div[5]/div/form/button'))
    )
    submit2.click()

    # Fill out form details
    name = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="formular"]/div[2]/div[1]/div/input'))
    )
    name.clear()
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

    # Select Category
    Category_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="formular"]/div[2]/div[6]/div/select'))
    )
    Select(Category_element).select_by_visible_text("Arts & Culture")

    # Filling the form details
    title = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="formular"]/div[2]/div[7]/div/input'))
    )
    title.clear()
    title.send_keys("Test title")

    text = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="inhalt"]'))
    )
    text.clear()

    # Add multiline text with newline characters
    multiline_text = """Global Smartphone Market Size, Trends, Growth, and Forecast Analysis 2025-2032

    The smartphone industry remains a pivotal sector driving digital transformation worldwide, characterized by rapid technological advancements and evolving consumer preferences. The continuous integration of 5G, AI capabilities, and enhanced camera systems fuel an intensely competitive landscape marked by swift innovation cycles and dynamic business growth.

    ### Market Size and Overview
    The Global Smartphone Market size is estimated to be valued at USD 520 billion in 2025 and is expected to reach USD 740 billion by 2032, exhibiting a compound annual growth rate (CAGR) of 5.1% from 2025 to 2032.

    ### Key Takeaways
    - **North America:** Mature smartphone ecosystem with a focus on 5G adoption and high-end device penetration.
    - **Latin America:** Increasing smartphone penetration backed by rising internet connectivity and affordable device options.
    - **Europe:** Strong inclination toward mid to premium smartphone segments fueled by growing enterprise adoption.
    """
    text.send_keys(multiline_text)

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

# Streamlit UI
st.title('Open PR Automation with Selenium')
st.write("This application automates the process of publishing articles on Open PR.")

# Input fields for user interaction
article_code = st.text_input("Enter Article Code", "6HA-2025-M6K439")
article_title = st.text_input("Enter Article Title", "Test title")

# Trigger the automation process when the user clicks the button
if st.button('Publish Article'):
    st.write("Starting automation process...")
    run_selenium_automation()
    st.write("Article published successfully!")

# Add a reset button if needed
if st.button('Reset'):
    st.write("Resetting the form.")
