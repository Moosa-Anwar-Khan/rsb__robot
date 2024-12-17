import os
import logging

from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF

# Creating a log file
logging.basicConfig(filename="bot_log.log", level=logging.INFO, format="%(asctime)s - %(message)s")

# MAIN TASK
@task
def robot_spare_bin_python():
    """Robot main workflow"""

    USERNAME, PASSWORD = load_credentials('credentials.txt')

    try:
        initialization(USERNAME, PASSWORD)
        download_excel_file()
        fill_form_with_excel_data() 
        collect_results() 
        export_as_pdf()
    except Exception as e:
        logging.error(f"Unexpected error occurred: {e}")
    finally:
        log_out()

def load_credentials(file_path):
    """Load credentials from a file."""
    with open(file_path, 'r') as f:
        lines = f.readlines()
        user = lines[0].strip().split('=')[1]  # Extracting username
        password = lines[1].strip().split('=')[1]  # Extracting password
    return user, password        


def initialization(username, password):
    """Initialization State: Setup browser and validate credentials"""
    logging.info("Initializing resources and browser.")
    if not username or not password:
        raise ValueError("Credentials not found. Ensure environment variables are set.")
    
    browser.configure(slowmo=100)
    open_the_intranet_website()
    log_in(username, password)

def open_the_intranet_website():
    """Navigates to the intranet website."""
    logging.info("Opening intranet website.")
    browser.goto("https://robotsparebinindustries.com/")

def log_in(username, password):
    """Logs in using secure credentials."""
    try:
        logging.info("Logging in to the website.")
        page = browser.page()
        page.fill("#username", username)
        page.fill("#password", password)
        page.click("button:text('Log in')")
    except Exception as e:
        logging.error(f"Error during login: {e}")
        raise

def download_excel_file():
    """Downloads excel file from the given URL"""
    try:
        logging.info("Downloading Excel file.")
        http = HTTP()
        http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)

    except Exception as e:
        logging.error(f"Error while downloading excel file: {e}")
        raise     

def fill_form_with_excel_data():
    """Read data from excel and fill in the sales form"""
    try:
        excel = Files()
        excel.open_workbook("SalesData.xlsx")
        worksheet = excel.read_worksheet_as_table("data", header=True)
        excel.close_workbook() 

        for ind, row in enumerate(worksheet):
           logging.info(f"Processing item {ind + 1}: {row}")
           fill_and_submit_sales_form(row)  

    except Exception as e:
        logging.error(f"Error while filling form: {e}")
        raise        


def fill_and_submit_sales_form(sales_rep):
    """Fills in the sales form and submits it."""
    page = browser.page()
    page.fill("#firstname", sales_rep["First Name"])
    page.fill("#lastname", sales_rep["Last Name"])
    page.select_option("#salestarget", str(sales_rep["Sales Target"]))
    page.fill("#salesresult", str(sales_rep["Sales"]))
    page.click("text=Submit")

def collect_results():
    """Takes a screenshot of the final results."""
    try:
        logging.info("Collecting results: Taking screenshot.")
        page = browser.page()
        page.screenshot(path="output/sales_summary.png")
    except Exception as e:
        logging.error(f"Error during result collection: {e}")
        raise

def export_as_pdf():
    """Exports sales results to a PDF."""
    try:
        logging.info("Exporting results to PDF.")
        page = browser.page()
        sales_results_html = page.locator("#sales-results").inner_html()
        pdf = PDF()
        pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")
    except Exception as e:
        logging.error(f"Error exporting PDF: {e}")
        raise

def log_out():
    """End State: Logs out and closes browser."""
    try:
        logging.info("Logging out and closing browser.")
        page = browser.page()
        page.click("text=Log out")

    except Exception as e:
        logging.warning(f"Error during logging out: {e}")
