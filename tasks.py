from time import sleep
from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
@task
def order_robots_from_RobotSpareBin():
    """Orders robots ofrom RobotSpareBin Indsutris Inc.
    Saves the order HTML receipt as PDF file.
    Saves the screenshot of the ordered bot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    download_excel_report()
    open_robot_order_website()

    orders = download_excel_report()
    for order in orders:
        remove_popups()
        loop_orders(order)
        preview_bot()
        submit_order()
        pdf_path = store_receipt_as_pdf(order["Order number"])
        screenshot_path = screenshot_robot(order["Order number"])
        embed_screenshot_to_recepit(screenshot_path, pdf_path)
        order_another_robot()
    create_zip_archive()

def open_robot_order_website():
    """Opens the RobotSpareBin website."""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_excel_report():
    """Downlaods the report from the RobotSpareBin website."""
    filepath = "output\orders.csv"
    http = HTTP()
    http.download("https://robotsparebinindustries.com/orders.csv", target_file=filepath, overwrite=True)
    return read_csv_file(filepath)

def read_csv_file(filepath):
    """Reads the CSV file and returns the data as a list of dictionaries."""
    table = Tables()
    orders = table.read_table_from_csv(filepath, columns=["Order number","Head","Body","Legs","Address"])
    return orders

def remove_popups():
    """Removes the popups from the website."""
    page = browser.page()
    page.click("text=OK")
    
def loop_orders(order):
    """Processes the order in the """
    page = browser.page()
    page.select_option("id=head", order["Head"])
    page.click(f"id=id-body-{order['Body']}")
    page.fill("xpath=/html/body/div/div/div[1]/div/div[1]/form/div[3]/input",order["Legs"])
    page.fill("xpath=/html/body/div/div/div[1]/div/div[1]/form/div[4]/input",order["Address"])

        

def preview_bot():
    """Previews the bot and takes a screenshot."""
    page = browser.page()
    page.click("text=Preview")
    page.screenshot(path="output/robot_preview.png")

def submit_order():
    """Submits the order."""
    page = browser.page()
    page.click("id=order")
    if page.is_visible("id=receipt"):
        print()
    else:
        page.click("id=order")
    
def store_receipt_as_pdf(order_number):
    """Stores the order receipt as PDF."""
    pdffilepath = f"output/receipt/order_receipt_{order_number}.pdf"
    page = browser.page()
    order_submission_html = page.locator("id=receipt").inner_html()

    pdf = PDF()
    pdf.html_to_pdf(order_submission_html, pdffilepath)
    return pdffilepath


def screenshot_robot(order_number):
    """Takes a screenshot of the ordered robot."""
    screenshot_path ="output/screenshot/robot_{order_number}.png"
    page = browser.page()
    page.screenshot(path=screenshot_path)
    return screenshot_path

def embed_screenshot_to_recepit(screenshot, pdf_file):
    list_of_files = [
        f'{pdf_file}',  
        f'{screenshot}:align=center',
    ]
    pdf = PDF()
    pdf.add_files_to_pdf(files = list_of_files, target_document = pdf_file, append=True)

def order_another_robot():
    """Orders another robot."""
    page = browser.page()
    page.click("id=order-another")

def create_zip_archive():
    """Creates a ZIP archive of the receipts and the images."""
    archive = Archive()
    archive.archive_folder_with_zip("output/receipt", "output/receipts.zip")