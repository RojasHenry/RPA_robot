from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
from RPA.FileSystem import FileSystem

import uuid

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=600,
    )
    open_robot_order_website()
    download_excel_file()
    orders = get_orders()
    for row in orders:
        fill_the_form(row)
    archive_receipts()
    clean_receipts()

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_excel_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders(): 
    """Get orders from .csv file"""
    return Tables().read_table_from_csv("orders.csv", header=True)

def fill_the_form(order):
    """Submit orders in form"""
    page = browser.page()
    close_annoying_modal()
    page.select_option("#head",value=str(order["Head"]))
    body = order["Body"]
    page.click(f"#id-body-{str(body)}")
    page.get_by_placeholder("Enter the part number for the legs").fill(str(order["Legs"]))
    page.fill('#address',order["Address"])
    while(not page.get_by_text("Receipt").is_visible()):
        page.click('#order')
    order_number = f"Order-{uuid.uuid4()}"
    pdf_file = store_receipt_as_pdf(order_number)
    screenshot = screenshot_robot(order_number)
    embed_screenshot_to_receipt(screenshot,pdf_file)
    page.click('#order-another')
    

def close_annoying_modal():
    """Remove initial modal"""
    page = browser.page()
    page.click("text=OK")

def store_receipt_as_pdf(order_number):
    """Create pdf of each order"""
    page = browser.page()
    file_path = f"output/receipt/{order_number}/{order_number}.pdf"
    receipt = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(receipt, file_path)
    return file_path

def screenshot_robot(order_number):
    """Create screenshot of robot preview"""
    page = browser.page()
    screenshot_path = f"output/receipt/{order_number}/{order_number}.png"
    page.locator("#robot-preview-image").screenshot(path=f"{screenshot_path}")
    return screenshot_path

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Append screenshot to pdf file"""
    pdf = PDF()
    list_of_files = [
        f'{screenshot}:x=0,y=0',
    ]
    pdf.add_files_to_pdf(list_of_files,target_document=pdf_file,append=True)

def archive_receipts():
    """Generate ZipFile of receipts"""
    lib = Archive()
    lib.archive_folder_with_zip('output/receipt', 'output/receipts.zip', recursive=True)

def clean_receipts():
    """Clear results"""
    lib = FileSystem()
    lib.remove_directory('output/receipt', recursive=True)