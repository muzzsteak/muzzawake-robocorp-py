from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Browser import Selenium

import shutil
from pathlib import Path

http_lib = HTTP()
tables_lib = Tables()

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(slowmo=1000,)
    open_robot_order_website()
    close_annoying_popup()
    download_orders_file()
    table = read_csv_as_table()
    get_orders(table)
    create_receipts_archive()

def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    # browser.goto("https://robotsparebinindustries.com")

def close_annoying_popup():
    page = browser.page()
    try:
        page.click("#root > div > div.modal > div > div > div > div > div > button.btn.btn-dark")
    except TimeoutError as e:
        print('annoying popup not found')

def download_orders_file():
    http_lib.download("https://robotsparebinindustries.com/orders.csv", None, True, False, True)

def read_csv_as_table():
    table = tables_lib.read_table_from_csv("orders.csv", True)
    return table

def get_orders(table):
    page = browser.page()

    # div #receipt
    # div #robot-preview-image
    # button #order-another

    for row in table:
        close_annoying_popup()
        fill_the_form(row)
        
        page.click('#order')
        if(page.query_selector('.alert-danger')):
            page.click('#order')

        create_pdf_receipt(row['Order Number'])

        if(page.locator('#order-another')):
            page.click('#order-another')
            
def fill_the_form(row):
    page = browser.page()
    
    page.select_option('#head', row['Head'])
    page.click(f"#id-body-{row['Body']}")
    page.get_by_placeholder('Enter the part number for the legs').fill(row['Legs'])
    page.fill('#address', row['Address'])

    page.click('#preview')

def create_pdf_receipt(orderNumber):
    page = browser.page()

    if(page.locator('#receipt')):
        receipt_html = page.locator('#receipt').inner_html()
        pdf = PDF()
        pdf.html_to_pdf(receipt_html, f"output/pdf_receipts/receipt for order num {orderNumber}.pdf")

        # embed screenshot of robot preview in pdf
        page.query_selector('#robot-preview-image').screenshot(path='output/robot_preview.jpg')
        pdf.add_files_to_pdf('output/robot_preview.jpg', 
                             f"output/pdf_receipts/receipt for order num {orderNumber}.pdf",
                             f"output/pdf_receipts/receipt for order num {orderNumber}.pdf")
        
def create_receipts_archive():
    source_dir = 'output/pdf_receipts'
    source_dir = Path(source_dir)
    shutil.make_archive('output/receipts_archive', 'zip', root_dir=source_dir)




