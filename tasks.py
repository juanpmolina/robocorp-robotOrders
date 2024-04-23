from robocorp.tasks import task
from robocorp import browser, vault

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF

import io
from PIL import Image

import shutil
import os

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
        slowmo=500,
    )
    log_in()
    get_orders()
    open_the_RobotOrders_website()
    fill_form_with_excel_data()
    zipFolder("output/receipts","output/receipts")
    shutil.rmtree("output/receiptsImgs")
    shutil.rmtree("output/receipts")
    log_out()

secrets = vault.get_secret('MariasVault')

def get_orders():
    """Downloads the orders.csv file given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def log_in():
    """Fills in the login form and clicks the 'Log in' button"""
    browser.goto("https://robotsparebinindustries.com/#/")

    page = browser.page()
    page.fill("#username", secrets['user'])
    page.fill("#password", secrets['pass'])
    page.click("button:text('Log in')")

def open_the_RobotOrders_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def fill_form_with_excel_data():
    """Read data from excel and fill in the sales form"""
    ordersTable = Tables().read_table_from_csv("orders.csv", header=True)

    if os.path.isdir("output/receiptsImgs"): 
        shutil.rmtree("output/receiptsImgs")
    if os.path.isdir("output/receipts"): 
        shutil.rmtree("output/receipts") 

    cont = 0
    for row in ordersTable:
        fill_and_submit_robotOrders_form(row)
        export_as_pdf(cont)
        cont = cont + 1

def fill_and_submit_robotOrders_form(robot_order_row):
    """Fills in the sales data and click the 'Submit' button for all the orders"""
    page = browser.page()

    page.click("text=OK")
    page.select_option("#head", robot_order_row["Head"])
    bodyID = "#id-body-" + str(robot_order_row["Body"]) + ""
    page.click(bodyID)
    page.fill("//input[@placeholder='Enter the part number for the legs']", robot_order_row["Legs"])
    page.fill("#address", robot_order_row["Address"])
    page.click("#preview")
    page.click("#order")
    errorExist = True
    while errorExist == True:
        errorExist = check_exists_by_xpath("@class='alert-danger'")
        if errorExist == True:
                page.click("#order")     

def export_as_pdf(cont):
    """Export the data to a pdf file"""
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdfPath = "output/receipts/receipt-" + str(cont) + "_html.pdf"
    pdf.html_to_pdf(receipt_html, pdfPath)
    
    imgPath = "output/receiptsImgs/robotCapture" + str(cont) + ".png"
    image_binary = page.locator("#robot-preview-image").screenshot()
    img = Image.open(io.BytesIO(image_binary))
    if not os.path.isdir("output/receiptsImgs"): 
        os.makedirs("output/receiptsImgs") 
    img.save(imgPath)
    pdf.add_watermark_image_to_pdf(image_path=imgPath,source_path=pdfPath,output_path=pdfPath)
 
    page.click("#order-another")

def check_exists_by_xpath(xpath):
    """Check if the error alert is visible"""
    page = browser.page()
    errorExist = page.is_visible(".alert.alert-danger[role='alert']")
    return errorExist

def zipFolder(surcePath,destinationPath):
    """zip folder from sourcePath to destinationPath"""
    shutil.make_archive(destinationPath, 'zip', surcePath)

def deleteFolder(folderPath):
    """delete folder in path"""
    shutil.rmtree(folderPath)

def log_out():
    """Presses the 'Log out' button"""
    page = browser.page()
    page.click("text=OK")
    page.click("text=Log out")