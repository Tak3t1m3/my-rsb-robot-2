from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    open_robot_order_website()
    collect_and_submit_orders()
    archive_receipts()

def open_robot_order_website():
    browser.goto('https://robotsparebinindustries.com/#/robot-order')
    
    
def close_annoying_modal():
    page = browser.page()
    page.click("text=OK")

def collect_and_submit_orders():
    orders = get_orders()

    for row in orders:
        fill_the_form(row)

def fill_the_form(order):
    page = browser.page()
    if page.is_visible('.modal'):
        close_annoying_modal()
    page.select_option("#head", value=order["Head"])
    page.click("#id-body-"+str(order["Body"]))
    page.fill("//input[@placeholder='Enter the part number for the legs']", value=order["Legs"])
    page.fill("#address", value=order["Address"])
    page.click("#order")
    while page.is_visible('.alert.alert-danger[role="alert"]'):
        page.click("#order")
    pdf_file = store_receipt_as_pdf(order["Order number"])
    screenshot = screenshot_robot(order["Order number"])
    embed_screenshot_to_receipt(screenshot, pdf_file)

    page.click("#order-another",click_count=2)

def get_orders():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv",overwrite=True)

    excel = Tables()
    orders = excel.read_table_from_csv("orders.csv",header=True)
    return orders

def read_model_info_table():
    page = browser.page()
    model_table = page.locator("model-info").inner_html()
    return model_table

def store_receipt_as_pdf(order_number):
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()

    pdf = PDF()
    pdf.html_to_pdf(receipt_html, f"output/receipt/robot-{order_number}-receipt.pdf")
    return f"output/receipt/robot-{order_number}-receipt.pdf"


 

def screenshot_robot(order_number):
    page = browser.page()
    robot = page.locator("#robot-preview-image")
    robot.screenshot(path=f"output/images/robot-{order_number}-preview.png")
    return f"output/images/robot-{order_number}-preview.png"

def embed_screenshot_to_receipt(screenshot, pdf_file):
    list_of_files =[
        screenshot+": align=center",
    ]
    pdf = PDF()
    pdf.add_files_to_pdf(files=list_of_files,target_document=pdf_file,append=True)


def archive_receipts():
    arc = Archive()

    arc.archive_folder_with_zip("output/receipt",archive_name="output/receipts.zip")

