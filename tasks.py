from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.Tables import Tables
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
    download_excel()
    data = read_excel()
    page = open_robot_order_website()
    for row in data:
        popup(page)
        fill_data(page, row)
        preview_robot(page)
        order_robot(page,row)
    archive_receipts()    
    
    
def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    page = browser.page()
    page.set_viewport_size(({"width": 1920, "height": 1080}))
    return page
    
    

    
def popup(page: browser.Page):
    page.wait_for_selector("button:text('OK')")
    page.click("button:text('OK')")
    
    
def download_excel():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    
def read_excel():
    table = Tables()
    data = table.read_table_from_csv("orders.csv")
    return data

def fill_data(page: browser.Page, row: dict):
    page.select_option("#head", row["Head"])
    page.click(f"id=id-body-{(row['Body'])}")
    page.fill(".form-control[type=number]",row["Legs"])
    page.fill("#address", row["Address"])


def preview_robot(page: browser.Page):
    page.click("button:text('Preview')")
    
def order_robot(page: browser.Page,row: dict):
    page.click("button:text('Order')")
    alert = page.query_selector(".alert.alert-danger")
    
    
    if alert is not None:
        while alert is not None:
            page.click("button:text('Order')")
            alert = page.query_selector(".alert.alert-danger")
    
    pdf_path = store_receipt_as_pdf(page,row["Order number"])
    screenshot_path = screenshot_robot(row["Order number"], page)
    embed_screenshot_to_receipt(screenshot_path, pdf_path)
    order_another = page.query_selector("id=order-another")
    if order_another is not None:
        page.click("id=order-another")    

def store_receipt_as_pdf(page: browser.Page,order_number: str) ->str:
    pdf = PDF()
    pdf_path = f"output/receipts/robot_receipts_{order_number}.pdf"    
    
    html = page.locator("#receipt").inner_html()
    pdf.html_to_pdf(html, pdf_path)
    return pdf_path
def embed_screenshot_to_receipt(screenshot_path, pdf_file):
    pdf = PDF()
    pdf.add_files_to_pdf([screenshot_path], pdf_file,append=True)
    
def screenshot_robot(order_number:str, page: browser.Page)-> str:
    screenshot_path = f"output/receipts/robot_{order_number}.png"
    robot_img = page.query_selector("#robot-preview-image")
    robot_img.screenshot(path=screenshot_path)
    return screenshot_path

def archive_receipts():
    zipper = Archive()
    zipper.archive_folder_with_zip("output/receipts", "output/robot_receipts.zip")