from robocorp.tasks import task
from robocorp import browser
from robocorp.workitems import ApplicationException

from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.Tables import Tables
import os
import shutil


@task
def Course2():
    browser.configure(  
        slowmo = 100,
    )
    """Creates Robots from RobotSpareBin Industries limited"""
    """Recieves order HTML as PDF file"""
    """Saves the screenshots of the made robot"""
    """Embed the screenshots with the pdf"""
    """Create a zip file of the pdfs"""

    order_robots_from_spare_bins()
    get_orders()
    close_annoying_modal()
    # archive_receipts()
    # save_html_as_pdf()

def order_robots_from_spare_bins():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def get_orders():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv",overwrite = True)
    csv_file="orders.csv"
    worksheet = Tables().read_table_from_csv(csv_file,columns=['Order number','Head','Body','Legs','Address'])
    final_path = "output/Receipt"
    for row in worksheet:
        close_annoying_modal()
        fill_form(row)
        order_click()
        # screenshot of the robot preview
        print("The order is:",row['Order number'])
        screenshot_path = screenshot_preview(row['Order number'])
        receipt_path = receipt_save(row['Order number'])
        embed_receipt_to_screenshot(row['Order number'],screenshot_path,receipt_path)
        order_another()
    archive_receipts(final_path)
    

def close_annoying_modal():
    page = browser.page()
    page.get_by_role("button",name="OK").click()

def fill_form(df):
    page = browser.page()
    page.select_option('#head',df["Head"])
    page.check(f"input[type='radio'][value='{df["Body"]}']")
    page.fill("//label[text()='3. Legs:']/following-sibling::input",df['Legs'])
    page.fill("#address",df['Address'])
    page.click("text=Preview")
    


def screenshot_preview(order_number):
    screenshot_path = f"output/Receipt/order {order_number}/{order_number}.png"
    page = browser.page()
    page.locator("//div[@id='robot-preview-image']").screenshot(path=screenshot_path)
    return screenshot_path

def order_click():
    page = browser.page()
    count = 0
    while not page.locator("//div[@id='order-completion']").is_visible() and count < 5 and page.locator("//button[@id='order']").is_visible():
        count += 1
        page.locator("//button[@id='order']").click()
    if(count ==0 or count >= 5 and not page.locator("//div[@id='order-completion']").is_visible()):
        raise ApplicationException("An unexpected error occurred while clicking order")

def order_another():
    page = browser.page()
    page.click("//button[@id='order-another']")

def receipt_save(order_number):
    receipt_path = f"output/Receipt/order {order_number}/{order_number}.pdf"
    pdf = PDF()
    page = browser.page()
    receipt_html = page.locator("//div[@id='receipt']").inner_html()
    pdf.html_to_pdf(receipt_html,receipt_path)
    return receipt_path

def embed_receipt_to_screenshot(order_number,screenshot_path,receipt_path):
    final_address = f"output/Receipt/order {order_number}/final {order_number}.pdf"
    pdf = PDF()
    pdf.add_files_to_pdf(
        files=[receipt_path,screenshot_path],
        target_document = final_address,
        append = True
    )
    os.remove(screenshot_path)
    os.remove(receipt_path)

def archive_receipts(path):
    shutil.make_archive(path,'zip',path)
    shutil.rmtree(path)





