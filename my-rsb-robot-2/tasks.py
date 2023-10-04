import os
from robocorp.tasks import task
from Browser import Browser, utils, SelectAttribute
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

browser = Browser()
http = HTTP()
csv = Tables()
pdf = PDF()
lib = Archive()
screen_bot_path = "C:\\Users\\nconte.ATG\\Desktop\\Progetti\\Robocorp\\my-rsb-robot-2\\output\\screenshots\\"


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
    orders = get_orders()
    process_orders(orders)
    archive_receipts()


def open_robot_order_website():
    """Open robot order website"""
    browser.new_browser(
        browser=utils.data_types.SupportedBrowsers.chromium)
    browser.new_page()
    browser.go_to("https://robotsparebinindustries.com/#/robot-order")
    close_annoying_modal()


def get_orders():
    """Downloads Orders csv file from the given URL"""
    http = HTTP()
    http.download(
        url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

    return csv.read_table_from_csv("orders.csv", header=True)


def process_orders(orders):
    for order in orders:
        fill_the_form(order)
        close_annoying_modal()


def fill_the_form(order):
    """Fill the robot form"""
    browser.select_options_by(
        "#head", SelectAttribute['value'], str(order["Head"]))
    browser.check_checkbox(
        'input[name="body"][value="' + str(order["Body"]) + '"]')
    browser.fill_text(
        'input[placeholder="Enter the part number for the legs"]', str(order["Legs"]))
    browser.fill_text("#address", order["Address"])

    exit = False

    while not (exit):
        browser.click("#order")

        try:
            browser.get_property(
                selector="#order", property="outerHTML")
        except:
            exit = True

    store_receipt_as_pdf(order["Order number"])
    browser.click("#order-another")


def close_annoying_modal():
    """Close annoying modal that opens up at page load"""
    browser.click("button.btn-warning:text('Yep')")


def screenshot_robot(order_number):
    """Take a screenshot of the receipt"""
    file_name = order_number + ".png"

    file_screen = os.path.join(
        screen_bot_path, file_name)
    browser.take_screenshot(
        filename=file_screen, selector="#robot-preview")

    return file_screen


def store_receipt_as_pdf(order_number):
    """Export the receipt to a pdf file with the screenshot"""
    receipt_html = browser.get_property(
        selector="#order-completion", property="outerHTML")
    pdf.html_to_pdf(receipt_html, "output/receipts/" + order_number + ".pdf")

    file_screenshot = screenshot_robot(order_number)

    list_of_files = [
        "output/receipts/" + order_number + ".pdf", file_screenshot]

    pdf.add_files_to_pdf(
        files=list_of_files,
        target_document="output/receipts_merged/" + order_number + ".pdf"
    )


def archive_receipts():
    lib.archive_folder_with_zip(
        folder="output/receipts_merged", archive_name="receipts.zip")
