from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.Dialogs import Dialogs
import pandas as pd
from time import sleep
from RPA.Archive import Archive

browser = Selenium(auto_close=False)
r = HTTP()
pdf_file = PDF()
lib = Archive()

def login():
    browser.open_chrome_browser(url='https://robotsparebinindustries.com/#/robot-order')
    browser.maximize_browser_window()
    browser.wait_until_element_is_visible(locator='//*[@class="modal-header"]')
    browser.click_button(locator='//*[@class="btn btn-dark"]')
    r.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def create_receipt(order):
    receipt = browser.get_element_attribute(locator='//*[@id="receipt"]', attribute='innerHTML')
    pdf_file.html_to_pdf(receipt, "output/pdf/"+order+".pdf")
    browser.wait_until_element_is_visible(locator='//*[@id="order-another"]')
    pdf_file.add_watermark_image_to_pdf(image_path=f"output/screenshot/{order}.png",source_path=f"output/pdf/{order}.pdf",output_path=f"output/receipts/{order}.pdf")

def fill_data(head,body,legs,address,order):
    browser.click_element(locator='//*[@id="head"]')
    browser.click_element(locator='//*[@value="{}"]'.format(head))
    browser.click_element(locator='//*[@for="id-body-{}"]'.format(body))
    browser.input_text(locator='//*[@placeholder="Enter the part number for the legs"]',text=legs) 
    browser.input_text(locator='//*[@id="address"]', text=address)
    browser.click_button(locator='//*[@id="preview"]')
    sleep(3)
    browser.wait_until_element_is_visible(locator='//*[@id="robot-preview-image"]')
    sleep(1)
    browser.click_button(locator='//*[@id="order"]')
    browser.wait_until_element_is_visible(locator='//*[@id="receipt"]',timeout='15')
    browser.wait_until_element_is_visible(locator='//*[@id="order-another"]')
    sleep(3)
    browser.screenshot(locator='//*[@id="robot-preview-image"]',filename= "output/screenshot/"+order+".png")
    create_receipt(order)
    browser.click_button_when_visible(locator='//*[@id="order-another"]')


def skip_modals():
    browser.wait_until_element_is_visible(locator='//*[@class="modal-header"]')
    browser.click_button(locator='//*[@class="btn btn-dark"]')

def buy_robots():
    df = pd.read_csv('orders.csv')
    df = df.astype(str)
    # print(df.head())
    for i in range(len(df)):
        try:
            skip_modals()
        except:
            pass    
        order = df['Order number'][i]
        head = df['Head'][i]
        body = df['Body'][i]
        legs = df['Legs'][i]
        address = df['Address'][i]
        fill_data(head,body,legs,address,order)


def zip_pdf():
    lib.archive_folder_with_zip(folder="output/receipts",archive_name="output/receipts.zip", recursive=True)
    

def main():
    try:
        login()
        buy_robots()
        zip_pdf()
        # fill_data(1,2,3,'address',"1")
    finally:
        browser.close_browser()


if __name__ == "__main__":
    main()
