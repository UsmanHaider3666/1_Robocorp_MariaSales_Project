from RPA.Browser.Selenium import Selenium
import time
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
import os
from RPA.Email.ImapSmtp import ImapSmtp
from RPA.Archive import Archive



class RpaRobocorp1:

    def __init__(self):
        self.mail = None
        self.browse = Selenium()
        self.url = "https://robotsparebinindustries.com/"
        self.http = HTTP()
        self.excel = Files()
        self.pdf = PDF()
        self.email = ImapSmtp()
        self.zip = Archive()


    def open_browser(self):
        self.browse.open_available_browser(self.url, maximized=True, )

    def log_in(self):
        self.browse.input_text('//input[@id="username"]', "maria")
        self.browse.input_text('//input[@id="password"]', "thoushallnotpass")
        self.browse.click_button_when_visible('//*[@type="submit"]')
        time.sleep(3)
        self.browse.auto_close = False

    def download_the_excel_file(self):
        self.http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)

    def fill_submit_form_for_person(self):
        self.excel.open_workbook("SalesData.xlsx")
        sales_data = self.excel.read_worksheet_as_table(header=True)
        self.excel.close_workbook()
        for sales in sales_data:
            self.browse.input_text('//*[@id="firstname"]', sales["First Name"])
            self.browse.input_text('//*[@id="lastname"]', sales["Last Name"])
            self.browse.select_from_list_by_value('//*[@id="salestarget"]', str(sales["Sales Target"]))
            self.browse.input_text('//*[@id="salesresult"]', str(sales["Sales"]))
            self.browse.click_button_when_visible('//*[@id="sales-form"]/button')
            time.sleep(3)
            break
        self.browse.click_button_when_visible('//*[@id="root"]/div/div/div/div[2]/div[3]/button[1]')

    def fill_the_form_using_the_data_from_the_excel_file(self):
        self.excel.open_workbook("SalesData.xlsx")
        sales_data = self.excel.read_worksheet_as_table(header=True)
        self.excel.close_workbook()
        for sales in sales_data:
            self.browse.input_text('//*[@id="firstname"]', sales["First Name"])
            self.browse.input_text('//*[@id="lastname"]', sales["Last Name"])
            self.browse.select_from_list_by_value('//*[@id="salestarget"]', str(sales["Sales Target"]))
            self.browse.input_text('//*[@id="salesresult"]', str(sales["Sales"]))
            self.browse.click_button_when_visible('//*[@id="sales-form"]/button')
            time.sleep(3)

        self.browse.click_button_when_visible('//*[@id="root"]/div/div/div/div[2]/div[3]/button[1]')

    def take_screenshot(self):
        self.browse.capture_element_screenshot('//*[@id="root"]/div/div/div/div[2]/div[1]',
                                               filename=f"{os.getcwd()}/output/sales_summary.png")

    def table_to_pdf(self):
        sales_results_html = self.browse.get_element_attribute('//*[@id="root"]/div/div/div/div[2]',
                                                               attribute="outerHTML")

        self.pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")

    def log_out(self):
        self.browse.click_button_when_visible('//*[@id="logout"]')

    def close_browser(self):
        self.browse.auto_close = True

    def make_zip(self):
        self.zip.archive_folder_with_zip("/home/usman/Python-RPA/1_Robocorp_MariaSales_Project/output", "final.zip", True, )

    def send_email(self):
        try:
            self.email.gmail_account = "usmanhaiderpk02@gmail.com"
            self.email.gmail_password = "davidnorteen!1122"
            self.email.sender = self.email.gmail_account
            self.email.result = ["norteendavid@gmail.com", "hdrmastoi@gmail.com"]

            self.mail = ImapSmtp(smtp_server="smtp.gmail.com", smtp_port=587)
            self.mail.authorize(account=self.email.gmail_account, password=self.email.gmail_password,
                                smtp_server="smtp.gmail.com")

            self.mail.send_message(
                sender=self.email.gmail_account,
                subject="Message from RPA Python",
                recipients=self.email.result,
                body="This is final result ",
                attachments=["/home/usman/Python-RPA/1_Robocorp_MariaSales_Project/final.zip"]
            )
            print("mail sent")
        except:
            print('Sending Error')


if __name__ == '__main__':
    res = RpaRobocorp1()
    res.open_browser()
    res.log_in()
    res.download_the_excel_file()
    # res.fill_submit_form_for_person()
    res.fill_the_form_using_the_data_from_the_excel_file()
    res.take_screenshot()
    res.table_to_pdf()
    res.log_out()
    res.close_browser()
    res.make_zip()
    res.send_email()
