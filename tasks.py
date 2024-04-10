from robocorp.tasks import task
from robocorp import browser
from RPA.PDF import PDF
import time
from robocorp import vault
from RPA.Excel.Files import Files

@task
def monthly_guarantee_to_excel():
    #open_sompo()
    #data_pdf = download_monthly_guarantee()
    data_pdf = "output/hoshosho.pdf"
    pdf_to_excel(data_pdf, 'output/test.xlsx')
    time.sleep(5)

def open_sompo():

    loginflag = 0
    loginInfo = vault.get_secret("Sompo")
    browser.goto("https://www.pw-japan.com/pwj/operation/auth/login")
    page = browser.page()
    while(loginflag == 0):
        page.locator("#memberId").fill(loginInfo["id"])
        page.locator("#memberPw").fill(loginInfo["password"])
        page.locator("#loginButton").click()
        loginflag = page.get_by_label("保証書：購入明細の通り").count()
        if loginflag == 0:
            page.get_by_role("button", name="ログイン画面").click()

def download_monthly_guarantee():
    page = browser.page()
    page.locator("#navi05").click()

    # Start waiting for the download
    with page.expect_download() as download_info:
        # Perform the action that initiates download
        page.get_by_alt_text("印刷").click()
    download = download_info.value
    download.save_as("output/" + download.suggested_filename)
    return "output/" + download.suggested_filename


def pdf_to_excel(pdf_file_path, excel_file_path):


