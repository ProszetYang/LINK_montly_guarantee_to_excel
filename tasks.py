from robocorp.tasks import task
from robocorp import browser
from RPA.PDF import PDF
import time
from robocorp import vault
from RPA.Excel.Files import Files
import fitz
import re
from datetime import datetime
from datetime import timedelta

HEADER = [[
    'NO.',
    '登録日',
    '保証番号',
    'プラン',
    '年数',
    'お客様名',
    '保証付与商品',
    '商品型番',
    '商品金額',
    '保証料金',
    '手数料',
    '当社ご請求金額'
]]


@task
def monthly_guarantee_to_excel():
    resultfile = "//192.168.0.178/contents/WinActor/シナリオ/価格チェックチーム/延長保証月次請求書Excel化/作業リスト/" + str(datetime.now().strftime("%Y")) + "年" + str((datetime.now() - timedelta(days=28)).strftime("%m")) + "月延長保証月次請求書.xlsx"

    browser.configure(
        headless=False
    )
    open_sompo()
    data_pdf = download_monthly_guarantee()
    pdf_to_excel(data_pdf, resultfile)

def open_sompo():

    loginflag = 0
    loginInfo = vault.get_secret("Sompo")
    browser.goto("https://www.pw-japan.com/pwj/operation/auth/login")
    page = browser.page()
    while(loginflag == 0):
        page.locator("#memberId").fill(loginInfo["id"])
        page.locator("#memberPw").fill(loginInfo["password"])
        page.locator("#loginButton").click()
        time.sleep(3)
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


def pdf_to_excel(pdf_path, excel_path):
    lib = Files()
    lib.create_workbook(sheet_name="リスト")
    lib.append_rows_to_worksheet(HEADER)
    doc = fitz.open(pdf_path)
    for page in doc:
        tabs = page.find_tables()
        if tabs.tables:
            for row in tabs[1].extract():
                if row[0].isdigit():
                    row[0] = int(row[0])
                    row[1] = datetime.strptime(row[1], '%y/%m/%d').strftime("%Y/%m/%d")
                    row[8] = int(re.sub(r'[^-\d]', "", row[8]))
                    row[9] = int(re.sub(r'[^-\d]', "", row[9]))
                    row[10] = int(re.sub(r'[^-\d]', "", row[10]))
                    row[11] = int(re.sub(r'[^-\d]', "", row[11]))
                    lib.append_rows_to_worksheet([row])
    lib.save_workbook(excel_path)