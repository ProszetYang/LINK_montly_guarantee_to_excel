from robocorp.tasks import task
from robocorp import browser
import json 
import time
from robocorp import vault

@task
def minimal_task():
    browser.configure(
        slowmo = 1000
    )
    open_sompo()
    download_monthly_guarantee()

    time.sleep(5)

def open_sompo():

    loginflag = 0
    while(loginflag == 0):
        loginInfo = vault.get_secret("Sompo")
        browser.goto("https://www.pw-japan.com/pwj/operation/auth/login")
        page = browser.page()
        page.locator("#memberId").fill(loginInfo["id"])
        page.locator("#memberPw").fill(loginInfo["password"])
        page.locator("#loginButton").click()
        loginflag = page.get_by_label("保証書：購入明細の通り").count()

def download_monthly_guarantee():
    page = browser.page()
    page.locator("#navi05").click()

    # Start waiting for the download
    with page.expect_download() as download_info:
        # Perform the action that initiates download
        page.get_by_alt_text("印刷").click()
    download = download_info.value
    download.save_as("output/" + download.suggested_filename)