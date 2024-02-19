import os
import random
import time
import warnings

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.chrome import service as fs
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

warnings.simplefilter(action='ignore', category=UserWarning)

CHROMEDRIVER = 'chromedriver'

options = Options()
options.add_argument(
    "--user-data-dir=" + os.environ['USERPROFILE'] + "/AppData/Local/Google/Chrome/User Data")


def main():
    chrome_service = fs.Service(executable_path=CHROMEDRIVER)
    driver = webdriver.Chrome(service=chrome_service, options=options)
    driver.get('https://teams.microsoft.com/l/message/19:gIK70q6P0m7vS4d9aTC-VR7WVb_yMqH_UJNFaeEgoHk1@thread.tacv2/1707981658150?groupId=0ddc3811-e676-47dd-b8ba-c7eb0e4e3ff8&parentMessageId=1707981658150&tenantId=72fe835d-5e95-4512-8ae0-a7b38af25fc8')

    wait = WebDriverWait(driver, 5)

    button = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@id="openTeamsClientInBrowser"]')))
    driver.execute_script("arguments[0].click();", button)

    parent = wait.until(EC.presence_of_element_located((By.XPATH, '//tbody')))
    rows = parent.find_elements(By.XPATH, './tr')

    # url_elemsは、指定されたXPathにマッチする<a>タグのリストを保持しています。
    # これらの要素からhref属性を取得してURLのリストを作成します。

    urls = []  # URLを格納するための空のリストを初期化

    for row in rows:
        # 教科名が「留学生」を含まない行を検索
        if "留学生" not in row.find_element(By.XPATH, "./td[2]").text:
            # 条件にマッチした行の中の<a>タグからhref属性を取得
            url = row.find_element(By.XPATH, "./td[3]/a").get_attribute('href')
            urls.append(url)

    for url in urls:
        driver.get(url)

        try:
            parent = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@id="question-list"]')))
            questions = parent.find_elements(By.XPATH, './div[@data-automation-id="questionItem"]')

            for question in questions:
                rand = random.randint(0, 1)

                try:
                    # ここで指定したXPathに対応する要素を検索します。
                    # f文字列を使って、rand変数の値に応じてXPathを動的に変更します。
                    input = question.find_element(By.XPATH, f'(.//input)[last()-{rand}]')

                    # 要素が見つかった場合の処理をここに書きます。
                    driver.execute_script("arguments[0].click();", input)
                    time.sleep(0.1)

                except NoSuchElementException:
                    # NoSuchElementExceptionが発生した場合、ここに書かれた処理を実行します。
                    # この場合、処理をスキップするため、passを使用します。
                    pass

            # フォームの送信ボタンをクリックする
            button = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@data-automation-id="submitButton"]')))
            driver.execute_script("arguments[0].click();", button)
            time.sleep(1)

        except TimeoutException:
            # TimeoutExceptionが発生した場合、ここに書かれた処理を実行します。
            # この場合、処理をスキップするため、passを使用します。
            pass


    driver.quit()


if __name__ == '__main__':
    main()
