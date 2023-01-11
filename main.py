from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import time
import os
from selenium.webdriver.common.by import By
import re
import requests
import xml.etree.ElementTree as ET
import json
import pandas as pd


def is_element_existed(driver, id):
    try:
        driver.find_element(By.ID, id)
        return True
    except:
        return False


def returnBI_address(pendo_src, hostname, cookies_fetch):

    cookies = {
        'CompanyAlias': '',
        'userCulture': '0',
        '_ga': '',
        '.ASPXANONYMOUS': '',
        'BIGipServern31.ultipro.com_https_pool': '',
        'DeviceIdentifier': '',
        'ASP.NET_SessionId': '',
        'UltiProNET': '',
        'loginToken': '',
    }
    headers = {
        'Accept': 'text/plain, */*; q=0.01',
        'Accept-Language': 'en-US,en;q=0.9',
        'Connection': 'keep-alive',
        'Origin': f'https://{hostname}.ultipro.com',
        'Referer': f'https://{hostname}.ultipro.com/pages/VIEW/BusinessIntelligenceLinks.aspx?pendopdk=cogno&pendoWfiHow=new-menu&{pendo_src}&destinationId=1932&USParams=PK=COGNO!MenuID=1932!PageRerId=1932!ParentRerId=2137',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
        'X-Requested-With': 'XMLHttpRequest',
        'sec-ch-ua': '"Not?A_Brand";v="8", "Chromium";v="108", "Google Chrome";v="108"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
    }
    data = {
        'url': f'https://c{hostname}.ultipro.com/ibmcognos/bi/v1/disp',
        'target': '_blank',
    }
    for cookie in cookies_fetch:
        name = cookie['name']
        value = cookie['value']
        cookies[name] = value
    response = requests.post(
        f'https://{hostname}.ultipro.com/services/CommonService.asmx/GetCognosSsoUrl?pendopdk=cogno&pendoWfiHow=new-menu&{pendo_src}&destinationId=1932&USParams=PK=COGNO!MenuID=1932!PageRerId=1932!ParentRerId=2137',
        cookies=cookies,
        headers=headers,
        data=data,
    )
    return response.text


def ukg_bi_cookies(hostname):

    service_object = Service(r"chromedriver_win32\chromedriver.exe")
    options_dvr = webdriver.ChromeOptions()
    options_dvr.add_argument('headless')
    options_dvr.add_argument("--incognito")
    driver = webdriver.Chrome(service=service_object, options=options_dvr)
    # hostname = input('Please enter your UKG host name: \n')
    url = f'https://{hostname}.ultipro.com/Login.aspx'
    driver.get(url)
    loginusername = input('Please Enter your UKG login name: \n')
    psw = input('please enter your UKG password: \n')
    driver.find_element(
        By.ID, 'ctl00_Content_Login1_UserName').send_keys(loginusername)
    driver.find_element(
        By.ID, 'ctl00_Content_Login1_Password').send_keys(psw)
    try:
        driver.find_element(By.ID, 'ctl00_Content_Login1_LoginButton').click()
        if is_element_existed(driver, 'ctl00_Content_radioButtonEmail'):
            driver.find_element(
                By.ID, 'ctl00_Content_radioButtonEmail').click()
            driver.find_element(
                By.ID, 'ctl00_Content_ButtonMultiFactDeliveryMethod').click()
            code = input('Please Enter Verfication Code: \n')
            driver.find_element(
                By.ID, 'ctl00_Content_txtMultiFactorAccessCode').send_keys(code)
            driver.find_element(
                By.ID, "ctl00_Content_chkRememberDevice").click()
            # login
            driver.find_element(
                By.ID, 'ctl00_Content_ButtonMultiFactorAccessCodeEntry').click()
        print('login successfully!')
    except:
        print('authentication failed. please check your credentials.')
    pendo_src = driver.find_element(By.ID, "ContentFrame").get_attribute("src")
    pendo_src = pendo_src.split("?")[1]
    link1st = f'https://{hostname}.ultipro.com/pages/VIEW/BusinessIntelligenceLinks.aspx?pendopdk=cogno&pendoWfiHow=new-menu&{pendo_src}&destinationId=1932&USParams=PK=COGNO!MenuID=1932!PageRerId=1932!ParentRerId=2137'

    driver.get(link1st)
    cookies_1st = driver.get_cookies()
    biaddress = returnBI_address(pendo_src, hostname, cookies_1st)
    bilink_request = parse_xml_bilink(biaddress)
    if bilink_request is None:
        return None
    bilink_request = bilink_request.replace("amp;", "")
    driver.get(bilink_request)
    driver.get(
        f"https://c{hostname}.ultipro.com/ibmcognos/bi/?perspective=schedule_mgmt&id=subscribe_mgmt")
    userSecret = driver.execute_script("return inlineContext")[
        'preferences']['id']

    schedule_cookies = driver.get_cookies()
    driver.get(
        f'https://c{hostname}.ultipro.com/ibmcognos/bi/v1/plugins/themes/current/favicon?v=1671661251082')

    print('read cookies')
    cookies = {}
    for cookie in schedule_cookies:
        name = cookie['name']
        value = cookie['value']
        cookies[name] = value
    with open('BI_cookies.json', 'w') as f:
        json.dump(cookies, f)
    print('write cookies to file.')

    return userSecret


def parse_xml_bilink(xmlstring):
    link = re.search("value](.+?)\|", xmlstring)
    if link is None:
        return None
    else:
        return link.group(1)


def read_storecookies():
    with open('BI_cookies.json', 'r') as f:
        cookies = json.load(f)
        return cookies


def makerequest_report(cookies, hostname, json_data):
    headers = {
        'Accept': '*/*',
        'Accept-Language': 'en-US,en;q=0.9',
        'Connection': 'keep-alive',
        'Content-Type': 'application/json',
        'Origin': f'https://c{hostname}.ultipro.com',
        'Referer': f'https://c{hostname}.ultipro.com/ibmcognos/bi/?perspective=schedule_mgmt&id=subscribe_mgmt',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
        'X-Requested-With': 'XMLHttpRequest',
        'X-XSRF-TOKEN': cookies['XSRF-TOKEN'],
        'sec-ch-ua': '"Not?A_Brand";v="8", "Chromium";v="108", "Google Chrome";v="108"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
    }

    response = requests.post(
        f'https://c{hostname}.ultipro.com/ibmcognos/bi/v1/activity_schedules',
        cookies=cookies,
        headers=headers,
        json=json_data,
    )
    return response


def write_dateto_Excel(response):
    if(response.status_code == 200):
        print('request successful! \n')
        user = os.getlogin()
        marker = time.strftime("%Y%m%d-%H%M%S")
        d = json.loads(response.text)
        schedules = d['schedules']
        header = list(schedules[0].keys())
        outputarr = []
        outputarr.append(header)  # type: ignore
        for schedule in schedules:
            row = list(schedule.values())
            outputarr.append(row)
        path = rf'C:\Users\{user}\Desktop\BIreport_schedule_list{marker}.xlsx'
        try:
            pd.DataFrame(outputarr).to_excel(
                path, header=False, index=False)
            return path
        except:
            print('export file failed... \n below is the data')
            print(outputarr)
            input("Press Enter to continue...")
    else:
        print('Error occured. Report request Failed.')


if __name__ == "__main__":

    try:
        hostname = input('Please Enter your UKG host name: \n')
        """loging UKG and getting the cookies of BI"""
        id = ukg_bi_cookies(hostname=hostname)
        print('cookie fetched \n')
        cookies = read_storecookies()
        print('cookie Read \n')
        print('making report request... \n')
        reqs = [
            {
                'type': 'schedule',
                'filters': {
                    'user': id,
                },
                'pageSize': 2000,
            }, {
                'type': 'schedule',
                'pageSize': 2000,
            }
        ]
        for i in reqs:
            response = makerequest_report(
                cookies=cookies, hostname=hostname, json_data=i)
            print('exporting data... \n')
            path = write_dateto_Excel(response=response)
            print(f"file expored! file path is: {path} \n Program completed.")

    except:
        print('Error Occured')
