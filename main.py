import time
import os
import re
import requests
import json
import pandas as pd
from bs4 import BeautifulSoup


def return_pendo(data):
    pattern = re.compile('var contentIframeInitialSrc = (.*?);')
    scripts = data.find_all('script', string=pattern)
    if len(scripts) > 0:
        pendo = scripts[0]
        pendo = re.search(r'\"(.*?)\"', str(pendo)
                          ).group(1).split("?")[1]  # type: ignore
        return pendo


def return_user_ID(data):
    pattern = re.compile('var inlineContext = (.*?);')
    scripts = data.find_all('script', string=pattern)
    if len(scripts) > 0:
        script = scripts[0]
        id = re.search(r"(\"id\"\:\")(.+?)(\")", str(script)
                       ).group(2)  # type: ignore
        return id


def session_request(hostname, username, password,mfa_method):
    s = requests.Session()
    default_page = s.get(f'https://{hostname}.ultipro.com/Login.aspx')
    if default_page.status_code == 200:
        default_page_html = default_page.text
        b = BeautifulSoup(default_page_html, "html.parser")
        VIEWSTATE = b.find(attrs={'id': '__VIEWSTATE'})['value']  # type: ignore
        VIEWSTATEGENERATOR = b.find(attrs={'id': '__VIEWSTATEGENERATOR'})['value']  # type: ignore
        EVENTVALIDATION = b.find(attrs={'id': '__EVENTVALIDATION'})['value']  # type: ignore

        login1st_data = {
            '__EVENTTARGET': '',
            '__EVENTARGUMENT': '',
            '__LASTFOCUS': '',
            '__VIEWSTATE': VIEWSTATE,
            '__VIEWSTATEGENERATOR': VIEWSTATEGENERATOR,
            '__EVENTVALIDATION': EVENTVALIDATION,
            'ctl00$Content$Login1$UserName': username,
            'ctl00$Content$Login1$Password': password,
            'ctl00$Content$Login1$LoginButton': 'Sign in',
            'ctl00$Content$languagesSelection': '0',
        }

        login1st_header = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'Accept-Language': 'en-US,en;q=0.9',
            'Cache-Control': 'max-age=0',
            'Connection': 'keep-alive',
            'Origin': f'https://{hostname}.ultipro.com',
            'Referer': f'https://{hostname}.ultipro.com/Login.aspx',
            'Sec-Fetch-Dest': 'document',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-Site': 'same-origin',
            'Sec-Fetch-User': '?1',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
            'sec-ch-ua': '"Not?A_Brand";v="8", "Chromium";v="108", "Google Chrome";v="108"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
        }

        fa2_page = s.post(f'https://{hostname}.ultipro.com/Login.aspx',
                          headers=login1st_header, data=login1st_data)

        if fa2_page.status_code == 200:
            """this step will send 2fa code to email,default email"""
            fa2_page_html = BeautifulSoup(fa2_page.text, 'html.parser')
            VIEWSTATE_2fa = fa2_page_html.find(attrs={'id': '__VIEWSTATE'})['value']  # type: ignore
            VIEWSTATEGENERATOR_2fa = fa2_page_html.find(
                attrs={'id': '__VIEWSTATEGENERATOR'})['value']  # type: ignore
            EVENTVALIDATION_2fa = fa2_page_html.find(attrs={'id': '__EVENTVALIDATION'})['value']  # type: ignore
            login2fa_data = {
                '__EVENTTARGET': '',
                '__EVENTARGUMENT': '',
                '__LASTFOCUS': '',
                '__VIEWSTATE': VIEWSTATE_2fa,
                '__VIEWSTATEGENERATOR': VIEWSTATEGENERATOR_2fa,
                '__EVENTVALIDATION': EVENTVALIDATION_2fa,
                'ctl00$Content$ddlMfaText': 'HomePhone',
                'ctl00$Content$ddlMfaVoice': 'BusinessPhone',
                'ctl00$Content$deliverySelection': mfa_method,  # radioButtonEmail
                'ctl00$Content$ddlMfaEmail': '1',
                'ctl00$Content$ButtonMultiFactDeliveryMethod': 'Submit',
                'ctl00$Content$languagesSelection': '0',
            }
            fa2_code_input = s.post(
                f'https://{hostname}.ultipro.com/Login.aspx', headers=login1st_header, data=login2fa_data)
            if fa2_code_input.status_code == 200:
                code_input_page = BeautifulSoup(
                    fa2_code_input.text, 'html.parser')
                VIEWSTATE_2fa = code_input_page.find(attrs={'id': '__VIEWSTATE'})['value']  # type: ignore
                VIEWSTATEGENERATOR_2fa = code_input_page.find(
                    attrs={'id': '__VIEWSTATEGENERATOR'})['value']  # type: ignore
                EVENTVALIDATION_2fa = code_input_page.find(attrs={'id': '__EVENTVALIDATION'})['value']  # type: ignore

                email_code = input("Please Enter Your 2fa Code: \n")
                print('please wait...')
                fa2_code_input_data = {
                    '__EVENTTARGET': '',
                    '__EVENTARGUMENT': '',
                    '__LASTFOCUS': '',
                    '__VIEWSTATE': VIEWSTATE_2fa,
                    '__VIEWSTATEGENERATOR': VIEWSTATEGENERATOR_2fa,
                    '__EVENTVALIDATION': EVENTVALIDATION_2fa,
                    'ctl00$Content$txtMultiFactorAccessCode': email_code,
                    'ctl00$Content$ButtonMultiFactorAccessCodeEntry': 'Submit',
                    'ctl00$Content$languagesSelection': '0',
                }
                s.post(f'https://{hostname}.ultipro.com/Login.aspx?ReturnUrl=%2fdefault.aspx',
                       headers=login1st_header, data=fa2_code_input_data)
                login1st_header['Referer'] = f'https://{hostname}.ultipro.com/Login.aspx?ReturnUrl=%2fdefault.aspx'
                s.get(
                    f'https://{hostname}.ultipro.com/pages/utility/newsession.aspx', headers=login1st_header)
                s.get(
                    f'https://{hostname}.ultipro.com/pages/utility/companyselector.aspx', headers=login1st_header)
                s.get(
                    f'https://{hostname}.ultipro.com/default.aspx', headers=login1st_header)
                params = {
                    'action': 'set',
                    'res': '3072x1728',
                    'd': '24',
                }
                headers = {
                    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
                    'Accept-Language': 'en-US,en;q=0.9',
                    'Connection': 'keep-alive',
                    'Referer': f'https://{hostname}.ultipro.com/default.aspx',
                    'Sec-Fetch-Dest': 'document',
                    'Sec-Fetch-Mode': 'navigate',
                    'Sec-Fetch-Site': 'same-origin',
                    'Upgrade-Insecure-Requests': '1',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
                    'sec-ch-ua': '"Not?A_Brand";v="8", "Chromium";v="108", "Google Chrome";v="108"',
                    'sec-ch-ua-mobile': '?0',
                    'sec-ch-ua-platform': '"Windows"',
                }
                s.get(f'https://{hostname}.ultipro.com/DetectScreen.aspx',
                      params=params, headers=headers)
                res = s.get(
                    f'https://{hostname}.ultipro.com/default.aspx',  headers=headers)
                pendo_data = BeautifulSoup(res.text, 'html.parser')
                pendoid = return_pendo(pendo_data)
                headers = {
                    'Accept': 'text/plain, */*; q=0.01',
                    'Accept-Language': 'en-US,en;q=0.9',
                    'Connection': 'keep-alive',
                    'Origin': f'https://{hostname}.ultipro.com',
                    'Referer': f'https://{hostname}.ultipro.com/pages/VIEW/BusinessIntelligenceLinks.aspx?pendopdk=cogno&pendoWfiHow=new-menu&{pendoid}&destinationId=1932&USParams=PK=COGNO!MenuID=1932!PageRerId=1932!ParentRerId=2137',
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
                response = s.post(
                    f'https://{hostname}.ultipro.com/services/CommonService.asmx/GetCognosSsoUrl?pendopdk=cogno&pendoWfiHow=new-menu&{pendoid}&destinationId=1932&USParams=PK=COGNO!MenuID=1932!PageRerId=1932!ParentRerId=2137',
                    headers=headers,
                    data=data,
                )
                bilink = parse_xml_bilink(response.text)  # return bi LINK new
                headers = {
                    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
                    'Accept-Language': 'en-US,en;q=0.9',
                    'Connection': 'keep-alive',
                    'Sec-Fetch-Dest': 'document',
                    'Sec-Fetch-Mode': 'navigate',
                    'Sec-Fetch-Site': 'none',
                    'Sec-Fetch-User': '?1',
                    'Upgrade-Insecure-Requests': '1',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
                    'sec-ch-ua': '"Not?A_Brand";v="8", "Chromium";v="108", "Google Chrome";v="108"',
                    'sec-ch-ua-mobile': '?0',
                    'sec-ch-ua-platform': '"Windows"',
                }
                params_data = returnparams_bilink(bilink)
                params = {
                    'Key': params_data['Key'],
                    'SiteID': params_data['SiteID'],
                }
                s.get(f'https://c{hostname}.ultipro.com/ibmcognos/bi/v1/disp',
                      params=params,  headers=headers)  # redirect
                s.get(
                    f'https://c{hostname}.ultipro.com/ibmcognos/bi/v1/disp?b_action=xts.run&m=portal/cc.xts&gohome=', headers=headers)
                response = s.get(
                    f'https://c{hostname}.ultipro.com/ibmcognos/bi/', headers=headers)

                id_text = BeautifulSoup(response.text, 'html.parser')
                id = return_user_ID(id_text)
                
                cookies =  s.cookies.get_dict()
                # make request to get schedule report:
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
                json_data = {
                    'type': 'schedule',
                    'filters': {
                        'user': id,
                    },
                    'pageSize': 2000,
                }
                schedule_report_response = s.post(
                    f'https://c{hostname}.ultipro.com/ibmcognos/bi/v1/activity_schedules',
                    headers=headers,
                    json=json_data,
                )
                #save first report schedule report
                report_schedule_to_Excel(schedule_report_response)
                #return report ids
                reportids = []
                response_json = json.loads(schedule_report_response.text)
                for schedule in response_json['schedules']:
                    reportid = schedule['reportId']
                    reportids.append(reportid)
                # fetch each of the report schedule details. Write to excel
                export_report_details(session=s,cookies=cookies,reportIDS=reportids,hostname=hostname)
                print("done")
                return pendoid



def parse_xml_bilink(xmlstring):
    link = re.search("(\\[value])(.+?)(\\|-|_blank)", xmlstring)
    if link is None:
        return None
    else:
        return link.group(2).replace("amp;", "")


def returnparams_bilink(link):
    obj = {}
    if link != None:
        linksplit = link.split("?")[1].split("&")
        for i in linksplit:
            keyval = i.split("=")
            obj[keyval[0]] = keyval[1]
    # {'Key': 'AF76fefaewfawDABC', 'SiteID': '6DCCfaewfae0E242E6A61'}
    return obj


def read_storecookies():
    with open('BI_cookies.json', 'r') as f:
        cookies = json.load(f)
        return cookies


def report_schedule_to_Excel(response):
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
            df = pd.DataFrame(outputarr, index=None)
            df.to_excel(
                path, header=False, index=False)
        except:
            print('export file failed... \n below is the data')
            print(outputarr)
            input("Press Enter to continue...")
    else:
        print('Error occured. Report request Failed.')


def get_schedule_detail(session,cookies, hostname, reportid):
    
    headers = {
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Language': 'en-US,en;q=0.9',
        'Connection': 'keep-alive',
        'Referer': f'https://c{hostname}.ultipro.com/ibmcognos/bi/?perspective=home',
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

    response = session.get(
        f'https://c{hostname}.ultipro.com/ibmcognos/bi/v1/reports/{reportid}/schedule',
        cookies=cookies,
        headers=headers,
    )
    if response.status_code == 200:
        return response.text
    else:
        return "error"


def export_report_details(session,cookies, reportIDS, hostname):

    user = os.getlogin()
    marker = time.strftime("%Y%m%d-%H%M%S")
    output_arr = []
    output_arr.append(['reportid', 'schedule_by', 'everyNPeriods', 'start_date', 'end_date',
                      'active', 'email_subject', 'email_to','email_body', 'burst', 'prompt', 'outputFormat']) #, 'email_body'

    for id in reportIDS:
        try:
            response_text = get_schedule_detail(session=session,  cookies=cookies, hostname=hostname, reportid=id)
            d_list = parse_schedule_response_detail(response_text)
            if d_list != None:
                output_arr.append(d_list)
        except:
            print('get report schedule detail failed.')
            continue

    path = rf'C:\Users\{user}\Desktop\BIreport_schedule_detail{marker}.xlsx'

    try:
        pd.DataFrame(output_arr).to_excel(
            path, header=False, index=False)
    except:
        print('export file failed... \n below is the data')
        print(output_arr)
        input("Press Enter x to continue...")


def parse_schedule_response_detail(response_text):
    data = json.loads(response_text,strict = False)

    try:
        reportid = data['reportId']
        repeat = data['scheduleInfo']
        """schedule"""
        schedule_by = repeat['type']  # daily
        everyNPeriods = repeat['everyNPeriods']
        end_date = repeat["endDate"]
        start_date = repeat["startDate"]
        active = repeat["active"]  # true or false
        """options"""
        options = data['options']
        deliver = options['delivery']
        emails = deliver['email']
        email_subject = emails['subject']
        email_body = emails['memoPart']
        # # return email list
        email_to =  json.dumps(emails['to'])
        #",".join(list(map(lambda x: x['address'], emails['to']))) 

        burst = options["burst"]
        prompt = options["prompt"]
        outputFormat = ",".join(options["outputFormat"])
        print('Success')
        return [reportid, schedule_by, everyNPeriods, start_date, end_date, active, email_subject, email_to,email_body, burst, prompt, outputFormat] #
    except:
        print("not a scheduled report")

if __name__ == "__main__":

    try:
        hostname = input("Please Enter hostname: \n")
        username = input("Please Enter UserName: \n")
        password = input("Please Enter password: \n")
        fa2_methods = input("Please Choose 2fa code delivery methods: \n 1: Email \n 2: SMS TEXT \n 3: Call.. \n")
        map = {
            "1":"radioButtonEmail",
            "2":"radioButtonSms",
            "3":"radioButtonVoice"
        }
        selection = map[fa2_methods]
        link = session_request(hostname=hostname, username=username, password=password,mfa_method=selection)
    except:
        print('Error Occured')






