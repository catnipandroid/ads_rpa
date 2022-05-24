import requests
import json
from datetime import date, timedelta, datetime
import openpyxl
# from oauth import access_token, refresh_token

report_start_year = input("조회하실 연도를 입력 형식: YYYY ")
report_start_month = input("조회하실 월을 입력 형식: MM ")
report_start_day = input("날짜는 언제부터 조회?: DD ")
report_end_day = input("언제까지의 날짜를 조회? 입력 형식: DD ")

# moment 광고 데이터 분류
moment_data_pc = []
moment_data_mob_Android = []
moment_data_mob_iOS = []


access_token = 'gI1WvnYyvvChY8GqVC0Ry9MLMidN2J0X3adz_Qo9dJkAAAF9WU4ugQ'
refresh_token = 'cNQ2Mi-Rg-gLrtbqPJw62Bvu5bCm60tM9gO99wo9dJkAAAF9WU4ugA'

api_uri = 'https://apis.moment.kakao.com/openapi/v4/adAccounts/report'
adAccountId = '224849'

headers = {
    'accept': 'application/json',
    'Authorization': 'Bearer ' + access_token,
    "adAccountId": adAccountId
}


params = (
    ('start', report_start_year + report_start_month + report_start_day),
    ('end', report_start_year + report_start_month + report_end_day),
    ('level', 'AD_ACCOUNT'),
    ('dimension', 'DEVICE_TYPE'),
    ('metricsGroup', 'BASIC'),
    ('metricsGroup', 'PIXEL_SDK_CONVERSION')
)

response = requests.get(api_uri, headers=headers, params=params, verify=True)
res_data = response.json()

# 기기 등의 보고서 차원, (Android, iOS, PC)
for report_data in res_data['data']:
    # print(report_data)
    if report_data['dimensions']['device_type'] == 'PC':
        moment_data_pc.append(report_data['metrics'])
    elif report_data['dimensions']['device_type'] == 'Android':
        moment_data_mob_Android.append(report_data['metrics'])
    elif report_data['dimensions']['device_type'] == 'iOS':
        moment_data_mob_iOS.append(report_data['metrics'])


# print(moment_data_mob_iOS)

# 워크북 시트 열기
wb = openpyxl.load_workbook(
    r'ADS_RPA\multipop\Reports\exist_excel\네이티브유니온 리포트.xlsx')
ws_moment = wb['카카오모먼트']


# NativeUnion kakaomoment writing in excel cells PC
def nu_moment_data_PC(data, cellNo, sheetName):

    for idx, i in enumerate(data):

        sheetName['C'+str(cellNo)] = data[idx]['imp']
        sheetName['D'+str(cellNo)] = data[idx]['click']
        sheetName['G'+str(cellNo)] = data[idx]['cost']
        sheetName['H'+str(cellNo)] = data[idx]['conv_purchase_7d']
        sheetName['K'+str(cellNo)] = data[idx]['conv_purchase_p_7d']

        cellNo += 1


nu_moment_data_PC(moment_data_pc, 119, ws_moment)


# NativeUnion kakaomoment writing in excel cells MOB
def nu_moment_data_mob(data_android, data_iOS, cellNo, sheetName):

    for idx, i in enumerate(data_android):

        sheetName['C'+str(cellNo)] = data_android[idx]['imp'] + \
            data_iOS[idx]['imp']
        sheetName['D'+str(cellNo)] = data_android[idx]['click'] + \
            data_iOS[idx]['click']
        sheetName['G'+str(cellNo)] = data_android[idx]['cost'] + \
            data_iOS[idx]['cost']
        sheetName['H'+str(cellNo)] = data_android[idx]['conv_purchase_7d'] + \
            data_iOS[idx]['conv_purchase_7d']
        sheetName['K'+str(cellNo)] = data_android[idx]['conv_purchase_p_7d'] + \
            data_iOS[idx]['conv_purchase_p_7d']

        cellNo += 1


nu_moment_data_mob(moment_data_mob_Android,
                   moment_data_mob_iOS, 154, ws_moment)

# 네이티브유니온 리포트 저장하기 위해서 문자열 포맷팅
reportDateName = '네이티브유니온 리포트'
# 리포트 저장
wb.save(r'ADS_RPA\multipop\Reports\completed_reports\{0}.xlsx'.format(
    reportDateName))
wb.close()
