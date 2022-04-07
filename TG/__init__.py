import openpyxl
from openpyxl.styles import Font, Color, Fill, Alignment, PatternFill, Border, Side
import requests
import time
import datetime
from datetime import date, timedelta
import json
from naversa_util import signaturehelper

########################
### TG API ###
########################
api_key = '735445466f6c523444456d4a362f4b6a624e442f2f413d3d'

tg_cell_count = 26

# 데이터를 일별 분리 (각 캠페인의 데이터를 합침)
impression_cnt = []
click_cnt = []
ad_cost = []
conversion_cnt = []
total_basket_value = []

for i in days31[0:Day]:

    imps_num = 0
    clicks_num = 0
    ads_cost_num = 0
    ccnt_num = 0
    convAmt_num = 0

    api_url = 'https://tgp.widerplanet.com/rest/v3/reportapi/performance-report?key={0}&date={1}&type=json'.format(
        api_key, Year+Month+str(i))
    url = requests.get(api_url)
    text = url.text
    data = json.loads(text)

    for idx, j in enumerate(data):
        imps_num += data[idx]['impression_cnt']
        clicks_num += data[idx]['click_cnt']
        ads_cost_num += data[idx]['ad_cost']
        ccnt_num += data[idx]['conversion_cnt']
        ccnt_num += data[idx]['view_through_conversions']
        convAmt_num += data[idx]['total_basket_value']
        convAmt_num += data[idx]['total_view_through_basket_value']

    impression_cnt.append(imps_num)
    click_cnt.append(clicks_num)
    ad_cost.append(ads_cost_num)
    conversion_cnt.append(ccnt_num)
    total_basket_value.append(convAmt_num)

# 티지 엑셀 시트 저장
for idx, i in enumerate(impression_cnt):
    imps_data = impression_cnt[idx]
    clicks_data = click_cnt[idx]
    adCost_data = ad_cost[idx]
    ccnt_data = conversion_cnt[idx]
    convAmt_data = total_basket_value[idx]

    ws_tg['D'+str(tg_cell_count)] = imps_data
    ws_tg['E'+str(tg_cell_count)] = clicks_data
    ws_tg['H'+str(tg_cell_count)] = adCost_data
    ws_tg['I'+str(tg_cell_count)] = ccnt_data
    ws_tg['K'+str(tg_cell_count)] = convAmt_data

    tg_cell_count += 1


# 리포트 저장하기 위해서 문자열 포맷팅
reportDateName = '블라이드 11월 리포트_{}{}{}'.format(str(datetime.datetime.now(
).year), str(datetime.datetime.now().month), datetime.datetime.now().day)
# 리포트 저장
wb.save(r'ADS_RPA\Blithe\Reports\completed_reports\{0}.xlsx'.format(
    reportDateName))
wb.close()
