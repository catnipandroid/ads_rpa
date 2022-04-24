import re
import openpyxl
import math
from datetime import datetime

# # Year, Month, Day
# year = str(datetime.now().year)
# month = str(datetime.now().month)
# day = str((datetime.now().day))

# Load Excels
qxpress_wb = openpyxl.load_workbook(r'static\completed_excel\큐익스프레스 리포트.xlsx', data_only=True)

# Select Sheets
qxpress_ws_summary = qxpress_wb['summary']

# summary Datas
qxpress_summary_imps = qxpress_ws_summary['E12'].value
qxpress_summary_click = qxpress_ws_summary['F12'].value
qxpress_summary_cost = qxpress_ws_summary['I12'].value
qxpress_summary_conv_amt = qxpress_ws_summary['J12'].value

# powerlink Total Datas
qxpress_powerlink_imps = qxpress_ws_summary['E6'].value
qxpress_powerlink_click = qxpress_ws_summary['F6'].value
qxpress_powerlink_cost = qxpress_ws_summary['I6'].value
qxpress_powerlink_conv_amt = qxpress_ws_summary['J6'].value

# naver brandsearch Total Datas
qxpress_brandsearch_imps = qxpress_ws_summary['E7'].value
qxpress_brandsearch_click = qxpress_ws_summary['F7'].value
qxpress_brandsearch_cost = qxpress_ws_summary['I7'].value
qxpress_brandsearch_conv_amt = qxpress_ws_summary['J7'].value

# google Total Datas
qxpress_google_imps = qxpress_ws_summary['E9'].value
qxpress_google_click = qxpress_ws_summary['F9'].value
qxpress_google_cost = qxpress_ws_summary['I9'].value
qxpress_google_conv_amt = qxpress_ws_summary['J9'].value

# facebook Total Datas
qxpress_facebook_imps = qxpress_ws_summary['E10'].value
qxpress_facebook_click = qxpress_ws_summary['F10'].value
qxpress_facebook_cost = qxpress_ws_summary['I10'].value
qxpress_facebook_conv_amt = qxpress_ws_summary['J10'].value


## Title Texts
total_title = '■ 종합 성과 요약'
total_powerlink_title = '■ 파워링크 누적'
total_brandsearch_title = '■ 파워링크 누적'
total_google_title = '■ 구글애즈 누적'
total_facebook_title = '■ 페이스북 누적'


# File Save
with open(r'comment\text_files\Qxpress.txt', 'w', encoding='UTF8') as f:
    f.write(total_title)
    f.write('\n')
    f.write(' - 노출 {:,} / 클릭수 {:,} / 광고비 {:,} / 전환수 {:,} \n'.format(qxpress_summary_imps, qxpress_summary_click, qxpress_summary_cost, qxpress_summary_conv_amt))
    
    # 파워링크 토탈
    f.write('\n')
    f.write(total_powerlink_title)
    f.write('\n')
    f.write(' - 노출 {:,} / 클릭수 {:,} / 광고비 {:,} / 전환수 {:,} \n'.format(qxpress_powerlink_imps, qxpress_powerlink_click, qxpress_powerlink_cost, qxpress_powerlink_conv_amt))
    
    # 브랜드검색 토탈
    f.write('\n')
    f.write(total_brandsearch_title)
    f.write('\n')
    f.write(' - 노출 {:,} / 클릭수 {:,} / 광고비 {:,} / 전환수 {:,} \n'.format(qxpress_brandsearch_imps, qxpress_brandsearch_click, qxpress_brandsearch_cost, qxpress_brandsearch_conv_amt))
    
    # 구글애즈 토탈
    f.write('\n')
    f.write(total_google_title)
    f.write('\n')
    f.write(' - 노출 {:,} / 클릭수 {:,} / 광고비 {:,} / 전환수 {:,} \n'.format(qxpress_google_imps, qxpress_google_click, qxpress_google_cost, qxpress_google_conv_amt))
    
    # 페이스북 토탈
    f.write('\n')
    f.write(total_facebook_title)
    f.write('\n')
    f.write(' - 노출 {:,} / 클릭수 {:,} / 광고비 {:,} / 전환수 {:,} \n'.format(qxpress_facebook_imps, qxpress_facebook_click, qxpress_facebook_cost, qxpress_facebook_conv_amt))
