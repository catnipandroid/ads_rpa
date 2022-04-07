import re
import openpyxl
import math
from datetime import datetime

griffin_wb = openpyxl.load_workbook(r'static\completed_excel\그리핀 리포트.xlsx', data_only=True)

# Select Daily Data Sheets
griffin_ws_dailySummary = griffin_wb['일일요약']

# Select Monthly Summary Data Sheets
griffin_ws_total = griffin_wb['종합요약']

# PowerLink Data
griffin_powerlink_imps = griffin_ws_dailySummary['C49'].value
griffin_powerlink_clicks = griffin_ws_dailySummary['D49'].value
griffin_powerlink_cost = math.floor(griffin_ws_dailySummary['G49'].value)
griffin_powerlink_conv = griffin_ws_dailySummary['H49'].value
griffin_powerlink_convAmt = griffin_ws_dailySummary['K49'].value

# Google Video DATA
griffin_google_video_imps = griffin_ws_dailySummary['C54'].value
griffin_google_video_clicks = griffin_ws_dailySummary['D54'].value
griffin_google_video_cost = math.floor(griffin_ws_dailySummary['G54'].value)
griffin_google_video_conv = griffin_ws_dailySummary['H54'].value
griffin_google_video_convAmt = griffin_ws_dailySummary['K54'].value

# TOTAL DATA
griffin_total_imps = griffin_ws_total['C63'].value
griffin_total_clicks = griffin_ws_total['D63'].value
griffin_total_conv = griffin_ws_total['H63'].value
griffin_total_convAmt = griffin_ws_total['K63'].value
griffin_total_cost = math.floor(griffin_ws_total['G63'].value)

# Title Texts
total_title = '■ 매체종합 (월 누적)\n'
brand_title = '■ 브랜드검색 일일\n'
powerlink_title = '■ 파워링크 일일\n'
googleSA_title = '■ 구글검색 일일\n'
googleDC_title = '■ 구글디스커버리 일일\n'
googleDA_title = '■ 구글DA 일일\n'
googleShopping_title = '■ 구글쇼핑 일일\n'
googleVideo_title = '■ 구글유튜브 일일\n'
NaverGFA_title = '■ 네이버GFA 일일\n'
google_pm_title = '■ 구글퍼포먼스맥스 일일\n'
kakaomoment_title = '■ 카카오모먼트 일일\n'

# File Save
with open(r'comment\text_files\Griffin.txt', 'w', encoding='UTF8') as f:
    f.write(total_title)
    f.write(' - 그리핀: 노출 {:,} / 클릭수 {:,} / 전환수 {:,} / 전환매출 {:,} / 총광고비 {:,}\n'.format(griffin_total_imps,
            griffin_total_clicks, griffin_total_conv, griffin_total_convAmt, griffin_total_cost))
    f.write('\n')
    
    f.write(powerlink_title)
    f.write(' - 그리핀: 노출 {:,} / 클릭수 {:,} / 전환수 {:,} / 전환매출 {:,} / 광고비 {:,}\n'.format(griffin_powerlink_imps,
            griffin_powerlink_clicks, griffin_powerlink_conv, griffin_powerlink_convAmt, griffin_powerlink_cost))
    f.write('\n')
    
    f.write(googleVideo_title)
    f.write(' - 그리핀: 노출 {:,} / 클릭수 {:,} / 전환수 {:,} / 전환매출 {:,} / 광고비 {:,}\n'.format(griffin_google_video_imps,
            griffin_google_video_clicks, griffin_google_video_conv, griffin_google_video_convAmt, griffin_google_video_cost))   
    f.write('\n')
    
    