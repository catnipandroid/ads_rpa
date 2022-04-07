import re
import openpyxl
import math
from datetime import datetime

# Load Excels
incipio_wb = openpyxl.load_workbook(r'static\completed_excel\인시피오 리포트.xlsx', data_only=True)

# Select Daily Data Sheets
incipio_ws_dailySummary = incipio_wb['일일요약']

# Select Monthly Summary Data Sheets
incipio_ws_total = incipio_wb['종합요약']

# PowerLink Data
incipio_powerlink_imps = incipio_ws_dailySummary['C51'].value
incipio_powerlink_clicks = incipio_ws_dailySummary['D51'].value
incipio_powerlink_cost = math.floor(incipio_ws_dailySummary['G51'].value)
incipio_powerlink_conv = incipio_ws_dailySummary['H51'].value
incipio_powerlink_convAmt = incipio_ws_dailySummary['K51'].value

# GoogleDA Data
incipio_googleDA_imps = incipio_ws_dailySummary['C56'].value
incipio_googleDA_clicks = incipio_ws_dailySummary['D56'].value
incipio_googleDA_cost = math.floor(incipio_ws_dailySummary['G56'].value)
incipio_googleDA_conv = incipio_ws_dailySummary['H56'].value
incipio_googleDA_convAmt = incipio_ws_dailySummary['K56'].value

# TOTAL DATA
incipio_total_imps = incipio_ws_total['C65'].value
incipio_total_clicks = incipio_ws_total['D65'].value
incipio_total_conv = incipio_ws_total['H65'].value
incipio_total_convAmt = incipio_ws_total['K65'].value
incipio_total_cost = math.floor(incipio_ws_total['G65'].value)

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
with open(r'comment\text_files\Incipio.txt', 'w', encoding='UTF8') as f:
    f.write(total_title)
    f.write(' - 인시피오: 노출 {:,} / 클릭수 {:,} / 전환수 {:,} / 전환매출 {:,} / 총광고비 {:,}\n'.format(incipio_total_imps,
            incipio_total_clicks, incipio_total_conv, incipio_total_convAmt, incipio_total_cost))
    f.write('\n')
    
    f.write(powerlink_title)
    f.write(' - 인시피오: 노출 {:,} / 클릭수 {:,} / 전환수 {:,} / 전환매출 {:,} / 광고비 {:,}\n'.format(incipio_powerlink_imps,
            incipio_powerlink_clicks, incipio_powerlink_conv, incipio_powerlink_convAmt, incipio_powerlink_cost))
    f.write('\n')
    
    f.write(googleDA_title)
    f.write(' - 인시피오: 노출 {:,} / 클릭수 {:,} / 전환수 {:,} / 전환매출 {:,} / 광고비 {:,}\n'.format(incipio_googleDA_imps,
            incipio_googleDA_clicks, incipio_googleDA_conv, incipio_googleDA_convAmt, incipio_googleDA_cost))
    f.write('\n')