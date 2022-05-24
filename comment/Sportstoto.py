import re
import openpyxl
import math
from datetime import datetime

# # Year, Month, Day
# year = str(datetime.now().year)
# month = str(datetime.now().month)
# day = str((datetime.now().day))

# Load Excels
toto_wb = openpyxl.load_workbook(r'static\completed_excel\스포츠토토_보고서.xlsx', data_only=True)

# Select Sheets
toto_ws_summary = toto_wb['summary']

# summary Datas
toto_summary_imps = toto_ws_summary['E11'].value
toto_summary_click = toto_ws_summary['F11'].value
toto_summary_cost = toto_ws_summary['I11'].value
toto_summary_conv_amt = toto_ws_summary['J11'].value

# TG Total Datas
toto_tg_imps = toto_ws_summary['E6'].value
toto_tg_click = toto_ws_summary['F6'].value
toto_tg_cost = toto_ws_summary['I6'].value
toto_tg_conv_amt = toto_ws_summary['J6'].value

# mobon Total Datas
toto_mobon_imps = toto_ws_summary['E7'].value
toto_mobon_click = toto_ws_summary['F7'].value
toto_mobon_cost = toto_ws_summary['I7'].value
toto_mobon_conv_amt = toto_ws_summary['J7'].value

# ADN Total Datas
toto_ADN_imps = toto_ws_summary['E8'].value
toto_ADN_click = toto_ws_summary['F8'].value
toto_ADN_cost = toto_ws_summary['I8'].value
toto_ADN_conv_amt = toto_ws_summary['J8'].value

# google Total Datas
toto_google_imps = toto_ws_summary['E9'].value
toto_google_click = toto_ws_summary['F9'].value
toto_google_cost = toto_ws_summary['I9'].value
toto_google_conv_amt = toto_ws_summary['J9'].value



#####################################
# 하기는 일자 데이터 보고서
####################################

# Select Daily Sheets
toto_ws_tg = toto_wb['타게팅게이츠']
toto_ws_mobon = toto_wb['모비온']
toto_ws_ADN = toto_wb['ADN']
toto_ws_google = toto_wb['구글']

# TG Daily Datas
toto_tg_daily_imps = toto_ws_tg['E12'].value
toto_tg_daily_click = toto_ws_tg['F12'].value
toto_tg_daily_cost = toto_ws_tg['I12'].value
toto_tg_daily_conv_amt = toto_ws_tg['J12'].value

# Mobon Daily Datas
toto_mobon_daily_imps = toto_ws_mobon['E12'].value
toto_mobon_daily_click = toto_ws_mobon['F12'].value
toto_mobon_daily_cost = toto_ws_mobon['I12'].value
toto_mobon_daily_conv_amt = toto_ws_mobon['J12'].value

# ADN Daily Datas
toto_ADN_daily_imps = toto_ws_ADN['E12'].value
toto_ADN_daily_click = toto_ws_ADN['F12'].value
toto_ADN_daily_cost = toto_ws_ADN['I12'].value
toto_ADN_daily_conv_amt = toto_ws_ADN['J12'].value

# Google Daily Datas
toto_google_daily_imps = toto_ws_google['E12'].value
toto_google_daily_click = toto_ws_google['F12'].value
toto_google_daily_cost = toto_ws_google['I12'].value
toto_google_daily_conv_amt = toto_ws_google['J12'].value


# Title Texts
total_title = '■ 종합 성과 요약'
total_mobon_title = '■ 모비온 누적'
total_tg_title = '■ 타게팅게이츠 누적'
total_ADN_title = '■ ADN 누적'
total_google_title = '■ 구글애즈 누적'

daily_title = '■ 일일 요약'

tg_daily_title = '■ 타게팅게이츠 일일'
mobon_daily_title = '■ 모비온 일일'
ADN_daily_title = '■ ADN 일일'
google_daily_title = '■ 구글애즈 일일'

# File Save
with open(r'comment\text_files\Sportstoto.txt', 'w', encoding='UTF8') as f:
    f.write(total_title)
    f.write('\n')
    f.write(' - 노출 {:,} / 클릭수 {:,} / 광고비 {:,} / 전환수 {:,} \n'.format(toto_summary_imps, toto_summary_click, toto_summary_cost, toto_summary_conv_amt))
    
    f.write('\n')
    f.write(total_mobon_title)
    f.write('\n')
    f.write(' - 노출 {:,} / 클릭수 {:,} / 광고비 {:,} / 전환수 {:,} \n'.format(toto_mobon_imps, toto_mobon_click, toto_mobon_cost, toto_mobon_conv_amt))
    
    f.write('\n')
    f.write(total_tg_title)
    f.write('\n')
    f.write(' - 노출 {:,} / 클릭수 {:,} / 광고비 {:,} / 전환수 {:,} \n'.format(toto_tg_imps, toto_tg_click, toto_tg_cost, toto_tg_conv_amt))
    
    f.write('\n')
    f.write(total_ADN_title)
    f.write('\n')
    f.write(' - 노출 {:,} / 클릭수 {:,} / 광고비 {:,} / 전환수 {:,} \n'.format(toto_ADN_imps, toto_ADN_click, toto_ADN_cost, toto_ADN_conv_amt))
    
    
    f.write('\n')
    f.write(total_google_title)
    f.write('\n')
    f.write(' - 노출 {:,} / 클릭수 {:,} / 광고비 {:,} / 전환수 {:,} \n'.format(toto_google_imps, toto_google_click, toto_google_cost, toto_google_conv_amt))
    
    ############################################ 일일
    f.write('\n')
    f.write(daily_title)
    f.write('\n')
    
    f.write(tg_daily_title)
    f.write(' - 노출 {:,} / 클릭수 {:,} / 광고비 {:,} / 전환수 {:,} \n'.format(toto_tg_daily_imps, toto_tg_daily_click, toto_tg_daily_cost, toto_tg_daily_conv_amt))
    f.write('\n')
    
    f.write(mobon_daily_title)
    f.write(' - 노출 {:,} / 클릭수 {:,} / 광고비 {:,} / 전환수 {:,} \n'.format(toto_mobon_daily_imps, toto_mobon_daily_click, toto_mobon_daily_cost, toto_mobon_daily_conv_amt))
    f.write('\n')
    
    f.write(ADN_daily_title)
    f.write(' - 노출 {:,} / 클릭수 {:,} / 광고비 {:,} / 전환수 {:,} \n'.format(toto_ADN_daily_imps, toto_ADN_daily_click, toto_ADN_daily_cost, toto_ADN_daily_conv_amt))
    f.write('\n')
    
    f.write(google_daily_title)
    f.write(' - 노출 {:,} / 클릭수 {:,} / 광고비 {:,} / 전환수 {:,} \n'.format(toto_google_daily_imps, toto_google_daily_click, toto_google_daily_cost, toto_google_daily_conv_amt))