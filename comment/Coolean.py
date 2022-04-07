import re
import openpyxl
import math
from datetime import datetime

# Year, Month, Day
year = str(datetime.now().year)
month = str(datetime.now().month)
day = str((datetime.now().day))

# Load Excels
coolean_wb = openpyxl.load_workbook(r'static\completed_excel\쿨린_보고서.xlsx', data_only=True)

# Select Sheets
coolean_ws_summary = coolean_wb['월 요약']

# weekly data arrays
coolean_facebook_imps = []
coolean_facebook_click = []
coolean_facebook_cost = []
coolean_facebook_conv_amt = []
coolean_facebook_profit = []

# Weekly for loop
def week_report_coolean(start_cell):
    
    for i in range(1,32):
        
        coolean_facebook_imps.append(coolean_ws_summary['D'+str(start_cell)].value)
        coolean_facebook_click.append(coolean_ws_summary['E'+str(start_cell)].value)
        coolean_facebook_cost.append(coolean_ws_summary['H'+str(start_cell)].value)
        coolean_facebook_conv_amt.append(coolean_ws_summary['I'+str(start_cell)].value)
        coolean_facebook_profit.append(coolean_ws_summary['L'+str(start_cell)].value)
        
        start_cell += 1        

# 스타트셀 선택
week_report_coolean(68)

# Title Texts
total_title = '■ 페이스북 성과 요약'


# File Save
with open(r'comment\text_files\Coolean.txt', 'w', encoding='UTF8') as f:
   
    f.write('\n')
    f.write(total_title)
    f.write('\n')
    for idx,i in enumerate(coolean_facebook_imps):
        f.write(str(idx+1)+'일')
        f.write(' - 쿨린 페이스북: 노출 {:,} / 클릭수 {:,} / 광고비 {:,} / 전환수 {:,} / 전환매출 {:,}\n'.format(coolean_facebook_imps[idx], coolean_facebook_click[idx], coolean_facebook_cost[idx], coolean_facebook_conv_amt[idx], coolean_facebook_profit[idx]))