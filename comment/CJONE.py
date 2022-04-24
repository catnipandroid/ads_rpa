import re
import openpyxl
import math
from datetime import datetime

# Year, Month, Day
year = str(datetime.now().year)
month = str(datetime.now().month)
day = str((datetime.now().day))

# Load Excels
coolean_wb = openpyxl.load_workbook(r'static\completed_excel\CJ_ONE_보고서.xlsx', data_only=True)

# Select Sheets
cjone_youtube_sheet = coolean_wb['유튜브']
cjone_instagram_sheet = coolean_wb['인스타그램']

# YOUTUBE TOTAL
cjone_total_youtube_imps = cjone_youtube_sheet['E6'].value
cjone_total_youtube_views = cjone_youtube_sheet['F6'].value
cjone_total_youtube_cost = cjone_youtube_sheet['G6'].value
cjone_total_youtube_participate = cjone_youtube_sheet['H6'].value

# INSTAGRAM TOTAL
cjone_total_instagram_imps = cjone_instagram_sheet['F6'].value
cjone_total_instagram_reach = cjone_instagram_sheet['E6'].value
cjone_total_instagram_cost = cjone_instagram_sheet['N6'].value
cjone_total_instagram_participate = cjone_instagram_sheet['G6'].value

#########################################################################

# weekly data arrays
cjone_youtube_imps = []
cjone_youtube_views = []
cjone_youtube_cost = []
cjone_youtube_participate = []

# weekly data arrays
cjone_instagram_imps = []
cjone_instagram_reach = []
cjone_instagram_cost = []
cjone_instagram_participate = []

# Weekly YOUTUBE for loop
def week_report_cjone(start_cell, start_cell2):
    
    for j in range(1,4):
        
        cjone_instagram_imps.append(cjone_instagram_sheet['F'+str(start_cell)].value)
        cjone_instagram_reach.append(cjone_instagram_sheet['E'+str(start_cell)].value)
        cjone_instagram_cost.append(cjone_instagram_sheet['N'+str(start_cell)].value)
        cjone_instagram_participate.append(cjone_instagram_sheet['G'+str(start_cell)].value)

        
        start_cell += 1        
    
    for i in range(1,13):
        
        cjone_youtube_imps.append(cjone_youtube_sheet['E'+str(start_cell2)].value)
        cjone_youtube_views.append(cjone_youtube_sheet['F'+str(start_cell2)].value)
        cjone_youtube_cost.append(cjone_youtube_sheet['G'+str(start_cell2)].value)
        cjone_youtube_participate.append(cjone_youtube_sheet['H'+str(start_cell2)].value)
        
        start_cell2 += 1      


# 스타트셀 선택 (youtube, instagram)
week_report_cjone(83, 73)


# Title Texts
total_youtube_title = '■ 유튜브 성과 누적'
total_instagram_title = '■ 인스타그램 성과 누적'

weekly_youtube_title = '■ 유튜브 주간'
weekly_instagram_title = '■ 인스타그램 주간'


# File Save
with open(r'comment\text_files\CJONE.txt', 'w', encoding='UTF8') as f:
    
    f.write(total_youtube_title)
    f.write('\n')
    f.write(' - 유튜브 누적 성과: 노출 {:,} / 조회수 {:,} / 광고비 {:,} / 참여수 {:,}'.format(cjone_total_youtube_imps, cjone_total_youtube_views, cjone_total_youtube_cost,cjone_total_youtube_participate))
    f.write('\n')
    
    f.write(total_instagram_title)
    f.write('\n')
    f.write(' - 인스타그램 누적 성과: 노출 {:,} / 도달 {:,} / 광고비 {:,} / 참여수 {:,}'.format(cjone_total_instagram_imps, cjone_total_instagram_reach, cjone_total_instagram_cost,cjone_total_instagram_participate))

    f.write('\n')

    ### Weekly loop ###
    f.write(weekly_instagram_title)
    f.write('\n')
    for idx,i in enumerate(cjone_instagram_imps):
        f.write(str(idx+1)+ '일')
        f.write( ' - CJONE 인스타그램: 노출 {:,} / 도달 {:,} / 광고비 {:,} / 참여수 {:,} \n'.format(cjone_instagram_imps[idx], cjone_instagram_reach[idx], cjone_instagram_cost[idx], cjone_instagram_participate[idx]) )
    
    f.write('\n')
    
    f.write(weekly_youtube_title)
    f.write('\n')
    for idx,i in enumerate(cjone_youtube_imps):
        f.write(str(idx+1) + '일')
        f.write(' - CJONE 유튜브: 노출 {:,} / 조회수 {:,} / 광고비 {:,} / 참여수 {:,} \n'.format(cjone_youtube_imps[idx], cjone_youtube_views[idx], cjone_youtube_cost[idx], cjone_youtube_participate[idx]))