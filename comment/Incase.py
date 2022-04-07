import re
import openpyxl
import math
from datetime import datetime

# Load Excels
incase_wb = openpyxl.load_workbook(r'static\completed_excel\인케이스 리포트.xlsx', data_only=True)

# Select Daily Data Sheets
incase_ws_dailySummary = incase_wb['일일요약']

# Select Monthly Summary Data Sheets
incase_ws_total = incase_wb['종합요약']

# BrandSearchData
incase_brand_imps = incase_ws_dailySummary['C51'].value
incase_brand_clicks = incase_ws_dailySummary['D51'].value
incase_brand_cost = math.floor(incase_ws_dailySummary['G51'].value)
incase_brand_conv = math.floor(incase_ws_dailySummary['H51'].value)
incase_brand_convAmt = math.floor(incase_ws_dailySummary['K51'].value)

# GoogleSA DATA
incase_googleSA_imps = incase_ws_dailySummary['C64'].value
incase_googleSA_clicks = incase_ws_dailySummary['D64'].value
incase_googleSA_cost = math.floor(incase_ws_dailySummary['G64'].value)
incase_googleSA_conv = math.floor(incase_ws_dailySummary['H64'].value)
incase_googleSA_convAmt = math.floor(incase_ws_dailySummary['K64'].value)

# GoogleDA DATA
incase_googleDA_imps = incase_ws_dailySummary['C69'].value
incase_googleDA_clicks = incase_ws_dailySummary['D69'].value
incase_googleDA_cost = math.floor(incase_ws_dailySummary['G69'].value)
incase_googleDA_conv = math.floor(incase_ws_dailySummary['H69'].value)
incase_googleDA_convAmt = math.floor(incase_ws_dailySummary['K69'].value)

# Google Shopping DATA
incase_googleShopping_imps = incase_ws_dailySummary['C59'].value
incase_googleShopping_clicks = incase_ws_dailySummary['D59'].value
incase_googleShopping_cost = math.floor(incase_ws_dailySummary['G59'].value)
incase_googleShopping_conv = math.floor(incase_ws_dailySummary['H59'].value)
incase_googleShopping_convAmt = math.floor(incase_ws_dailySummary['K59'].value)

# TOTAL DATA
incase_total_imps = incase_ws_total['C80'].value
incase_total_clicks = incase_ws_total['D80'].value
incase_total_conv = math.floor(incase_ws_total['H80'].value)
incase_total_convAmt = math.floor(incase_ws_total['K80'].value)
incase_total_cost = math.floor(incase_ws_total['G80'].value)

# Title Texts
total_title = '■ 매체종합 (월 누적)\n'
brand_title = '■ 브랜드검색 일일\n'
googleSA_title = '■ 구글검색 일일\n'
googleDA_title = '■ 구글DA 일일\n'
googleShopping_title = '■ 구글쇼핑 일일\n'

# File Save
with open(r'comment\text_files\Incase.txt', 'w', encoding='UTF8') as f:
    
    f.write(total_title)
    f.write(' - 인케이스: 노출 {:,} / 클릭수 {:,} / 전환수 {:,} / 전환매출 {:,} / 총광고비 {:,}\n'.format(
        incase_total_imps, incase_total_clicks, incase_total_conv, incase_total_convAmt, incase_total_cost))
    f.write('\n')
    
    f.write(brand_title)
    f.write(' - 인케이스: 노출 {:,} / 클릭수 {:,} / 전환수 {:,} / 전환매출 {:,} / 광고비 {:,}\n'.format(
        incase_brand_imps, incase_brand_clicks, incase_brand_conv, incase_brand_convAmt, incase_brand_cost))        
    f.write('\n')
    
    f.write(googleSA_title)
    f.write(' - 인케이스: 노출 {:,} / 클릭수 {:,} / 전환수 {:,} / 전환매출 {:,} / 광고비 {:,}\n'.format(incase_googleSA_imps,
            incase_googleSA_clicks, incase_googleSA_conv, incase_googleSA_convAmt, incase_googleSA_cost))
    f.write('\n')
    
    f.write(googleDA_title)
    f.write(' - 인케이스: 노출 {:,} / 클릭수 {:,} / 전환수 {:,} / 전환매출 {:,} / 광고비 {:,}\n'.format(incase_googleDA_imps,
    incase_googleDA_clicks, incase_googleDA_conv, incase_googleDA_convAmt, incase_googleDA_cost))
    f.write('\n')
    
    f.write(googleShopping_title)
    f.write(' - 인케이스: 노출 {:,} / 클릭수 {:,} / 전환수 {:,} / 전환매출 {:,} / 광고비 {:,}\n'.format(incase_googleShopping_imps,
            incase_googleShopping_clicks, incase_googleShopping_conv, incase_googleShopping_convAmt, incase_googleShopping_cost))
        