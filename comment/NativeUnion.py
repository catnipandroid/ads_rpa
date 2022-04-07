import re
import openpyxl
import math
from datetime import datetime

nu_wb = openpyxl.load_workbook(r'static\completed_excel\네이티브유니온 리포트.xlsx', data_only=True)

# Select Daily Data Sheets
nu_ws_dailySummary = nu_wb['일일요약']

# Select Monthly Summary Data Sheets
nu_ws_total = nu_wb['종합요약']

# BrandSearchData
nu_brand_imps = nu_ws_dailySummary['C49'].value
nu_brand_clicks = nu_ws_dailySummary['D49'].value
nu_brand_cost = math.floor(nu_ws_dailySummary['G49'].value)
nu_brand_conv = nu_ws_dailySummary['H49'].value
nu_brand_convAmt = nu_ws_dailySummary['K49'].value

# PowerLink Data
nu_powerlink_imps = nu_ws_dailySummary['C52'].value
nu_powerlink_clicks = nu_ws_dailySummary['D52'].value
nu_powerlink_cost = math.floor(nu_ws_dailySummary['G52'].value)
nu_powerlink_conv = nu_ws_dailySummary['H52'].value
nu_powerlink_convAmt = nu_ws_dailySummary['K52'].value

# KAKAO Moment Data
nu_moment_imps = nu_ws_dailySummary['C55'].value
nu_moment_clicks = nu_ws_dailySummary['D55'].value
nu_moment_cost = math.floor(nu_ws_dailySummary['G55'].value)
nu_moment_conv = nu_ws_dailySummary['H55'].value
nu_moment_convAmt = nu_ws_dailySummary['K55'].value

# GoogleSA DATA
nu_googleSA_imps = nu_ws_dailySummary['C68'].value
nu_googleSA_clicks = nu_ws_dailySummary['D68'].value
nu_googleSA_cost = math.floor(nu_ws_dailySummary['G68'].value)
nu_googleSA_conv = nu_ws_dailySummary['H68'].value
nu_googleSA_convAmt = nu_ws_dailySummary['K68'].value

# GoogleDA DATA
nu_googleDA_imps = nu_ws_dailySummary['C73'].value
nu_googleDA_clicks = nu_ws_dailySummary['D73'].value
nu_googleDA_cost = math.floor(nu_ws_dailySummary['G73'].value)
nu_googleDA_conv = nu_ws_dailySummary['H73'].value
nu_googleDA_convAmt = nu_ws_dailySummary['K73'].value

# Google Shopping DATA
nu_googleShopping_imps = nu_ws_dailySummary['C60'].value
nu_googleShopping_clicks = nu_ws_dailySummary['D60'].value
nu_googleShopping_cost = math.floor(nu_ws_dailySummary['G60'].value)
nu_googleShopping_conv = nu_ws_dailySummary['H60'].value
nu_googleShopping_convAmt = nu_ws_dailySummary['K60'].value

# TOTAL DATA
nu_total_imps = nu_ws_total['C78'].value
nu_total_clicks = nu_ws_total['D78'].value
nu_total_conv = nu_ws_total['H78'].value
nu_total_convAmt = nu_ws_total['K78'].value
nu_total_cost = math.floor(nu_ws_total['G78'].value)


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
with open(r'comment\text_files\nativeUnion.txt', 'w', encoding='UTF8') as f:
    f.write(total_title)
    f.write(' - 네이티브유니온: 노출 {:,} / 클릭수 {:,} / 전환수 {:,} / 전환매출 {:,} / 총광고비 {:,}\n'.format(
            nu_total_imps, nu_total_clicks, nu_total_conv, nu_total_convAmt, nu_total_cost))
    f.write('\n')
    
    f.write(brand_title)
    f.write(' - 네이티브유니온: 노출 {:,} / 클릭수 {:,} / 전환수 {:,} / 전환매출 {:,} / 광고비 {:,}\n'.format(
        nu_brand_imps, nu_brand_clicks, nu_brand_conv, nu_brand_convAmt, nu_brand_cost))    
    f.write('\n')
    
    f.write(powerlink_title)
    f.write(' - 네이티브유니온: 노출 {:,} / 클릭수 {:,} / 전환수 {:,} / 전환매출 {:,} / 광고비 {:,}\n'.format(
        nu_powerlink_imps, nu_powerlink_clicks, nu_powerlink_conv, nu_powerlink_convAmt, nu_powerlink_cost))
    f.write('\n')
    
    f.write(kakaomoment_title)
    f.write(' - 네이티브유니온: 노출 {:,} / 클릭수 {:,} / 전환수 {:,} / 전환매출 {:,} / 광고비 {:,}\n'.format(
        nu_moment_imps, nu_moment_clicks, nu_moment_conv, nu_moment_convAmt, nu_moment_cost))
    f.write('\n')
    
    f.write(googleSA_title)
    f.write(' - 네이티브유니온: 노출 {:,} / 클릭수 {:,} / 전환수 {:,} / 전환매출 {:,} / 광고비 {:,}\n'.format(
        nu_googleSA_imps, nu_googleSA_clicks, nu_googleSA_conv, nu_googleSA_convAmt, nu_googleSA_cost))    
    f.write('\n')
    
    f.write(googleDA_title)
    f.write(' - 네이티브유니온: 노출 {:,} / 클릭수 {:,} / 전환수 {:,} / 전환매출 {:,} / 광고비 {:,}\n'.format(
        nu_googleDA_imps, nu_googleDA_clicks, nu_googleDA_conv, nu_googleDA_convAmt, nu_googleDA_cost))    
    f.write('\n')
 
    f.write(googleShopping_title)   
    f.write(' - 네이티브유니온: 노출 {:,} / 클릭수 {:,} / 전환수 {:,} / 전환매출 {:,} / 광고비 {:,}\n'.format(nu_googleShopping_imps,
         nu_googleShopping_clicks, nu_googleShopping_conv, nu_googleShopping_convAmt, nu_googleShopping_cost))        
    f.write('\n')
    
    