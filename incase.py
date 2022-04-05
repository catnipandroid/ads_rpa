from NaverSA import NaverSA_API
from excelCreate import Create_Excel
from GoogleAds import GoogleAdsAPI, GoogleAdsPerform
import configparser

config = configparser.ConfigParser()
config.read('config\info.ini')


################### ################### ################### 
################### Naver Search Ads 캠페인별 #############
################### ################### ################### 

# Day by Day Data를 담을 배열
incase_brand_search_pc_data = []
incase_brand_search_mob_data = []

# 브랜드검색 캠페인
incaseBrandAdCampaign = {
    'nccCampaignIdPC': 'cmp-a001-04-000000003086886',
    'nccCampaignIdMob': 'cmp-a001-04-000000003086868',
}

# config 파일 필요함
naver_sa_api = NaverSA_API(config['naver_sa']['multipop_API_KEY'], config['naver_sa']['multipop_SECRET_KEY'], config['naver_sa']['multipop_CUSTOMER_ID'])

# 브랜드검색 호출
naver_sa_api.naverSearch_API_Get(incaseBrandAdCampaign['nccCampaignIdPC'], incase_brand_search_pc_data)
naver_sa_api.naverSearch_API_Get(incaseBrandAdCampaign['nccCampaignIdMob'], incase_brand_search_mob_data)

# Excel
brandsa_excel = Create_Excel('static\exist_excel\인케이스 리포트.xlsx', '네이버 브랜드검색')
brandsa_excel.naver_sa_write(incase_brand_search_pc_data, 56)
brandsa_excel.naver_sa_write(incase_brand_search_mob_data, 91)

brandsa_excel.save('static\completed_excel\인케이스 리포트.xlsx')


################### ################### ################### 
################### Google Ads 캠페인별 ###################
################### ################### ################### 

## Google DA
GoogleDA_excel = Create_Excel('static\completed_excel\인케이스 리포트.xlsx', '구글애즈DA')
# CID 및 캠페인 아이디 전달
googleads_da_api = GoogleAdsAPI('8370773952', 13946306322)
# DB 쿼리 실행 (batch api)
googleads_da_api.get_data()
# 데이터 가져오기 
googleAds_Perform = GoogleAdsPerform()
## Google DA
GoogleDA_excel.google_ads_write(googleAds_Perform.pc_data, 58)
GoogleDA_excel.google_ads_write(googleAds_Perform.mob_data, 93)
GoogleDA_excel.google_ads_write(googleAds_Perform.tablet_data, 128)
GoogleDA_excel.google_ads_write(googleAds_Perform.others_data, 163)
GoogleDA_excel.save('static\completed_excel\인케이스 리포트.xlsx')
googleAds_Perform.clear_data()

## Google SA
GoogleSA_excel = Create_Excel('static\completed_excel\인케이스 리포트.xlsx', '구글애즈SA')
# CID 및 캠페인 아이디 전달
googleads_sa_api = GoogleAdsAPI('8370773952', 13230191890)
# DB 쿼리 실행 (batch api)
googleads_sa_api.get_data()
# 데이터 가져오기 
googleAds_Perform = GoogleAdsPerform()
## Google SA
GoogleSA_excel.google_ads_write(googleAds_Perform.pc_data, 58)
GoogleSA_excel.google_ads_write(googleAds_Perform.mob_data, 93)
GoogleSA_excel.google_ads_write(googleAds_Perform.tablet_data, 128)
GoogleSA_excel.google_ads_write(googleAds_Perform.others_data, 163)
GoogleSA_excel.save('static\completed_excel\인케이스 리포트.xlsx')
googleAds_Perform.clear_data()

## Google Shopping
GoogleShopping_excel = Create_Excel('static\completed_excel\인케이스 리포트.xlsx', '구글애즈-스마트쇼핑')
# CID 및 캠페인 아이디 전달
googleads_shopping_api = GoogleAdsAPI('8370773952', 15687855740)
# DB 쿼리 실행 (batch api)
googleads_shopping_api.get_data()
# 데이터 가져오기 
googleAds_Perform = GoogleAdsPerform()
## Google Shopping
GoogleShopping_excel.google_ads_write(googleAds_Perform.pc_data, 58)
GoogleShopping_excel.google_ads_write(googleAds_Perform.mob_data, 93)
GoogleShopping_excel.google_ads_write(googleAds_Perform.tablet_data, 128)
GoogleShopping_excel.google_ads_write(googleAds_Perform.others_data, 163)
GoogleShopping_excel.save('static\completed_excel\인케이스 리포트.xlsx')
googleAds_Perform.clear_data()