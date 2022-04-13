from NaverSA import NaverSA_API
from excelCreate import Create_Excel
from GoogleAds import GoogleAdsAPI, GoogleAdsPerform
import configparser

config = configparser.ConfigParser()
config.read('config\info.ini')

################### ################### ################### 
################### Naver Search Ads 캠페인별 #############
################### ################### ################### 

# Excel
nu_wb = Create_Excel(r'static\exist_excel\네이티브유니온 리포트.xlsx')

# Day by Day Data를 담을 배열
nu_brand_search_pc_data = []
nu_brand_search_mob_data = []
nu_powerlink_pc_data = []
nu_powerlink_mob_data = []

# 브랜드검색 캠페인
nuBrandAdCampaign = {
    'nccCampaignIdPC': 'cmp-a001-04-000000003175217',
    'nccCampaignIdMob': 'cmp-a001-04-000000003175223',
}

# 파워링크 캠페인
nuPlAdCampaign = {
    'nccCampaignIdPC': 'cmp-a001-01-000000003192699',
    'nccCampaignIdMob': 'cmp-a001-01-000000003193097',
}


# config 파일 필요함
naver_sa_api = NaverSA_API(config['naver_sa']['multipop_API_KEY'], config['naver_sa']['multipop_SECRET_KEY'], config['naver_sa']['multipop_CUSTOMER_ID'])

# 브랜드검색 호출
naver_sa_api.naverSearch_API_Get(nuBrandAdCampaign['nccCampaignIdPC'], nu_brand_search_pc_data)
naver_sa_api.naverSearch_API_Get(nuBrandAdCampaign['nccCampaignIdMob'], nu_brand_search_mob_data)

# 파워링크 호출
naver_sa_api.naverSearch_API_Get(nuPlAdCampaign['nccCampaignIdPC'], nu_powerlink_pc_data)
naver_sa_api.naverSearch_API_Get(nuPlAdCampaign['nccCampaignIdMob'], nu_powerlink_mob_data)

# 브랜드검색
nu_wb.naver_brand_sa_write('네이버 브랜드검색', nu_brand_search_pc_data, 56, 18333)
nu_wb.naver_brand_sa_write('네이버 브랜드검색', nu_brand_search_mob_data, 91, 25667)

# 파워링크
nu_wb.naver_sa_write('파워링크', nu_powerlink_pc_data, 56)
nu_wb.naver_sa_write('파워링크', nu_powerlink_mob_data, 91)


################### ################### ################### 
################### Google Ads 캠페인별 ###################
################### ################### ################### 

## Google DA
# CID 및 캠페인 아이디 전달
googleads_da_api = GoogleAdsAPI('7041360197', 16792036770)
# DB 쿼리 실행 (batch api)
googleads_da_api.get_data()
# 데이터 가져오기 
googleAds_Perform = GoogleAdsPerform()
## Google DA
nu_wb.google_ads_write('구글애즈DA', googleAds_Perform.pc_data, 63)
nu_wb.google_ads_write('구글애즈DA', googleAds_Perform.mob_data, 98)
nu_wb.google_ads_write('구글애즈DA', googleAds_Perform.tablet_data, 133)
nu_wb.google_ads_write('구글애즈DA', googleAds_Perform.others_data, 168)
googleAds_Perform.clear_data()

## Google Search Ads
# CID 및 캠페인 아이디 전달
googleads_sa_api = GoogleAdsAPI('7041360197', 12305055279)
# DB 쿼리 실행 (batch api)
googleads_sa_api.get_data()
# 데이터 가져오기 
googleAds_Perform = GoogleAdsPerform()
## Google DA
nu_wb.google_ads_write('구글애즈SA', googleAds_Perform.pc_data, 58)
nu_wb.google_ads_write('구글애즈SA', googleAds_Perform.mob_data, 93)
nu_wb.google_ads_write('구글애즈SA', googleAds_Perform.tablet_data, 128)
nu_wb.google_ads_write('구글애즈SA', googleAds_Perform.others_data, 163)
googleAds_Perform.clear_data()


## Google Shopping
# CID 및 캠페인 아이디 전달
googleads_shopping_api = GoogleAdsAPI('7041360197', 15660552908)
# DB 쿼리 실행 (batch api)
googleads_shopping_api.get_data()
# 데이터 가져오기 
googleAds_Perform = GoogleAdsPerform()
## Google DA
nu_wb.google_ads_write('구글애즈 쇼핑', googleAds_Perform.pc_data, 58)
nu_wb.google_ads_write('구글애즈 쇼핑', googleAds_Perform.mob_data, 93)
nu_wb.google_ads_write('구글애즈 쇼핑', googleAds_Perform.tablet_data, 128)
nu_wb.google_ads_write('구글애즈 쇼핑', googleAds_Perform.others_data, 163)
googleAds_Perform.clear_data()


# 저장
nu_wb.save('static\completed_excel/네이티브유니온 리포트.xlsx')