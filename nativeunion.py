from NaverSA import NaverSA_API
from excelCreate import Create_Excel

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

# Excel Select
brandsa_excel = Create_Excel('static\exist_excel\네이티브유니온 리포트.xlsx', '네이버 브랜드검색')

# 브랜드검색
brandsa_excel.naver_sa_write(nu_brand_search_pc_data, 56)
brandsa_excel.naver_sa_write(nu_brand_search_mob_data, 91)
brandsa_excel.save('static\completed_excel/네이티브유니온 리포트.xlsx')

# 파워링크
powerlink_excel = Create_Excel('static\completed_excel/네이티브유니온 리포트.xlsx', '파워링크')
powerlink_excel.naver_sa_write(nu_powerlink_pc_data, 56)
powerlink_excel.naver_sa_write(nu_powerlink_mob_data, 91)
powerlink_excel.save('static\completed_excel/네이티브유니온 리포트.xlsx')



################### ################### ################### 
################### Google Ads 캠페인별 ###################
################### ################### ################### 

## Google SA (Excel파일명, 시트명)
GoogleSA_excel = Create_Excel('static\completed_excel\네이티브유니온 리포트.xlsx', '구글애즈SA')
# CID(Str) 및 캠페인아이디(Int) 전달
googleads_da_api = GoogleAdsAPI('7041360197', 12305055279)
# DB 쿼리 실행 (batch api)
googleads_da_api.get_data()
# 데이터 가져오기 
googleAds_Perform = GoogleAdsPerform()
## Google SA
GoogleSA_excel.google_ads_write(googleAds_Perform.pc_data, 58)
GoogleSA_excel.google_ads_write(googleAds_Perform.mob_data, 93)
GoogleSA_excel.google_ads_write(googleAds_Perform.tablet_data, 128)
GoogleSA_excel.google_ads_write(googleAds_Perform.others_data, 163)
GoogleSA_excel.save('static\completed_excel\네이티브유니온 리포트.xlsx')
googleAds_Perform.clear_data()


## Google Shopping (Excel파일명, 시트명)
GoogleShopping_excel = Create_Excel('static\completed_excel\네이티브유니온 리포트.xlsx', '구글애즈 쇼핑')
# CID(Str) 및 캠페인아이디(Int) 전달
googleads_da_api = GoogleAdsAPI('7041360197', 15660552908)
# DB 쿼리 실행 (batch api)
googleads_da_api.get_data()
# 데이터 가져오기 
googleAds_Perform = GoogleAdsPerform()
## Google Shopping
GoogleShopping_excel.google_ads_write(googleAds_Perform.pc_data, 58)
GoogleShopping_excel.google_ads_write(googleAds_Perform.mob_data, 93)
GoogleShopping_excel.google_ads_write(googleAds_Perform.tablet_data, 128)
GoogleShopping_excel.google_ads_write(googleAds_Perform.others_data, 163)
GoogleShopping_excel.save('static\completed_excel\네이티브유니온 리포트.xlsx')
googleAds_Perform.clear_data()


