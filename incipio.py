from NaverSA import NaverSA_API
from excelCreate import Create_Excel


################### ################### ################### 
################### Naver Search Ads 캠페인별 #############
################### ################### ################### 

# Day by Day Data를 담을 배열
incipio_powerlink_pc_data = []
incipio_powerlink_mob_data = []

# 파워링크 캠페인
incipioPlAdCampaign = {
    'nccCampaignIdPC': 'cmp-a001-01-000000003193310',
    'nccCampaignIdMob': 'cmp-a001-01-000000003193324',
}

# config 파일 필요함
naver_sa_api = NaverSA_API(config['naver_sa']['multipop_API_KEY'], config['naver_sa']['multipop_SECRET_KEY'], config['naver_sa']['multipop_CUSTOMER_ID'])

# 파워링크 호출
naver_sa_api.naverSearch_API_Get(incipioPlAdCampaign['nccCampaignIdPC'], incipio_powerlink_pc_data)
naver_sa_api.naverSearch_API_Get(incipioPlAdCampaign['nccCampaignIdMob'], incipio_powerlink_mob_data)

# Excel
powerlink_excel = Create_Excel('static\exist_excel\인시피오 리포트.xlsx', '파워링크')
powerlink_excel.naver_sa_write(incase_brand_search_pc_data, 56)
powerlink_excel.naver_sa_write(incase_brand_search_mob_data, 91)
powerlink_excel.save('static\completed_excel/인시피오 리포트.xlsx')


################### ################### ################### 
################### Google Ads 캠페인별 ###################
################### ################### ################### 

## Google DA
GoogleDA_excel = Create_Excel('static\completed_excel\인시피오 리포트.xlsx', '구글애즈DA')
# CID 및 캠페인 아이디 전달
googleads_da_api = GoogleAdsAPI('1821610713', 15559749119)
# DB 쿼리 실행 (batch api)
googleads_da_api.get_data()
# 데이터 가져오기 
googleAds_Perform = GoogleAdsPerform()
## Google DA
GoogleDA_excel.google_ads_write(googleAds_Perform.pc_data, 58)
GoogleDA_excel.google_ads_write(googleAds_Perform.mob_data, 93)
GoogleDA_excel.google_ads_write(googleAds_Perform.tablet_data, 128)
GoogleDA_excel.google_ads_write(googleAds_Perform.others_data, 163)
GoogleDA_excel.save('static\completed_excel\인시피오 리포트.xlsx')
googleAds_Perform.clear_data()