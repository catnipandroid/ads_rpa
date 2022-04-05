from NaverSA import NaverSA_API
from excelCreate import Create_Excel


################### ################### ################### 
################### Naver Search Ads 캠페인별 #############
################### ################### ################### 

# Day by Day Data를 담을 배열
griffin_powerlink_pc_data = []
griffin_powerlink_mob_data = []

griffinPlAdCampaign = {
    'nccCampaignIdPC': 'cmp-a001-01-000000003193151',
    'nccCampaignIdMob': 'cmp-a001-01-000000003193279',
}

# config 파일 필요함
naver_sa_api = NaverSA_API(config['naver_sa']['multipop_API_KEY'], config['naver_sa']['multipop_SECRET_KEY'], config['naver_sa']['multipop_CUSTOMER_ID'])


# 파워링크 호출
naver_sa_api.naverSearch_API_Get(griffinPlAdCampaign['nccCampaignIdPC'], griffin_powerlink_pc_data)
naver_sa_api.naverSearch_API_Get(griffinPlAdCampaign['nccCampaignIdMob'], griffin_powerlink_mob_data)

# Excel
powerlink_excel = Create_Excel('static\exist_excel\그리핀 리포트.xlsx', '파워링크')
powerlink_excel.naver_sa_write(griffin_powerlink_pc_data, 56)
powerlink_excel.naver_sa_write(griffin_powerlink_mob_data, 91)
powerlink_excel.save('static\completed_excel/그리핀 리포트.xlsx')


################### ################### ################### 
################### Google Ads 캠페인별 ###################
################### ################### ################### 

## Google Youtube
GoogleYoutube_excel = Create_Excel('static\completed_excel\그리핀 리포트.xlsx', '구글애즈 유튜브')
# CID 및 캠페인 아이디 전달
googleads_da_api = GoogleAdsAPI('2912045005', 16253502937)
# DB 쿼리 실행 (batch api)
googleads_da_api.get_data()
# 데이터 가져오기 
googleAds_Perform = GoogleAdsPerform()
## Google Youtube
GoogleYoutube_excel.google_ads_write(googleAds_Perform.pc_data, 58)
GoogleYoutube_excel.google_ads_write(googleAds_Perform.mob_data, 93)
GoogleYoutube_excel.google_ads_write(googleAds_Perform.tablet_data, 128)
GoogleYoutube_excel.google_ads_write(googleAds_Perform.others_data, 163)
GoogleYoutube_excel.save('static\completed_excel\그리핀 리포트.xlsx')
googleAds_Perform.clear_data()