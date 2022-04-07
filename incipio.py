from NaverSA import NaverSA_API
from excelCreate import Create_Excel
from GoogleAds import GoogleAdsAPI, GoogleAdsPerform
import configparser

####### 엑셀은 한번만 불러와서 하나의 변수에만 담아서 사용할 것! #############
####### 엑셀 SAVE는 한번만!! #######

config = configparser.ConfigParser()
config.read('config\info.ini')

################### ################### ################### 
################### Naver Search Ads 캠페인별 #############
################### ################### ################### 

# Excel
incipio_wb = Create_Excel(r'static\exist_excel\인시피오 리포트.xlsx')

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

# 파워링크
incipio_wb.naver_sa_write('파워링크', incipio_powerlink_pc_data, 56)
incipio_wb.naver_sa_write('파워링크', incipio_powerlink_mob_data, 91)


################### ################### ################### 
################### Google Ads 캠페인별 ###################
################### ################### ################### 

## Google DA
# CID 및 캠페인 아이디 전달
googleads_da_api = GoogleAdsAPI('1821610713', 15559749119)
# DB 쿼리 실행 (batch api)
googleads_da_api.get_data()
# 데이터 가져오기 
googleAds_Perform = GoogleAdsPerform()
## Google DA
incipio_wb.google_ads_write('구글애즈DA', googleAds_Perform.pc_data, 58)
incipio_wb.google_ads_write('구글애즈DA', googleAds_Perform.mob_data, 93)
incipio_wb.google_ads_write('구글애즈DA', googleAds_Perform.tablet_data, 128)
incipio_wb.google_ads_write('구글애즈DA', googleAds_Perform.others_data, 163)
googleAds_Perform.clear_data()


# 저장
incipio_wb.save('static\completed_excel/인시피오 리포트.xlsx')