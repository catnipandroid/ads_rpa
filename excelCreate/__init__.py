import openpyxl

class Create_Excel:
    
    def __init__(self, wb):
        self.wb = openpyxl.load_workbook(wb)

    def naver_sa_write(self, sheet, data, cellNo):
        
        selected_sheet = self.wb[sheet]
        
        for idx, i in enumerate(data):
            imps_data = data[idx]['data'][0]['impCnt']
            clicks_data = data[idx]['data'][0]['clkCnt']
            cost_data = data[idx]['data'][0]['salesAmt']
            ccnt_data = data[idx]['data'][0]['ccnt']
            convAmt_data = data[idx]['data'][0]['convAmt']

            selected_sheet['C'+str(cellNo)] = imps_data
            selected_sheet['D'+str(cellNo)] = clicks_data
            selected_sheet['G'+str(cellNo)] = cost_data
            selected_sheet['H'+str(cellNo)] = ccnt_data
            selected_sheet['K'+str(cellNo)] = convAmt_data

            cellNo += 1  
            
    def google_ads_write(self, sheet, data, cellNo):
        
        selected_sheet = self.wb[sheet]
        
        for idx, i in enumerate(data['impCnt']):
            imps_data = data['impCnt'][idx]
            clicks_data = data['clkCnt'][idx]
            cost_data = data['salesAmt'][idx] * data['clkCnt'][idx]
            ccnt_data = data['ccnt'][idx]
            convAmt_data = data['convAmt'][idx]

            selected_sheet['C'+str(cellNo)] = imps_data
            selected_sheet['D'+str(cellNo)] = clicks_data
            selected_sheet['N'+str(cellNo)] = cost_data
            selected_sheet['H'+str(cellNo)] = ccnt_data
            selected_sheet['K'+str(cellNo)] = convAmt_data

            cellNo += 1  

    def save(self, report_name):
        self.wb.save(report_name)
        self.wb.close()
        