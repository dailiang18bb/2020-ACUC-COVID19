# -*- coding: utf-8 -*-
from datetime import datetime
import pandas as pd
import numpy as np
import collections
import xlsxwriter







### Read Excel
excel_file_name = 'testExcel.xlsx'
df = pd.read_excel(excel_file_name, sheet_names='ACUCOrganizationDonationVerific')
# complete_df = df[df['Entry_Status'] == 'Complete']
complete_df = df.fillna(0) 
# complete_df['机构名称'].fillna('',inplace = True) 
num_rows = len(complete_df.index)


#Dict
translate_dict = {
'TotalPiecesOfN95InUS' : 'N95 Mask (Piece)',
'TotalPiecesOfSurgicalMaskInUS' : 'Medical Mask (Piece)',
'TotalPiecesOfMedicalProtectiveGownInUS' : 'Protective Garment',
'_3TotalPiecesOfTestKitsInUS' : 'Test Kits (set)',
'TotalPiecesOfMechanicalVentilator' : 'Ventilator',
'TotalPiecesOfHandSanitizerOrHandSoapInUS' : 'Hand Sanitizer (package)',
'TotalPiecesOfMedicalProtectiveHatInUS' : 'Protective Hat',
'TotalPiecesOfShoesCover' : 'Protective Shoes Cover',
'Gloves' : 'Protective Gloves',
'TotalPiecesOfGogglesInUS' : 'Goggles',
'TotalPiecesOfFaceShieldInUS' : 'Face Shield',
'DisinfectWipes' : 'Disinfect Wipes',
'TotalQtOfMeals' : 'Meals',
'_10TotalPiecesOfProtectiveKitsInUS' : 'ProtectiveKits',
# '_15OtherInKindSuppliesOrDonationsInUS' : 'Others'
}

summary_dict = {
'N95 Mask (Piece)' : 0,
'Medical Mask (Piece)' : 0,
'Protective Garment' : 0,
'Test Kits (set)' : 0,
'Ventilator' : 0,
'Hand Sanitizer (package)' : 0,
'Protective Hat' : 0,
'Protective Shoes Cover' : 0,
'Protective Gloves' : 0,
'Goggles' : 0,
'Face Shield' : 0,
'Disinfect Wipes' : 0,
'Meals' : 0,
'ProtectiveKits' : 0,
}

### Date
now = datetime.now().strftime('%m_%d_%Y %Hh_%Mm_%Ss')
update_date = datetime.now().strftime('%m/%d/%Y %H:%M:%S')


# Excel
# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('./Excel/ACUC Donation summary Excel light' + now + '.xlsx')
worksheet = workbook.add_worksheet()

# format
number_format = workbook.add_format({'num_format': '#,##0.00'})


# ### Write Word
# document = Document()
# # document.styles['Normal'].font.name = 'SimHei'
# document.styles['Normal'].font.name = 'SimHei'

# p = document.add_paragraph()
# p_run = p.add_run('ACUC Covid19 Donation Summary')
# p2= document.add_paragraph('Last Update: ' + str(update_date))
# p.alignment = WD_ALIGN_PARAGRAPH.CENTER
# p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
# p_run.font.size = Pt(24)



# table = document.add_table(rows=num_rows + 1, cols=1)
# table.style = 'Table Grid'

### Ranking and summary
rank_dict = {}
summary_amount = 0
summary_cash_amount = 0
summary_transfer_amount = 0 # 正在中国转运物资价值：

summary_amount_acuc_cash = 0 # 汇集ACUC现金捐赠
summary_amount_other_cash = 0 # 其余现金捐赠
summary_amount_supplies = 0 # 其他物资价值

summary_N95_Mask = 0
summary_Medical_Mask = 0
summary_Protective_Garment = 0
# summary_Test_Kits = 0
summary_Ventilator = 0
summary_Hand_Sanitizer = 0
summary_Protective_Hat = 0
summary_Protective_Shoes_Cover = 0
summary_Protective_Gloves = 0
summary_Goggles = 0
summary_Face_Shield = 0
summary_Disinfect_Wipes = 0
summary_Meals = 0
summary_ProtectiveKits = 0
summary_other = 0
# outside NY
summary_outside_ny = 0;


# Excel dict

excel_dict = {}




for index, row in complete_df.iterrows():

	# new dict
	row_dict = {}
	# Col1_ID
	row_dict['ID'] = str(int(row['OrganizationSignUpListNumber您的机构在接龙里的序号']))

	# Col2_OrgEnglishName
	row_dict['OrgEnglishName'] = row['OrganizationNameInEnglish']

	# Col3_OrgChineseName
	if row['机构名称'] != 0:
		row_dict['OrgChineseName'] = row['机构名称']
	else:
		row_dict['OrgChineseName'] = ''

	# Col _AUCU现金捐款
	row_dict['AUCU现金捐款'] = row['CashDonationAmountThroughACUC']
	summary_amount_acuc_cash += row['CashDonationAmountThroughACUC']

	# Col _直接现金捐款
	row_dict['直接现金捐款'] = row['AmountOfDonatedCash']
	summary_amount_other_cash += row['AmountOfDonatedCash']

	# Col_总现金捐款
	cash_amount = row['CashDonationAmountThroughACUC'] + row['AmountOfDonatedCash']
	row_dict['总现金捐款'] = cash_amount
	summary_cash_amount += cash_amount

	# Col _N95 Mask
	row_dict['N95 Mask'] = row['ValuePriceOfTotalN95']
	summary_N95_Mask += row['ValuePriceOfTotalN95']
	
	# Col _Medical Mask
	row_dict['Medical Mask'] = row['ValuePriceOfTotalSurgicalMask']
	summary_Medical_Mask += row['ValuePriceOfTotalSurgicalMask']

	# Col _Protective Garment
	row_dict['Protective Garment'] = row['ValuePriceOfTotalGown']
	summary_Protective_Garment += row['ValuePriceOfTotalGown']

	# Col _Test Kits
	# row_dict['Test Kits'] = row['ValuePriceOfTotalTestKits']
	# summary_Test_Kits += row['ValuePriceOfTotalTestKits']
	
	# Col _Ventilator
	row_dict['Ventilator'] = row['ValuePriceOfTotalVentilator']
	summary_Ventilator += row['ValuePriceOfTotalVentilator']

	# Col _ Hand Sanitizer
	row_dict['Hand Sanitizer'] = row['ValuePriceOfTotalHandSoapSanitizer']
	summary_Hand_Sanitizer += row['ValuePriceOfTotalHandSoapSanitizer']

	# Col _ Protective Hat
	row_dict['Protective Hat'] = row['ValuePriceOfTotalProtectiveHat']
	summary_Protective_Hat += row['ValuePriceOfTotalProtectiveHat']

	# Col _ Protective Shoes Cover
	row_dict['Protective Shoes Cover'] = row['ValuePriceOfTotalShoesCover']
	summary_Protective_Shoes_Cover += row['ValuePriceOfTotalShoesCover']

	# Col _ Protective Gloves
	row_dict['Protective Gloves'] = row['ValuePriceOfTotalGloves']
	summary_Protective_Gloves += row['ValuePriceOfTotalGloves']
	
	# Col _ Goggles
	row_dict['Goggles'] = row['ValuePriceOfTotalGoggles']
	summary_Goggles += row['ValuePriceOfTotalGoggles']

	# Col _ Face Shield
	row_dict['Face Shield'] = row['ValuePriceOfTotalFaceShield']
	summary_Face_Shield += row['ValuePriceOfTotalFaceShield']

	# Col _ Disinfect Wipes
	row_dict['Disinfect Wipes'] = row['ValuePriceOfTotalDisinfectWipes']
	summary_Disinfect_Wipes += row['ValuePriceOfTotalDisinfectWipes']

	# Col _ Meals
	row_dict['Meals'] = row['ValuePriceOfTotalMeals']
	summary_Meals += row['ValuePriceOfTotalMeals']

	# Col _ ProtectiveKits
	row_dict['ProtectiveKits'] = row['ValuePriceOfTotalProtectiveKits']
	summary_ProtectiveKits += row['ValuePriceOfTotalProtectiveKits']

	# Col _ Other
	row_dict['Other'] = row['TotalOfOthersValuePrice']
	summary_other += row['TotalOfOthersValuePrice']

	# Col_物品捐赠总价值
	supplie_value = row['ValuePriceOfTotalN95'] + \
	row['ValuePriceOfTotalSurgicalMask'] + row['ValuePriceOfTotalTestKits'] + \
	row['ValuePriceOfTotalGown'] + row['ValuePriceOfTotalVentilator'] + \
	row['ValuePriceOfTotalHandSoapSanitizer'] + row['ValuePriceOfTotalProtectiveHat'] + \
	row['ValuePriceOfTotalShoesCover'] + row['ValuePriceOfTotalFaceShield'] + \
	row['ValuePriceOfTotalProtectiveKits'] + row['ValuePriceOfTotalGloves'] + \
	row['ValuePriceOfTotalGoggles'] + row['ValuePriceOfTotalDisinfectWipes'] + \
	row['ValuePriceOfTotalMeals'] + row['TotalOfOthersValuePrice']
	summary_amount_supplies += supplie_value
	row_dict['物品捐赠总价值'] = supplie_value

	# Col_正在海外转运物资价值
	if row['ValuePriceOfYourPurchaseThatAreNotYetArrived'] != 0:
		row_dict['正在海外转运物资价值'] = row['ValuePriceOfYourPurchaseThatAreNotYetArrived']
		summary_transfer_amount += row['ValuePriceOfYourPurchaseThatAreNotYetArrived']
	else:
		row_dict['正在海外转运物资价值'] = 0

	# Col7_总计捐赠价值
	total_amount = cash_amount + supplie_value + row['ValuePriceOfYourPurchaseThatAreNotYetArrived']
	row_dict['总计捐赠价值'] = total_amount
	summary_amount += total_amount

	# special outside ny donation
	if row['OrganizationSignUpListNumber您的机构在接龙里的序号'] == 75 or row['OrganizationSignUpListNumber您的机构在接龙里的序号'] == 62:
		summary_outside_ny += total_amount * 0.6

	# add to dict
	excel_dict.update({ str(int(row['OrganizationSignUpListNumber您的机构在接龙里的序号'])) : row_dict} )



# write title

# for i in range(7):
worksheet.write(0, 0, 'ID')
worksheet.write(0, 1, 'OrgEnglishName')
worksheet.write(0, 2, 'OrgChineseName')

worksheet.write(0, 3, 'AUCU现金捐款')
worksheet.write(0, 4, '直接现金捐款')
worksheet.write(0, 5, '总现金捐款')

worksheet.write(0, 6, 'N95_Mask_Value')
worksheet.write(0, 7, 'Medical_Mask_Value')
worksheet.write(0, 8, 'Protective_Garment_Value')
worksheet.write(0, 9, 'Ventilator_Value')
worksheet.write(0, 10, 'Hand_Sanitizer_Value')
worksheet.write(0, 11, 'Protective_Hat_Value')
worksheet.write(0, 12, 'Protective_Shoes_Cover_Value')
worksheet.write(0, 13, 'Protective_Gloves_Value')
worksheet.write(0, 14, 'Goggles_Value')
worksheet.write(0, 15, 'Face_Shield_Value')
worksheet.write(0, 16, 'Disinfect_Wipes_Value')
worksheet.write(0, 17, 'Meals_Value')
worksheet.write(0, 18, 'ProtectiveKits_Value')
worksheet.write(0, 19, 'Other_Value')
worksheet.write(0, 20, '物品捐赠总价值')

worksheet.write(0, 21, '正在海外转运物资价值')
worksheet.write(0, 22, '总计捐赠价值')


i = 1
for key, value in excel_dict.items():
	x = 0
	for k, v in value.items():
		worksheet.write(i, x, v, number_format)
		x += 1
	i += 1

i += 1
worksheet.write(i, 1, '代码总计计算')
worksheet.write(i, 3, summary_amount_acuc_cash, number_format)
worksheet.write(i, 4, summary_amount_other_cash, number_format)
worksheet.write(i, 5, summary_cash_amount, number_format)

worksheet.write(i, 6, summary_N95_Mask, number_format)
worksheet.write(i, 7, summary_Medical_Mask, number_format)
worksheet.write(i, 8, summary_Protective_Garment, number_format)
worksheet.write(i, 9, summary_Ventilator, number_format)
worksheet.write(i, 10, summary_Hand_Sanitizer, number_format)
worksheet.write(i, 11, summary_Protective_Hat, number_format)
worksheet.write(i, 12, summary_Protective_Shoes_Cover, number_format)
worksheet.write(i, 13, summary_Protective_Gloves, number_format)
worksheet.write(i, 14, summary_Goggles, number_format)
worksheet.write(i, 15, summary_Face_Shield, number_format)
worksheet.write(i, 16, summary_Disinfect_Wipes, number_format)
worksheet.write(i, 17, summary_Meals, number_format)
worksheet.write(i, 18, summary_ProtectiveKits, number_format)
worksheet.write(i, 19, summary_other, number_format)
worksheet.write(i, 20, summary_amount_supplies, number_format)

worksheet.write(i, 21, summary_transfer_amount, number_format)
worksheet.write(i, 22, summary_amount, number_format)


# formula
i += 1

worksheet.write(i, 1, '公式总计验证')
for y in range(68, 88): # [D, E, F ... W]
	worksheet.write_formula(str(chr(y))+str(i+1), '=SUM(' +str(chr(y))+ '2:'+ str(chr(y)) + str(i - 2) + ')')


# special outside ny
i += 2
worksheet.write(i, 1, '总计捐款价值')
worksheet.write(i, 3, summary_amount, number_format)

i += 1
worksheet.write(i, 1, '捐赠给 NY Tri-state area 价值')
worksheet.write(i, 3, summary_amount - summary_outside_ny, number_format)



print('Excel file generate successful!')

workbook.close()


# document.save('./output/ACUC Donation Summary ' + now + '.docx')
# print('Word file generate successful!')



# 
# 1. 正在转运列改为数字
# 2. 改127号，[海外捐款]



# Style
# 1. Page Margin to Normal
# 2. Chinese font SimHei, English font Calibri
# 3. table not across page
# 4. table font size 12
# 5. table layout
# 6. table autofit



# 0418 Update Notes:
# 1. change [Surgical Mask] to [Medical Mask]
# 2. ignore row['_15OtherInKindSuppliesOrDonationsInUS'] '0' item


# Update Notes:
# 1. 所有物资数量使用千分符(,)显示
# 2. 在明细中，增加[海外转运捐赠价值]
# 3. 隐藏明细中[现金捐赠，物资捐赠，海外转运捐赠]为空的，对应标题。

# Notes： 
# 1. 同意中英文两行
# 2. 总捐助价值下上面，明细在下
# 3. 英文就可以，不需要翻译中文



# index
# total_value_price = ['ValuePriceOfTotalN95',
# 'ValuePriceOfTotalSurgicalMask',
# 'ValuePriceOfTotalTestKits',
# 'ValuePriceOfTotalGown',
# 'ValuePriceOfTotalVentilator',
# 'ValuePriceOfTotalHandSoapSanitizer',
# 'ValuePriceOfTotalProtectiveHat',
# 'ValuePriceOfTotalShoesCover',
# 'ValuePriceOfTotalFaceShield',
# 'ValuePriceOfTotalProtectiveKits',
# 'ValuePriceOfTotalGloves',
# 'ValuePriceOfTotalGoggles',
# 'ValuePriceOfTotalDisinfectWipes',
# 'ValuePriceOfTotalMeals',
# 'TotalValuePrice',
# 'ValuePriceOfYourPurchaseThatAreNotYetArrived']


###Update Notes
# 1. re-arrange sorting algorithm and Num.
# 2. cover all company
# 3. summary header
### Style
# 1. font SimHei
# 2. Table not across page
# 3. table horizontal line