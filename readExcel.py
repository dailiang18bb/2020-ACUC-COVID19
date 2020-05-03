import pandas as pd

excel_file_name = 'testExcel.xlsx'

df = pd.read_excel(excel_file_name, sheet_names='ACUCOrganizationDonationVerific')
complete_df = df[df['Entry_Status'] == 'Complete']
num_rows = len(complete_df.index)
complete_df = complete_df.fillna(0) 
# df = df.replace('nan', '')

# print (df)
# print (complete_df.head(5))

for index, row in complete_df.head(1).iterrows():
     # access data using column names
     first_row = '###' + '. ' + row['OrganizationNameInEnglish'] + '_' + row['机构名称']
     # print(index, first_row)
     row_part = row.iloc[12:58] 
     for i , v in row_part.iteritems():
     	print(i)
     	print(v)

# print (complete_df.dtypes)
