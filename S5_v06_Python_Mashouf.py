# Course   : Data Science with Python
# Teacher  : Mr. Pouriya Baghdadi
# Student  : Mohammadreza Mashouf
# Session  : 5


import numpy as np
import pandas as pd


path = 'C:\Mohammad\Learning\Python\Python-P.Baghdadi-My Course\Session_5\S5_Assignment\customerfeedback.xlsx'
df_CustomerTable   =    pd.DataFrame(pd.read_excel (path,sheet_name=0))
df_SupportTicket   =    pd.DataFrame(pd.read_excel (path,sheet_name=1))
df_Device          =    pd.DataFrame(pd.read_excel (path,sheet_name=2))
df_Company         =    pd.DataFrame(pd.read_excel (path,sheet_name=3))
df_Subscription    =    pd.DataFrame(pd.read_excel (path,sheet_name=4))
df_Geography       =    pd.DataFrame(pd.read_excel (path,sheet_name=5))
df_Role            =    pd.DataFrame(pd.read_excel (path,sheet_name=6))

'*****************************************************************************************************************************'

# Question : 01
original_score_avg = int(df_CustomerTable['Original Score'].mean())
df_CustomerTable['Original Score'] = df_CustomerTable['Original Score'].fillna(original_score_avg)               # Method 1
#df_CustomerTable['Original Score'] = df_CustomerTable['Original Score'].replace(np.nan, original_score_avg      # Method 2
'*****************************************************************************************************************************'

# Question : 02
df_CustomerTable = df_CustomerTable.dropna(subset=['Customer ID'])
#print(df_CustomerTable.head(10))
'*****************************************************************************************************************************'

# Question : 03
df_CustomerTable = df_CustomerTable.merge(
    df_Geography,
    how='left',
    left_on='Geography ID',
    right_on='Geography ID'
    )
'*****************************************************************************************************************************'

# Question : 04
df_CustomerTable_pivot = df_CustomerTable.pivot_table (
    index   = 'Country-Region',
    columns = 'Completed tutorial',
    values  = 'Original Score',
    aggfunc = np.sum
    )
'*****************************************************************************************************************************'

# Question : 05
df_CustomerTable['Completed tutorial'] = df_CustomerTable['Completed tutorial'].replace({'no': 0, 'yes': 1})
'*****************************************************************************************************************************'

# Question : 06
print ('Question 06 : \n' +
       'In total ' + 
       str('{:,.0f}'.format(len(df_SupportTicket.index))) +
       ' tickets are recorded.' +
       '\n')
'*****************************************************************************************************************************'

# Question : 07
df_CustomerTable = df_CustomerTable.merge(
    df_Role,
    how='left',
    left_on='Role',
    right_on='Role ID'
    )
df_CustomerTable_filtered = df_CustomerTable.loc[
    (df_CustomerTable['Role in Org']=='administrator') &
    (df_CustomerTable['Original Score'] == 10)
    ]   # Method 1

# df_CustomerTable_filtered = np.where(
#    (df_CustomerTable['Role in Org']=='administrator') &
#    (df_CustomerTable['Original Score'] == 10)
#    )  # Method 2

print ('Question 07 : \n' +
       'There are ' +
       str('{:,.0f}'.format(len(df_CustomerTable_filtered.index))) +
       ' customers with administrator role, which have recorded the 10 score.'+
       '\n')
'*****************************************************************************************************************************'

# Question : 08
df_CustomerTable_pivot_Continent = df_CustomerTable.pivot_table (
    index   = 'Continent',
    values  = 'Original Score',
    aggfunc = np.mean
        )
df_CustomerTable_pivot_Country = df_CustomerTable.pivot_table (
    index   = 'Country-Region',
    values  = 'Original Score',
    aggfunc = np.mean
        )
df_CustomerTable_pivot_Continent.plot.bar()
df_CustomerTable_pivot_Country.plot.bar()
'*****************************************************************************************************************************'

# Question : 09
Customer_Total = df_Device['Customer ID'].nunique()
# Customer_Total = len(df_Device.groupby(['Customer ID']).count()) # Method 2
Customer_Tablet = len(df_Device.loc[(df_Device['device']=='tablet')].index)
Customer_Tablet_percentage = Customer_Tablet / Customer_Total * 100
print ('Question 09 : \n' +
       str('{:,.0f}'.format(Customer_Tablet))  +
        ' customers, out of ' +
        str('{:,.0f}'.format(Customer_Total))   +
        ' have used tablet, which means about ' +
        str('{:,.0f}'.format(Customer_Tablet_percentage)) +
        '%.'+
       '\n')
'*****************************************************************************************************************************'

# Question : 10
df_Subscription = df_Subscription.drop(['Subscription New'], axis=1)
'*****************************************************************************************************************************'

# Question : 11
Company_Type = pd.DataFrame({'Company Type': ['Small','Medium','Large'] ,'Company ID': [1,2,3]})
df_Company = df_Company.merge(
    Company_Type,
    how='left',
    left_on='Company ID',
    right_on='Company ID'
    )
'*****************************************************************************************************************************'

# Question : 12
writer = pd.ExcelWriter('Customer Table_cleaned.xlsx')
df_CustomerTable.to_excel(writer)
writer.save()
print('Question 12 : \n' +
      'Customer Table DataFrame is written successfully to an Excel File, named : "Customer Table_cleaned" .')