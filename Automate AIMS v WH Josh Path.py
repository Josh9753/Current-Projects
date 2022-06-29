# Authors: John Ayres
# Desc: Script to compare AIMS system INV to Warehouse self-reported INV, institute Fail conditions, flag fails
# PARAM: WH ACTUALS xls, WH AIMS.xls
# OUT:Single XLSX file containing tabs for AIMS vs ACTUALS, and fails summary tab, which contains all failed Styles

import pandas as pd
import numpy as np
#import xlrd
#import openpyxl
#import xlsxwriter


######################################################## PATHS #########################################################

on7_inv_p = r'C:\Users\Joshua Kemperman\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\7-ON-Inv.xls'
on13_inv_p = r'C:\Users\Joshua Kemperman\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\13-ON-Inv.xls'
sv7_inv_p = r'C:\Users\Joshua Kemperman\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\7-SV-Inv.xls'
sv13_inv_p = r'C:\Users\Joshua Kemperman\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\13-SV-Inv.xls'

on_A_p = r'C:\Users\Joshua Kemperman\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\Inventory ON.xlsx'
sv_A_p = r'C:\Users\Joshua Kemperman\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\Inventory ON.xlsx'

#### Output: Change Date to Today!!!!!!!!!!!!!!!
save_Loc = r'C:\Users\Joshua Kemperman\OneDrive - Enchante Living\Documents\39 Joint Project\WH Stock vs AIMS stock 06.29.22.xlsx'


################################################## CREATE DATA FRAMES ##################################################

on7_inv = pd.read_excel(on7_inv_p,engine= 'xlrd')
on13_inv = pd.read_excel(on13_inv_p,engine= 'xlrd')
sv7_inv = pd.read_excel(sv7_inv_p,engine= 'xlrd')
sv13_inv = pd.read_excel(sv13_inv_p,engine= 'xlrd')

frame = [on7_inv, on13_inv]
on_AIMS = pd.concat(frame,ignore_index=True)

frame = [sv7_inv, sv13_inv]
sv_AIMS = pd.concat(frame,ignore_index=True)

on_ACTUAL = pd.read_excel(on_A_p,engine= 'openpyxl')
sv_ACTUAL = pd.read_excel(sv_A_p,engine= 'openpyxl')

##################################################### Changes ##########################################################
# Creating WH col
on_AIMS["WH"] = 'ON'
sv_AIMS["WH"] = 'SV'

# Rename units case cubic
on_AIMS['unit']=on_AIMS['on']
del on_AIMS['on']
sv_AIMS['unit']=sv_AIMS['sv']
del sv_AIMS['sv']

on_AIMS['caseqty']=on_AIMS['on_caseqty']
del on_AIMS['on_caseqty']
sv_AIMS['caseqty']=sv_AIMS['sv_caseqty']
del sv_AIMS['sv_caseqty']

on_AIMS['cubic']=on_AIMS['on_cubic']
del on_AIMS['on_cubic']
sv_AIMS['cubic']=sv_AIMS['sv_cubic']
del sv_AIMS['sv_cubic']

on_ACTUAL['Sty_Color'] = on_ACTUAL['Style']
del on_ACTUAL['Style']
sv_ACTUAL['Sty_Color'] = sv_ACTUAL['Style']
del sv_ACTUAL['Style']

on_AIMS = on_AIMS[['WH','style','grp','desc','color','division','cubic_ft','weight','master_pack','unit',
                               'caseqty', 'cubic']]
sv_AIMS = sv_AIMS[['WH','style','grp','desc','color','division','cubic_ft','weight','master_pack','unit',
                               'caseqty', 'cubic']]

# Creating Sty_col column
on_AIMS["Sty_Color"] = on_AIMS["style"] + "_" + on_AIMS["color"]
sv_AIMS["Sty_Color"] = sv_AIMS["style"] + "_" + sv_AIMS["color"]

# Creating WH QTY INFO COL
on_AIMS_master =  pd.merge(on_AIMS,on_ACTUAL,on='Sty_Color',how='inner')
sv_AIMS_master =  pd.merge(sv_AIMS,sv_ACTUAL,on='Sty_Color',how='inner')

#rename
on_AIMS_master['WH Qty Info'] = on_AIMS_master['Available']
del on_AIMS_master['Available']
sv_AIMS_master['WH Qty Info'] = sv_AIMS_master['Available']
del sv_AIMS_master['Available']

on_AIMS_master = on_AIMS_master[['WH','style','grp','desc','color','division','cubic_ft','weight','master_pack','unit',
                               'caseqty', 'cubic', 'Sty_Color','WH Qty Info']]
sv_AIMS_master = sv_AIMS_master[['WH','style','grp','desc','color','division','cubic_ft','weight','master_pack','unit',
                               'caseqty', 'cubic', 'Sty_Color','WH Qty Info']]

on_AIMS_master['unit'] = on_AIMS_master['unit'].astype(float)
sv_AIMS_master['unit'] = sv_AIMS_master['unit'].astype(float)
on_AIMS_master['WH Qty Info'] = on_AIMS_master['WH Qty Info'].astype(float)
sv_AIMS_master['WH Qty Info'] = sv_AIMS_master['WH Qty Info'].astype(float)
# difference col to create fails
on_AIMS_master['Diff'] = on_AIMS_master["unit"] - on_AIMS_master["WH Qty Info"]
sv_AIMS_master['Diff'] = sv_AIMS_master["unit"] - sv_AIMS_master["WH Qty Info"]


# Data type conversion
on_AIMS_master['Diff'] = on_AIMS_master['Diff'].astype(float)
on_AIMS_master['Diff'] = on_AIMS_master['Diff'].abs()
sv_AIMS_master['Diff'] = sv_AIMS_master['Diff'].astype(float)
sv_AIMS_master['Diff'] = sv_AIMS_master['Diff'].abs()
# Create Fail Col
conditions =[(on_AIMS_master['Diff']>=10),(on_AIMS_master['Diff']<10)]
values = ['Fail','Pass']
on_AIMS_master['Qty Close'] = np.select(conditions,values)
conditions =[(sv_AIMS_master['Diff']>=10),(sv_AIMS_master['Diff']<10)]
values = ['Fail','Pass']
sv_AIMS_master['Qty Close'] = np.select(conditions,values)

# creating  Fails summ
frame = [on_AIMS_master,sv_AIMS_master]
fail_sum = pd.concat(frame,ignore_index=True)
fail_sum = fail_sum.drop(fail_sum[fail_sum['Qty Close']=='Pass'].index)

# Cleaning Fails Sum
fail_sum = fail_sum[['Sty_Color','WH','unit','WH Qty Info','Diff']]
fail_sum['Style & Color'] = fail_sum['Sty_Color']
del fail_sum['Sty_Color']
fail_sum['AIMS Stock'] = fail_sum['unit']
del fail_sum['unit']
fail_sum['Difference'] = fail_sum['Diff']
del fail_sum['Diff']
fail_sum = fail_sum[['Style & Color','WH','AIMS Stock','WH Qty Info','Difference']]

##################################################### FILE OUTPUTS #####################################################

fileName = pd.ExcelWriter(save_Loc, engine = 'xlsxwriter')

# Sales OUT
on_AIMS_master.to_excel(fileName, sheet_name='ON AIMS Inv', index = False)
sv_AIMS_master.to_excel(fileName, sheet_name='SV AIMS Inv', index = False)

#INV OUT
on_ACTUAL.to_excel(fileName, sheet_name='ON WH INV', index = False)
sv_ACTUAL.to_excel(fileName, sheet_name='SV WH INV', index = False)

#SUMM OUT
fail_sum.to_excel(fileName, sheet_name='Fail Sum', index = False)
fileName.save()
