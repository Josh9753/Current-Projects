# Authors: John Ayres
# Desc: Script to compare AIMS system INV to Warehouse self-reported INV, institute Fail conditions, flag fails
# PARAM: WH ACTUALS xls, WH AIMS.xls
# OUT:Single XLSX file containing tabs for AIMS vs ACTUALS, and fails summary tab, which contains all failed Styles

import pandas as pd
import numpy as np

######################################################## PATHS #########################################################

on7_inv_p = r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\7-ON-Inv.xls'
on13_inv_p = r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\13-ON-Inv.xls'
sv7_inv_p = r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\7-SV-Inv.xls'
sv13_inv_p = r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\13-SV-Inv.xls'

on_A_p = r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\Inventory ON.xlsx'
sv_A_p = r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\Inventory SV.xlsx'

oo7p = r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\7 - Open Order.csv'
oo13p = r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\13 - Open Order.csv'
#### Output: Change Date to Today!!!!!!!!!!!!!!!
save_Loc = r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\WH Stock vs AIMS stock TODAY.xlsx'

# Difference Tolerance in units(Fail tolerance)
tol = 10

################################################## CREATE DATA FRAMES ##################################################

on7_inv = pd.read_excel(on7_inv_p, engine='xlrd')
on13_inv = pd.read_excel(on13_inv_p, engine='xlrd')
sv7_inv = pd.read_excel(sv7_inv_p, engine='xlrd')
sv13_inv = pd.read_excel(sv13_inv_p, engine='xlrd')

frame = [on7_inv, on13_inv]
on_AIMS = pd.concat(frame, ignore_index=True)

frame = [sv7_inv, sv13_inv]
sv_AIMS = pd.concat(frame, ignore_index=True)

on_ACTUAL = pd.read_excel(on_A_p, engine='openpyxl')
sv_ACTUAL = pd.read_excel(sv_A_p, engine='openpyxl')

oo7 = pd.read_csv(oo7p)
oo13 = pd.read_csv(oo13p)
frame = [oo7, oo13]
oo = pd.concat(frame, ignore_index=True)

##################################################### Changes ##########################################################
# Creating WH col
on_AIMS["WH"] = 'ON'
sv_AIMS["WH"] = 'SV'

# Rename units case cubic
on_AIMS['AIMS Stock'] = on_AIMS['on']
del on_AIMS['on']
sv_AIMS['AIMS Stock'] = sv_AIMS['sv']
del sv_AIMS['sv']

on_AIMS['caseqty'] = on_AIMS['on_caseqty']
del on_AIMS['on_caseqty']
sv_AIMS['caseqty'] = sv_AIMS['sv_caseqty']
del sv_AIMS['sv_caseqty']

on_AIMS['cubic'] = on_AIMS['on_cubic']
del on_AIMS['on_cubic']
sv_AIMS['cubic'] = sv_AIMS['sv_cubic']
del sv_AIMS['sv_cubic']

on_ACTUAL['Sty_Color'] = on_ACTUAL['Style']
del on_ACTUAL['Style']
sv_ACTUAL['Sty_Color'] = sv_ACTUAL['Style']
del sv_ACTUAL['Style']

on_AIMS = on_AIMS[['WH', 'style', 'grp', 'desc', 'color', 'division', 'cubic_ft', 'weight', 'master_pack', 'AIMS Stock',
                   'caseqty', 'cubic']]
sv_AIMS = sv_AIMS[['WH', 'style', 'grp', 'desc', 'color', 'division', 'cubic_ft', 'weight', 'master_pack', 'AIMS Stock',
                   'caseqty', 'cubic']]

# Creating Sty_col column
on_AIMS["Sty_Color"] = on_AIMS["style"] + "_" + on_AIMS["color"]
sv_AIMS["Sty_Color"] = sv_AIMS["style"] + "_" + sv_AIMS["color"]

# Creating WH QTY INFO COL
on_AIMS_master = pd.merge(on_AIMS, on_ACTUAL, on='Sty_Color', how='left')
sv_AIMS_master = pd.merge(sv_AIMS, sv_ACTUAL, on='Sty_Color', how='left')

# rename
on_AIMS_master['WH Stock'] = on_AIMS_master['Available']
del on_AIMS_master['Available']
sv_AIMS_master['WH Stock'] = sv_AIMS_master['Available']
del sv_AIMS_master['Available']

on_AIMS_master['description'] = on_AIMS_master['desc']
del on_AIMS_master['desc']
sv_AIMS_master['description'] = sv_AIMS_master['desc']
del sv_AIMS_master['desc']

on_AIMS_master['group'] = on_AIMS_master['grp']
del on_AIMS_master['grp']
sv_AIMS_master['group'] = sv_AIMS_master['grp']
del sv_AIMS_master['grp']

on_ACTUAL['Description'] = on_ACTUAL['Descr']
del on_ACTUAL['Descr']
sv_ACTUAL['Description'] = sv_ACTUAL['Descr']
del sv_ACTUAL['Descr']

on_AIMS_master = on_AIMS_master[['WH', 'style', 'group', 'description', 'color', 'division', 'cubic_ft', 'weight',
                                 'master_pack', 'caseqty', 'cubic', 'Sty_Color', 'AIMS Stock', 'WH Stock']]
sv_AIMS_master = sv_AIMS_master[['WH', 'style', 'group', 'description', 'color', 'division', 'cubic_ft', 'weight',
                                 'master_pack', 'caseqty', 'cubic', 'Sty_Color', 'AIMS Stock', 'WH Stock']]

# Handling nulls
on_AIMS_master["WH Stock"] = on_AIMS_master["WH Stock"].fillna(0)
sv_AIMS_master["WH Stock"] = sv_AIMS_master["WH Stock"].fillna(0)

on_AIMS_master['AIMS Stock'] = on_AIMS_master['AIMS Stock'].astype(float)
sv_AIMS_master['AIMS Stock'] = sv_AIMS_master['AIMS Stock'].astype(float)
on_AIMS_master['WH Stock'] = on_AIMS_master['WH Stock'].astype(float)
sv_AIMS_master['WH Stock'] = sv_AIMS_master['WH Stock'].astype(float)


####################################### Method to remove Open Orders from stock ########################################

oo["Sty_Color"] = oo["style"] + "_" + oo["color"]
oo_ON = oo[oo["loc"] == "ON"]
oo_SV = oo[oo["loc"] == "SV"]
oo_ON.reindex()
oo_SV.reindex()

#ON Logic
dupON = oo_ON[oo_ON["Sty_Color"].duplicated(keep='first')]
oo_ON = oo_ON.drop_duplicates(subset="Sty_Color")
dupON.reindex()
oo_ON.reindex()

totOO = pd.DataFrame()
dupON["pikqty2"] = dupON["pikqty"]
dupON = dupON[["Sty_Color", "pikqty2"]]
oo_ON = oo_ON[["Sty_Color", "pikqty"]]

for index, row in enumerate(oo_ON.itertuples(), 0):
    temp = oo_ON.iloc[index].loc["pikqty"]
    for index2, row2 in enumerate(dupON.itertuples(), 0):
        if oo_ON.iloc[index].loc["Sty_Color"] == dupON.iloc[index2].loc["Sty_Color"]:
            temp = dupON.iloc[index2].loc["pikqty2"] + temp
    size = len(totOO.index)
    totOO.loc[size,"pikqtyTOT"] = temp
    totOO.loc[size, "Sty_Color"] = oo_ON.iloc[index].loc["Sty_Color"]

totOO.reset_index()


on_AIMS_master = pd.merge(on_AIMS_master, totOO, how='left', on="Sty_Color")
on_AIMS_master["pikqtyTOT"] = on_AIMS_master["pikqtyTOT"].fillna(0)
on_AIMS_master["Adjusted Stock"] = on_AIMS_master["AIMS Stock"] - on_AIMS_master["pikqtyTOT"]
on_AIMS_master = on_AIMS_master[['WH', 'style', 'group', 'description', 'color', 'division', 'cubic_ft', 'weight',
                                 'master_pack', 'caseqty', 'cubic', 'Sty_Color', 'AIMS Stock', 'Adjusted Stock', 'WH Stock']]

#SV Logic
dupSV = oo_SV[oo_SV["Sty_Color"].duplicated(keep='first')]
oo_SV = oo_SV.drop_duplicates(subset="Sty_Color")
dupSV.reindex()
oo_SV.reindex()

totOOSV = pd.DataFrame()
dupSV["pikqty2"] = dupSV["pikqty"]
dupSV = dupSV[["Sty_Color", "pikqty2"]]
oo_SV = oo_SV[["Sty_Color", "pikqty"]]

for index, row in enumerate(oo_SV.itertuples(), 0):
    temp = oo_SV.iloc[index].loc["pikqty"]
    for index2, row2 in enumerate(dupSV.itertuples(), 0):
        if oo_SV.iloc[index].loc["Sty_Color"] == dupSV.iloc[index2].loc["Sty_Color"]:
            temp = dupSV.iloc[index2].loc["pikqty2"] + temp
    size = len(totOOSV.index)
    totOOSV.loc[size,"pikqtyTOT"] = temp
    totOOSV.loc[size, "Sty_Color"] = oo_SV.iloc[index].loc["Sty_Color"]

totOOSV.reset_index()


sv_AIMS_master = pd.merge(sv_AIMS_master, totOOSV, how='left', on="Sty_Color")
sv_AIMS_master["pikqtyTOT"] = sv_AIMS_master["pikqtyTOT"].fillna(0)
sv_AIMS_master["Adjusted Stock"] = sv_AIMS_master["AIMS Stock"] - sv_AIMS_master["pikqtyTOT"]
sv_AIMS_master = sv_AIMS_master[['WH', 'style', 'group', 'description', 'color', 'division', 'cubic_ft', 'weight',
                                 'master_pack', 'caseqty', 'cubic', 'Sty_Color', 'AIMS Stock', 'Adjusted Stock', 'WH Stock']]


# difference col to create fails
on_AIMS_master['Diff'] = on_AIMS_master["Adjusted Stock"] - on_AIMS_master["WH Stock"]
sv_AIMS_master['Diff'] = sv_AIMS_master["Adjusted Stock"] - sv_AIMS_master["WH Stock"]

# Data type conversion
on_AIMS_master['Diff'] = on_AIMS_master['Diff'].astype(float)
on_AIMS_master['Diff'] = on_AIMS_master['Diff'].abs()
sv_AIMS_master['Diff'] = sv_AIMS_master['Diff'].astype(float)
sv_AIMS_master['Diff'] = sv_AIMS_master['Diff'].abs()


# Create Fail Col
conditions = [(on_AIMS_master['Diff'] >= tol), (on_AIMS_master['Diff'] < tol), (on_AIMS_master['Diff'] == None)]
values = ['Fail', 'Pass', 'Pass']
on_AIMS_master['Qty Close'] = np.select(conditions, values, default=0)
conditions = [(sv_AIMS_master['Diff'] >= tol), (sv_AIMS_master['Diff'] < tol), (sv_AIMS_master['Diff'] == None)]
values = ['Fail', 'Pass', 'Pass']
sv_AIMS_master['Qty Close'] = np.select(conditions, values, default=0)

# creating  Fails sum
frame = [on_AIMS_master, sv_AIMS_master]
fail_sum = pd.concat(frame, ignore_index=True)
fail_sum = fail_sum.loc[fail_sum['Qty Close'] == 'Fail']

# Cleaning Fails Sum
fail_sum = fail_sum[['Sty_Color', 'WH', 'AIMS Stock', 'WH Stock', 'Diff']]
fail_sum['Style & Color'] = fail_sum['Sty_Color']
del fail_sum['Sty_Color']
fail_sum['Difference'] = fail_sum['Diff']
del fail_sum['Diff']
fail_sum = fail_sum[['Style & Color', 'WH', 'AIMS Stock', 'WH Stock', 'Difference']]

# Cleaning data
del on_ACTUAL['Sty_Color']
del sv_ACTUAL['Sty_Color']
del on_ACTUAL['Code']
del sv_ACTUAL['Code']
del on_ACTUAL['NMFC']
del sv_ACTUAL['NMFC']
on_ACTUAL['WH Stock'] = on_ACTUAL['Qty']
del on_ACTUAL['Qty']
sv_ACTUAL['WH Stock'] = sv_ACTUAL['Qty']
del sv_ACTUAL['Qty']
on_ACTUAL = on_ACTUAL[['Customer', 'Facility', 'Item', 'Description', 'Color', 'Size', 'WH Stock', 'Available',
                       'Case Qty', 'Length', 'Height', 'Width', 'Weight', 'Cube Each', 'CFT Each Per Line', 'Group',
                       'Date']]
sv_ACTUAL = sv_ACTUAL[['Customer', 'Facility', 'Item', 'Description', 'Color', 'Size', 'WH Stock', 'Available',
                       'Case Qty', 'Length', 'Height', 'Width', 'Weight', 'Cube Each', 'CFT Each Per Line', 'Group',
                       'Date']]

print(on_AIMS_master[on_AIMS_master["Sty_Color"]=="P212_CLR"],"WH Stock")
##################################################### FILE OUTPUTS #####################################################

fileName = pd.ExcelWriter(save_Loc, engine='xlsxwriter')

# Sales OUT
on_AIMS_master.to_excel(fileName, sheet_name='ON AIMS Inv', index=False)
sv_AIMS_master.to_excel(fileName, sheet_name='SV AIMS Inv', index=False)

# INV OUT
on_ACTUAL.to_excel(fileName, sheet_name='ON WH INV', index=False)
sv_ACTUAL.to_excel(fileName, sheet_name='SV WH INV', index=False)

# SUM OUT
fail_sum.to_excel(fileName, sheet_name='Fail Sum', index=False)


fileName.save()
