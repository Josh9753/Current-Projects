# Authors: John Ayres, Joshua Kemperman
# Desc: Script to process INV/Sales, and Summary containing all INV/Sales JOHN PATH IN PROGRESS, ADDING FUNCTIONALITY
# PARAM: Input files as listed under pathing; Warehouse INV/OHS by data group and INV Report
# OUTPUT: Multipage excel Document, listing total INV/OHS by warehouse, Sheets containing totals of INV/OHS

import pandas as pd

######################################################## PATHS #########################################################
si7_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-SI.xls'
si13_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-SI.xls'
on7_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-ON.xls'
on13_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-ON.xls'
sv7_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-SV.xls'
sv13_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-SV.xls'

si7_inv_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-SI-Inv.xls'
si13_inv_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-SI-Inv.xls'
on7_inv_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-ON-Inv.xls'
on13_inv_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-ON-Inv.xls'
sv7_inv_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-SV-Inv.xls'
sv13_inv_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-SV-Inv.xls'

iv_rep = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\05 - INVENTORY\Inventory Report - For Audit - June 20, 2022.xlsx'

#### CHANGE DATE TO TODAY
save_Loc = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\US, CA WHSE Sales and OH 06.23.22.xlsx'


################################################## CREATE DATA FRAMES ##################################################

si7 = pd.read_excel(si7_p, engine='xlrd')
si13 = pd.read_excel(si13_p, engine='xlrd')
frame = [si7, si13]
si_sales_master = pd.concat(frame, ignore_index=True)

on7 = pd.read_excel(on7_p, engine='xlrd')
on13 = pd.read_excel(on13_p, engine='xlrd')
frame = [on7, on13]
on_sales_master = pd.concat(frame, ignore_index=True)

sv7 = pd.read_excel(sv7_p, engine='xlrd')
sv13 = pd.read_excel(sv13_p, engine='xlrd')
frame = [sv7, sv13]
sv_sales_master = pd.concat(frame, ignore_index=True)

si7_inv = pd.read_excel(si7_inv_p, engine='xlrd')
si13_inv = pd.read_excel(si13_inv_p, engine='xlrd')
on7_inv = pd.read_excel(on7_inv_p, engine='xlrd')
on13_inv = pd.read_excel(on13_inv_p, engine='xlrd')
sv7_inv = pd.read_excel(sv7_inv_p, engine='xlrd')
sv13_inv = pd.read_excel(sv13_inv_p, engine='xlrd')

frame = [si7_inv, si13_inv]
si_inv_master = pd.concat(frame, ignore_index=True)

frame = [on7_inv, on13_inv]
on_inv_master = pd.concat(frame, ignore_index=True)

frame = [sv7_inv, sv13_inv]
sv_inv_master = pd.concat(frame, ignore_index=True)

inv_rep = pd.read_excel(iv_rep, header=2)

##################################################### Changes ##########################################################

# Creating Month Col
si_sales_master['invdate'] = pd.to_datetime(si_sales_master['invdate'])
si_sales_master['month'] = si_sales_master['invdate'].dt.strftime('%m')
on_sales_master['invdate'] = pd.to_datetime(on_sales_master['invdate'])
on_sales_master['month'] = on_sales_master['invdate'].dt.strftime('%m')
sv_sales_master['invdate'] = pd.to_datetime(sv_sales_master['invdate'])
sv_sales_master['month'] = sv_sales_master['invdate'].dt.strftime('%m')

# Creating WH col
si_sales_master["WH"] = 'SI'
on_sales_master["WH"] = 'ON'
sv_sales_master["WH"] = 'SV'

si_inv_master["WH"] = 'SI'
on_inv_master["WH"] = 'ON'
sv_inv_master["WH"] = 'SV'

# Rename units case cubic name group description
si_inv_master['unit'] = si_inv_master['si']
on_inv_master['unit'] = on_inv_master['on']
sv_inv_master['unit'] = sv_inv_master['sv']
del si_inv_master['si']
del on_inv_master['on']
del sv_inv_master['sv']

si_inv_master['caseqty'] = si_inv_master['si_caseqty']
on_inv_master['caseqty'] = on_inv_master['on_caseqty']
sv_inv_master['caseqty'] = sv_inv_master['sv_caseqty']
del si_inv_master['si_caseqty']
del on_inv_master['on_caseqty']
del sv_inv_master['sv_caseqty']

si_inv_master['cubic'] = si_inv_master['si_cubic']
on_inv_master['cubic'] = on_inv_master['on_cubic']
sv_inv_master['cubic'] = sv_inv_master['sv_cubic']
del si_inv_master['si_cubic']
del on_inv_master['on_cubic']
del sv_inv_master['sv_cubic']

si_sales_master['customer'] = si_sales_master['name']
on_sales_master['customer'] = on_sales_master['name']
sv_sales_master['customer'] = sv_sales_master['name']
del sv_sales_master['name']
del si_sales_master['name']
del on_sales_master['name']

si_sales_master['description'] = si_sales_master['desc']
on_sales_master['description'] = on_sales_master['desc']
sv_sales_master['description'] = sv_sales_master['desc']
del sv_sales_master['desc']
del si_sales_master['desc']
del on_sales_master['desc']

si_inv_master['description'] = si_inv_master['desc']
on_inv_master['description'] = on_inv_master['desc']
sv_inv_master['description'] = sv_inv_master['desc']
del sv_inv_master['desc']
del si_inv_master['desc']
del on_inv_master['desc']

si_inv_master['group'] = si_inv_master['grp']
on_inv_master['group'] = on_inv_master['grp']
sv_inv_master['group'] = sv_inv_master['grp']
del sv_inv_master['grp']
del si_inv_master['grp']
del on_inv_master['grp']

# Col Reorg
si_sales_master = si_sales_master[['WH', 'month', 'division', 'linecode', 'style', 'color', 'description', 'catcode',
                                   'upcno', 'shptot', 'price', 'amount', 'cost', 'account', 'customer', 'custpo',
                                   'invoice', 'invdate', 'order', 'entered', 'edifile', 'note1', 'ststate', 'stzip',
                                   'piktkt', 'piktktdate']]
on_sales_master = on_sales_master[['WH', 'month', 'division', 'linecode', 'style', 'color', 'description', 'catcode',
                                   'upcno', 'shptot', 'price', 'amount', 'cost', 'account', 'customer', 'custpo',
                                   'invoice', 'invdate', 'order', 'entered', 'edifile', 'note1', 'ststate', 'stzip',
                                   'piktkt', 'piktktdate']]
sv_sales_master = sv_sales_master[['WH', 'month', 'division', 'linecode', 'style', 'color', 'description', 'catcode',
                                   'upcno', 'shptot', 'price', 'amount', 'cost', 'account', 'customer', 'custpo',
                                   'invoice', 'invdate', 'order', 'entered', 'edifile', 'note1', 'ststate', 'stzip',
                                   'piktkt', 'piktktdate']]

si_inv_master = si_inv_master[['WH', 'style', 'group', 'description', 'color', 'division', 'cubic_ft', 'weight',
                               'master_pack', 'unit', 'caseqty', 'cubic']]
on_inv_master = on_inv_master[['WH', 'style', 'group', 'description', 'color', 'division', 'cubic_ft', 'weight',
                               'master_pack', 'unit', 'caseqty', 'cubic']]
sv_inv_master = sv_inv_master[['WH', 'style', 'group', 'description', 'color', 'division', 'cubic_ft', 'weight',
                               'master_pack', 'unit', 'caseqty', 'cubic']]

# data type manip
si_sales_master['WH'] = si_sales_master['WH'].astype(str)
si_sales_master['amount'] = si_sales_master['amount'].astype(float)

################################################### Creating Master sheets #############################################
frame = [si_sales_master, on_sales_master, sv_sales_master]
sales_master = pd.concat(frame, ignore_index=True)

frame = [si_inv_master, on_inv_master, sv_inv_master]
inv_master = pd.concat(frame, ignore_index=True)


####################################################### Changes ########################################################
# DF merge to obtain CBM and CasePack from INV for SalesMaster(VLOOKUP SIM)
sales_master["ID"] = sales_master["style"]+sales_master["color"]

inv_rep = inv_rep[['ID', 'CRTN CBM', 'Casepack']]

sales_master = pd.merge(sales_master, inv_rep, on='ID', how='inner')

inv_master = inv_master[['WH', 'style', 'group', 'description', 'color', 'division', 'cubic_ft', 'weight', 'master_pack'
                                , 'unit', 'caseqty', 'cubic']]

# Rounding
sales_master['Total CBM'] = round((sales_master['CRTN CBM']/sales_master['Casepack'])*sales_master['shptot'], 4)
sales_master['CRTN CBM'] = round(sales_master['CRTN CBM'], 4)
sales_master['Total CBM'] = sales_master['Total CBM'].astype(float)


############################################# Summary Functions ########################################################

# whname: single quote string; 'SI', 'ON', 'SV'
# colname: single quote string; 'amount', 'shptot', 'Total CBM'
# sheetname: df variable name; sales_month_master, unit_month_master, cbm_month_master
def maketable(sheetname, colname, whname):
    sales_month = sales_master[sales_master.WH == whname]
    sales_month = sales_month[['WH', 'month', colname]]
    month_01 = sales_month[sales_month.month == '01']
    month_02 = sales_month[sales_month.month == '02']
    month_03 = sales_month[sales_month.month == '03']
    month_04 = sales_month[sales_month.month == '04']
    month_05 = sales_month[sales_month.month == '05']
    month_06 = sales_month[sales_month.month == '06']
    month_07 = sales_month[sales_month.month == '07']
    month_08 = sales_month[sales_month.month == '08']
    month_09 = sales_month[sales_month.month == '09']
    month_10 = sales_month[sales_month.month == '10']
    month_11 = sales_month[sales_month.month == '11']
    month_12 = sales_month[sales_month.month == '12']
    mo_01 = sum(month_01[colname])
    mo_02 = sum(month_02[colname])
    mo_03 = sum(month_03[colname])
    mo_04 = sum(month_04[colname])
    mo_05 = sum(month_05[colname])
    mo_06 = sum(month_06[colname])
    mo_07 = sum(month_07[colname])
    mo_08 = sum(month_08[colname])
    mo_09 = sum(month_09[colname])
    mo_10 = sum(month_10[colname])
    mo_11 = sum(month_11[colname])
    mo_12 = sum(month_12[colname])

    mo_sum = mo_01 + mo_02 + mo_03 + mo_04 + mo_05 + mo_06 + mo_07 + mo_08 + mo_09 + mo_10 + mo_11 + mo_12
    molist = [mo_01, mo_02, mo_03, mo_04, mo_05, mo_06, mo_07, mo_08, mo_09, mo_10, mo_11, mo_12, mo_sum]
    sheetname[whname] = molist


# sheetname: df variable name; sales_month_master, unit_month_master, cbm_month_master
def maketot(sheetname):
    tot01 = sheetname.loc[0, 'ON'] + sheetname.loc[0, 'SI'] + sheetname.loc[0, 'SV']
    tot02 = sheetname.loc[1, 'ON'] + sheetname.loc[1, 'SI'] + sheetname.loc[1, 'SV']
    tot03 = sheetname.loc[2, 'ON'] + sheetname.loc[2, 'SI'] + sheetname.loc[2, 'SV']
    tot04 = sheetname.loc[3, 'ON'] + sheetname.loc[3, 'SI'] + sheetname.loc[3, 'SV']
    tot05 = sheetname.loc[4, 'ON'] + sheetname.loc[4, 'SI'] + sheetname.loc[4, 'SV']
    tot06 = sheetname.loc[5, 'ON'] + sheetname.loc[5, 'SI'] + sheetname.loc[5, 'SV']
    tot07 = sheetname.loc[6, 'ON'] + sheetname.loc[6, 'SI'] + sheetname.loc[6, 'SV']
    tot08 = sheetname.loc[7, 'ON'] + sheetname.loc[7, 'SI'] + sheetname.loc[7, 'SV']
    tot09 = sheetname.loc[8, 'ON'] + sheetname.loc[8, 'SI'] + sheetname.loc[8, 'SV']
    tot10 = sheetname.loc[9, 'ON'] + sheetname.loc[9, 'SI'] + sheetname.loc[9, 'SV']
    tot11 = sheetname.loc[10, 'ON'] + sheetname.loc[10, 'SI'] + sheetname.loc[10, 'SV']
    tot12 = sheetname.loc[11, 'ON'] + sheetname.loc[11, 'SI'] + sheetname.loc[11, 'SV']
    totsum = sheetname.loc[12, 'ON'] + sheetname.loc[12, 'SI'] + sheetname.loc[12, 'SV']
    totlist = [tot01, tot02, tot03, tot04, tot05, tot06, tot07, tot08, tot09, tot10, tot11, tot12, totsum]
    sheetname['Grand Total'] = totlist


########################################### Sales by wh by month #######################################################

sales_month = pd.DataFrame()
months = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12', 'Total']
sales_month_master = pd.DataFrame(months, columns=['month'])

maketable(sales_month_master,'amount','ON')
maketable(sales_month_master,'amount','SI')
maketable(sales_month_master,'amount','SV')

maketot(sales_month_master)


#################################### Units by WH by month ##############################################################
unit_month = pd.DataFrame()
unit_month_master = pd.DataFrame(months, columns=['Month'])

maketable(unit_month_master,'shptot','ON')
maketable(unit_month_master,'shptot','SI')
maketable(unit_month_master,'shptot','SV')

maketot(unit_month_master)


############################################### Creating CBM by WH by mth ##############################################
cbm_month = pd.DataFrame()
cbm_month_master = pd.DataFrame(months, columns=['Month'])

maketable(cbm_month_master,'Total CBM','ON')
maketable(cbm_month_master,'Total CBM','SI')
maketable(cbm_month_master,'Total CBM','SV')

maketot(cbm_month_master)


############################################### Creating CBF by WH by mth ##############################################

cbf_month_master = pd.DataFrame()
cbf_month_master['Month'] = cbm_month_master['Month']
cbf_month_master['ON'] = cbm_month_master['ON'] * 35.3146667
cbf_month_master['SI'] = cbm_month_master['SI'] * 35.3146667
cbf_month_master['SV'] = cbm_month_master['SV'] * 35.3146667
cbf_month_master['Grand Total'] = cbm_month_master['Grand Total'] * 35.3146667

############################################### Creating Unit by WH Total ##############################################

unit_tot_on = sum(on_inv_master['unit'])
unit_tot_si = sum(si_inv_master['unit'])
unit_tot_sv = sum(sv_inv_master['unit'])
unit_tot_tot = unit_tot_on+unit_tot_si+unit_tot_sv
unit_tot = pd.DataFrame({'': ["Sum of Unit"], 'ON': [unit_tot_on], 'SI': [unit_tot_si], 'SV': [unit_tot_sv],
                         'GrandTotal': [unit_tot_tot]})

############################################## Creating Cubic by WH Total ##############################################

cube_tot_on = sum(on_inv_master['cubic'])
cube_tot_si = sum(si_inv_master['cubic'])
cube_tot_sv = sum(sv_inv_master['cubic'])
cube_tot_tot = cube_tot_on+cube_tot_si+cube_tot_sv
cube_tot = pd.DataFrame({'': ["Sum of Cubic"], 'ON': [cube_tot_on], 'SI': [cube_tot_si], 'SV': [cube_tot_sv],
                         'GrandTotal': [cube_tot_tot]})
##################################################### FILE OUTPUTS #####################################################


fileConstructor = pd.ExcelWriter(save_Loc, engine='xlsxwriter')

# Sales OUT
si_sales_master.to_excel(fileConstructor, sheet_name='SI Sales', index=False)
on_sales_master.to_excel(fileConstructor, sheet_name='ON Sales', index=False)
sv_sales_master.to_excel(fileConstructor, sheet_name='SV Sales', index=False)

#INV OUT
si_inv_master.to_excel(fileConstructor, sheet_name='SI INV', index=False)
on_inv_master.to_excel(fileConstructor, sheet_name='ON INV', index=False)
sv_inv_master.to_excel(fileConstructor, sheet_name='SV INV', index=False)

# SUMMARY OUT
sales_master.to_excel(fileConstructor, sheet_name='Total Sales', index=False)
inv_master.to_excel(fileConstructor, sheet_name='Total INV', index=False)

sales_month_master.to_excel(fileConstructor, sheet_name='Sales by WH by Month', index=False)
unit_month_master.to_excel(fileConstructor, sheet_name='Units by WH by Month', index=False)
cbm_month_master.to_excel(fileConstructor, sheet_name='CBM by WH by Month', index=False)
cbf_month_master.to_excel(fileConstructor, sheet_name='CBF by WH by Month', index=False)

unit_tot.to_excel(fileConstructor, sheet_name='unit totals', index=False)
cube_tot.to_excel(fileConstructor, sheet_name='cubic totals', index=False)

fileConstructor.save()
