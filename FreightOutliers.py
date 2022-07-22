# Authors: John Ayres
# Desc: identifies freight cost outliers
# PARAM: Inventory Report.xlsx
# OUT: xlsx file cleanly listing PO's and relative information in their own cell
import pandas as pd
import numpy as np
import grubbs
################################################### PATHS AND READ #####################################################

invp = r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\05 - INVENTORY\Inventory Report_July 21, 2022.xlsx'

inv = pd.read_excel(invp, engine='openpyxl',header=2)

############################################# Categorization and data type #############################################

inv["Category"] = inv["Category"].astype('category')
inv["Ocean Freight Cost"] = inv["Ocean Freight Cost"].astype(float)
grouped = inv.groupby(inv["Category"])
basi = grouped.get_group("Basin")
bacc = grouped.get_group("Bathroom Acc")
bfau = grouped.get_group("Bathroom Faucet")
chan = grouped.get_group("Chandelier")
disp = grouped.get_group("Display")
flam = grouped.get_group("Floor Lamp")
flus = grouped.get_group("Flushmount")
isli = grouped.get_group("Island Light")
kfac = grouped.get_group("Kitchen Faucet")
ksin = grouped.get_group("Kitchen Sink")
kwar = grouped.get_group("Kitchenware")
mirr = grouped.get_group("Mirror")
olig = grouped.get_group("Outdoor Light")
pend = grouped.get_group("Pendant")
repp = grouped.get_group("Replacement Parts")
scur = grouped.get_group("Shower Curtain")
sdor = grouped.get_group("Shower Door")
sfau = grouped.get_group("Shower Faucet")
shea = grouped.get_group("Showerhead")
tlam = grouped.get_group("Table Lamp")
toil = grouped.get_group("Toilet")
tlig = grouped.get_group("Track Lighting")
ucab = grouped.get_group("Under Cabinet")
vani = grouped.get_group("Vanity")
vlig = grouped.get_group("Vanity Light")
whar = grouped.get_group("Wall Hardware")
wsco = grouped.get_group("Wall Sconce")


############################################# Log Transformation and tests #############################################

basi['Ocean Freight Cost'] = np.log(basi['Ocean Freight Cost'])
bacc['Ocean Freight Cost'] = np.log(bacc['Ocean Freight Cost'])
bfau['Ocean Freight Cost'] = np.log(bfau['Ocean Freight Cost'])
chan['Ocean Freight Cost'] = np.log(chan['Ocean Freight Cost'])
disp['Ocean Freight Cost'] = np.log(disp['Ocean Freight Cost'])
flam['Ocean Freight Cost'] = np.log(flam['Ocean Freight Cost'])
flus['Ocean Freight Cost'] = np.log(flus['Ocean Freight Cost'])
isli['Ocean Freight Cost'] = np.log(isli['Ocean Freight Cost'])
kfac['Ocean Freight Cost'] = np.log(kfac['Ocean Freight Cost'])
ksin['Ocean Freight Cost'] = np.log(ksin['Ocean Freight Cost'])
kwar['Ocean Freight Cost'] = np.log(kwar['Ocean Freight Cost'])
mirr['Ocean Freight Cost'] = np.log(mirr['Ocean Freight Cost'])
olig['Ocean Freight Cost'] = np.log(olig['Ocean Freight Cost'])
pend['Ocean Freight Cost'] = np.log(pend['Ocean Freight Cost'])
repp['Ocean Freight Cost'] = np.log(repp['Ocean Freight Cost'])
scur['Ocean Freight Cost'] = np.log(scur['Ocean Freight Cost'])
sdor['Ocean Freight Cost'] = np.log(sdor['Ocean Freight Cost'])
sfau['Ocean Freight Cost'] = np.log(sfau['Ocean Freight Cost'])
shea['Ocean Freight Cost'] = np.log(shea['Ocean Freight Cost'])
tlam['Ocean Freight Cost'] = np.log(tlam['Ocean Freight Cost'])
toil['Ocean Freight Cost'] = np.log(toil['Ocean Freight Cost'])
tlig['Ocean Freight Cost'] = np.log(tlig['Ocean Freight Cost'])
ucab['Ocean Freight Cost'] = np.log(ucab['Ocean Freight Cost'])
vani['Ocean Freight Cost'] = np.log(vani['Ocean Freight Cost'])
vlig['Ocean Freight Cost'] = np.log(vlig['Ocean Freight Cost'])
whar['Ocean Freight Cost'] = np.log(whar['Ocean Freight Cost'])
wsco['Ocean Freight Cost'] = np.log(wsco['Ocean Freight Cost'])

dframes = [basi, bacc, bfau, chan, disp, flam, flus, isli, kfac, ksin, kwar, mirr, olig, pend, repp, scur, sdor, sfau,
           shea, tlam, toil, tlig, ucab, vani, vlig, whar, wsco]

basi.reset_index(drop=True,inplace=True)
bacc.reset_index(drop=True,inplace=True)
bfau.reset_index(drop=True,inplace=True)
chan.reset_index(drop=True,inplace=True)
disp.reset_index(drop=True,inplace=True)
flam.reset_index(drop=True,inplace=True)
flus.reset_index(drop=True,inplace=True)
isli.reset_index(drop=True,inplace=True)
kfac.reset_index(drop=True,inplace=True)
ksin.reset_index(drop=True,inplace=True)
kwar.reset_index(drop=True,inplace=True)
mirr.reset_index(drop=True,inplace=True)
olig.reset_index(drop=True,inplace=True)
pend.reset_index(drop=True,inplace=True)
repp.reset_index(drop=True,inplace=True)
scur.reset_index(drop=True,inplace=True)
sdor.reset_index(drop=True,inplace=True)
sfau.reset_index(drop=True,inplace=True)
shea.reset_index(drop=True,inplace=True)
tlam.reset_index(drop=True,inplace=True)
toil.reset_index(drop=True,inplace=True)
tlig.reset_index(drop=True,inplace=True)
ucab.reset_index(drop=True,inplace=True)
vani.reset_index(drop=True,inplace=True)
vlig.reset_index(drop=True,inplace=True)
whar.reset_index(drop=True,inplace=True)
wsco.reset_index(drop=True,inplace=True)


test01 = grubbs.detect_outliers(basi['Ocean Freight Cost'],alternative="max")
#test02 = grubbs.detect_outliers(bacc['Ocean Freight Cost'],alternative="max")
test03 = grubbs.detect_outliers(bfau['Ocean Freight Cost'],alternative="max")
test04 = grubbs.detect_outliers(chan['Ocean Freight Cost'],alternative="max")
test05 = grubbs.detect_outliers(disp['Ocean Freight Cost'],alternative="max")
test06 = grubbs.detect_outliers(flam['Ocean Freight Cost'],alternative="max")
test07 = grubbs.detect_outliers(flus['Ocean Freight Cost'],alternative="max")
test08 = grubbs.detect_outliers(isli['Ocean Freight Cost'],alternative="max")
test09 = grubbs.detect_outliers(kfac['Ocean Freight Cost'],alternative="max")
test10 = grubbs.detect_outliers(ksin['Ocean Freight Cost'],alternative="max")
test11 = grubbs.detect_outliers(kwar['Ocean Freight Cost'],alternative="max")
test12 = grubbs.detect_outliers(mirr['Ocean Freight Cost'],alternative="max")
test13 = grubbs.detect_outliers(olig['Ocean Freight Cost'],alternative="max")
test14 = grubbs.detect_outliers(pend['Ocean Freight Cost'],alternative="max")
test15 = grubbs.detect_outliers(repp['Ocean Freight Cost'],alternative="max")
test16 = grubbs.detect_outliers(scur['Ocean Freight Cost'],alternative="max")
test17 = grubbs.detect_outliers(sdor['Ocean Freight Cost'],alternative="max")
test18 = grubbs.detect_outliers(sfau['Ocean Freight Cost'],alternative="max")
test19 = grubbs.detect_outliers(shea['Ocean Freight Cost'],alternative="max")
test20 = grubbs.detect_outliers(tlam['Ocean Freight Cost'],alternative="max")
test21 = grubbs.detect_outliers(toil['Ocean Freight Cost'],alternative="max")
test22 = grubbs.detect_outliers(tlig['Ocean Freight Cost'],alternative="max")
test23 = grubbs.detect_outliers(ucab['Ocean Freight Cost'],alternative="max")
test24 = grubbs.detect_outliers(vani['Ocean Freight Cost'],alternative="max")
test25 = grubbs.detect_outliers(vlig['Ocean Freight Cost'],alternative="max")
test26 = grubbs.detect_outliers(whar['Ocean Freight Cost'],alternative="max")
test27 = grubbs.detect_outliers(wsco['Ocean Freight Cost'],alternative="max")

print(test14)

#So i dont have to type a bunch
# basi
# bacc
# bfau
# chan
# disp
# flam
# flus
# isli
# kfac
# ksin
# kwar
# mirr
# olig
# pend
# repp
# scur
# sdor
# sfau
# shea
# tlam
# toil
# tlig
# ucab
# vani
# vlig
# whar
# wsco
