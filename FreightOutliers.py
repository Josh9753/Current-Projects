# Authors: John Ayres
# Desc: identifies freight cost outliers through a normalized z-score. data is log normal.
# PARAM: Inventory Report.xlsx
# OUT: xlsx file cleanly listing PO's and relative information in their own cell
import pandas as pd
import numpy as np


################################################### PATHS AND READ #####################################################

invp = r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\05 - INVENTORY\Inventory Report_July 25, 2022.xlsx'

inv = pd.read_excel(invp, engine='openpyxl', header=2)

threshold = 1.86 #Z-score threshold (2 is less conservative, 1.86 is traditional value used)

save_Loc = r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\Freight Outlier Report TODAY.xlsx'
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

basi.reset_index(drop=True, inplace=True)
bacc.reset_index(drop=True, inplace=True)
bfau.reset_index(drop=True, inplace=True)
chan.reset_index(drop=True, inplace=True)
disp.reset_index(drop=True, inplace=True)
flam.reset_index(drop=True, inplace=True)
flus.reset_index(drop=True, inplace=True)
isli.reset_index(drop=True, inplace=True)
kfac.reset_index(drop=True, inplace=True)
ksin.reset_index(drop=True, inplace=True)
kwar.reset_index(drop=True, inplace=True)
mirr.reset_index(drop=True, inplace=True)
olig.reset_index(drop=True, inplace=True)
pend.reset_index(drop=True, inplace=True)
repp.reset_index(drop=True, inplace=True)
scur.reset_index(drop=True, inplace=True)
sdor.reset_index(drop=True, inplace=True)
sfau.reset_index(drop=True, inplace=True)
shea.reset_index(drop=True, inplace=True)
tlam.reset_index(drop=True, inplace=True)
toil.reset_index(drop=True, inplace=True)
tlig.reset_index(drop=True, inplace=True)
ucab.reset_index(drop=True, inplace=True)
vani.reset_index(drop=True, inplace=True)
vlig.reset_index(drop=True, inplace=True)
whar.reset_index(drop=True, inplace=True)
wsco.reset_index(drop=True, inplace=True)


################################################## Calculating Fails ###################################################

mean = np.mean(basi.loc[:, "Ocean Freight Cost"])
std = np.std(basi.loc[:, "Ocean Freight Cost"])
count = 0
for i in basi["Ocean Freight Cost"]:
    z = (i-mean)/std
    basi.loc[count, "Z-score"] = z
    if z > threshold:
        basi.loc[count, "Outlier"] = "True"
    else:
        basi.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(bacc.loc[:, "Ocean Freight Cost"])
std = np.std(bacc.loc[:, "Ocean Freight Cost"])
count = 0
for i in bacc["Ocean Freight Cost"]:
    z = (i-mean)/std
    bacc.loc[count, "Z-score"] = z
    if z > threshold:
        bacc.loc[count, "Outlier"] = "True"
    else:
        bacc.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(bfau.loc[:, "Ocean Freight Cost"])
std = np.std(bfau.loc[:, "Ocean Freight Cost"])
count = 0
for i in bfau["Ocean Freight Cost"]:
    z = (i-mean)/std
    bfau.loc[count, "Z-score"] = z
    if z > threshold:
        bfau.loc[count, "Outlier"] = "True"
    else:
        bfau.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(chan.loc[:, "Ocean Freight Cost"])
std = np.std(chan.loc[:, "Ocean Freight Cost"])
count = 0
for i in chan["Ocean Freight Cost"]:
    z = (i-mean)/std
    chan.loc[count, "Z-score"] = z
    if z > threshold:
        chan.loc[count, "Outlier"] = "True"
    else:
        chan.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(disp.loc[:, "Ocean Freight Cost"])
std = np.std(disp.loc[:, "Ocean Freight Cost"])
count = 0
for i in disp["Ocean Freight Cost"]:
    z = (i-mean)/std
    disp.loc[count, "Z-score"] = z
    if z > threshold:
        disp.loc[count, "Outlier"] = "True"
    else:
        disp.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(flam.loc[:, "Ocean Freight Cost"])
std = np.std(flam.loc[:, "Ocean Freight Cost"])
count = 0
for i in flam["Ocean Freight Cost"]:
    z = (i-mean)/std
    flam.loc[count, "Z-score"] = z
    if z > threshold:
        flam.loc[count, "Outlier"] = "True"
    else:
        flam.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(flus.loc[:, "Ocean Freight Cost"])
std = np.std(flus.loc[:, "Ocean Freight Cost"])
count = 0
for i in flus["Ocean Freight Cost"]:
    z = (i-mean)/std
    flus.loc[count, "Z-score"] = z
    if z > threshold:
        flus.loc[count, "Outlier"] = "True"
    else:
        flus.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(isli.loc[:, "Ocean Freight Cost"])
std = np.std(isli.loc[:, "Ocean Freight Cost"])
count = 0
for i in isli["Ocean Freight Cost"]:
    z = (i-mean)/std
    isli.loc[count, "Z-score"] = z
    if z > threshold:
        isli.loc[count, "Outlier"] = "True"
    else:
        isli.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(kfac.loc[:, "Ocean Freight Cost"])
std = np.std(kfac.loc[:, "Ocean Freight Cost"])
count = 0
for i in kfac["Ocean Freight Cost"]:
    z = (i-mean)/std
    kfac.loc[count, "Z-score"] = z
    if z > threshold:
        kfac.loc[count, "Outlier"] = "True"
    else:
        kfac.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(ksin.loc[:, "Ocean Freight Cost"])
std = np.std(ksin.loc[:, "Ocean Freight Cost"])
count = 0
for i in ksin["Ocean Freight Cost"]:
    z = (i-mean)/std
    ksin.loc[count, "Z-score"] = z
    if z > threshold:
        ksin.loc[count, "Outlier"] = "True"
    else:
        ksin.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(kwar.loc[:, "Ocean Freight Cost"])
std = np.std(kwar.loc[:, "Ocean Freight Cost"])
count = 0
for i in kwar["Ocean Freight Cost"]:
    z = (i-mean)/std
    kwar.loc[count, "Z-score"] = z
    if z > threshold:
        kwar.loc[count, "Outlier"] = "True"
    else:
        kwar.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(mirr.loc[:, "Ocean Freight Cost"])
std = np.std(mirr.loc[:, "Ocean Freight Cost"])
count = 0
for i in mirr["Ocean Freight Cost"]:
    z = (i-mean)/std
    mirr.loc[count, "Z-score"] = z
    if z > threshold:
        mirr.loc[count, "Outlier"] = "True"
    else:
        mirr.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(olig.loc[:, "Ocean Freight Cost"])
std = np.std(olig.loc[:, "Ocean Freight Cost"])
count = 0
for i in olig["Ocean Freight Cost"]:
    z = (i-mean)/std
    olig.loc[count, "Z-score"] = z
    if z > threshold:
        olig.loc[count, "Outlier"] = "True"
    else:
        olig.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(pend.loc[:, "Ocean Freight Cost"])
std = np.std(pend.loc[:, "Ocean Freight Cost"])
count = 0
for i in pend["Ocean Freight Cost"]:
    z = (i-mean)/std
    pend.loc[count, "Z-score"] = z
    if z > threshold:
        pend.loc[count, "Outlier"] = "True"
    else:
        pend.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(repp.loc[:, "Ocean Freight Cost"])
std = np.std(repp.loc[:, "Ocean Freight Cost"])
count = 0
for i in repp["Ocean Freight Cost"]:
    z = (i-mean)/std
    repp.loc[count, "Z-score"] = z
    if z > threshold:
        repp.loc[count, "Outlier"] = "True"
    else:
        repp.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(scur.loc[:, "Ocean Freight Cost"])
std = np.std(scur.loc[:, "Ocean Freight Cost"])
count = 0
for i in scur["Ocean Freight Cost"]:
    z = (i-mean)/std
    scur.loc[count, "Z-score"] = z
    if z > threshold:
        scur.loc[count, "Outlier"] = "True"
    else:
        scur.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(sdor.loc[:, "Ocean Freight Cost"])
std = np.std(sdor.loc[:, "Ocean Freight Cost"])
count = 0
for i in sdor["Ocean Freight Cost"]:
    z = (i-mean)/std
    sdor.loc[count, "Z-score"] = z
    if z > threshold:
        sdor.loc[count, "Outlier"] = "True"
    else:
        sdor.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(sfau.loc[:, "Ocean Freight Cost"])
std = np.std(sfau.loc[:, "Ocean Freight Cost"])
count = 0
for i in sfau["Ocean Freight Cost"]:
    z = (i-mean)/std
    sfau.loc[count, "Z-score"] = z
    if z > threshold:
        sfau.loc[count, "Outlier"] = "True"
    else:
        sfau.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(shea.loc[:, "Ocean Freight Cost"])
std = np.std(shea.loc[:, "Ocean Freight Cost"])
count = 0
for i in shea["Ocean Freight Cost"]:
    z = (i-mean)/std
    shea.loc[count, "Z-score"] = z
    if z > threshold:
        shea.loc[count, "Outlier"] = "True"
    else:
        shea.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(tlam.loc[:, "Ocean Freight Cost"])
std = np.std(tlam.loc[:, "Ocean Freight Cost"])
count = 0
for i in tlam["Ocean Freight Cost"]:
    z = (i-mean)/std
    tlam.loc[count, "Z-score"] = z
    if z > threshold:
        tlam.loc[count, "Outlier"] = "True"
    else:
        tlam.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(toil.loc[:, "Ocean Freight Cost"])
std = np.std(toil.loc[:, "Ocean Freight Cost"])
count = 0
for i in toil["Ocean Freight Cost"]:
    z = (i-mean)/std
    toil.loc[count, "Z-score"] = z
    if z > threshold:
        toil.loc[count, "Outlier"] = "True"
    else:
        toil.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(tlig.loc[:, "Ocean Freight Cost"])
std = np.std(tlig.loc[:, "Ocean Freight Cost"])
count = 0
for i in tlig["Ocean Freight Cost"]:
    z = (i-mean)/std
    tlig.loc[count, "Z-score"] = z
    if z > threshold:
        tlig.loc[count, "Outlier"] = "True"
    else:
        tlig.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(ucab.loc[:, "Ocean Freight Cost"])
std = np.std(ucab.loc[:, "Ocean Freight Cost"])
count = 0
for i in ucab["Ocean Freight Cost"]:
    z = (i-mean)/std
    ucab.loc[count, "Z-score"] = z
    if z > threshold:
        ucab.loc[count, "Outlier"] = "True"
    else:
        ucab.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(vani.loc[:, "Ocean Freight Cost"])
std = np.std(vani.loc[:, "Ocean Freight Cost"])
count = 0
for i in vani["Ocean Freight Cost"]:
    z = (i-mean)/std
    vani.loc[count, "Z-score"] = z
    if z > threshold:
        vani.loc[count, "Outlier"] = "True"
    else:
        vani.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(vlig.loc[:, "Ocean Freight Cost"])
std = np.std(vlig.loc[:, "Ocean Freight Cost"])
count = 0
for i in vlig["Ocean Freight Cost"]:
    z = (i-mean)/std
    vlig.loc[count, "Z-score"] = z
    if z > threshold:
        vlig.loc[count, "Outlier"] = "True"
    else:
        vlig.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(whar.loc[:, "Ocean Freight Cost"])
std = np.std(whar.loc[:, "Ocean Freight Cost"])
count = 0
for i in whar["Ocean Freight Cost"]:
    z = (i-mean)/std
    whar.loc[count, "Z-score"] = z
    if z > threshold:
        whar.loc[count, "Outlier"] = "True"
    else:
        whar.loc[count, "Outlier"] = "False"
    count = count + 1
mean = np.mean(wsco.loc[:, "Ocean Freight Cost"])
std = np.std(wsco.loc[:, "Ocean Freight Cost"])
count = 0
for i in wsco["Ocean Freight Cost"]:
    z = (i-mean)/std
    wsco.loc[count, "Z-score"] = z
    if z > threshold:
        wsco.loc[count, "Outlier"] = "True"
    else:
        wsco.loc[count, "Outlier"] = "False"
    count = count + 1


################################################## Creating Fails DF ###################################################
outlier = pd.DataFrame()
outlier = pd.concat([outlier, basi[basi["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, bacc[bacc["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, bfau[bfau["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, chan[chan["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, disp[disp["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, flam[flam["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, flus[flus["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, isli[isli["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, kfac[kfac["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, ksin[ksin["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, kwar[kwar["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, mirr[mirr["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, olig[olig["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, pend[pend["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, repp[repp["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, scur[scur["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, sdor[sdor["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, sfau[sfau["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, shea[shea["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, tlam[tlam["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, toil[toil["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, tlig[tlig["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, ucab[ucab["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, vani[vani["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, vlig[vlig["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, whar[whar["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, wsco[wsco["Outlier"] == "True"]], ignore_index=True)

##################################################### FILE OUTPUTS #####################################################

fileName = pd.ExcelWriter(save_Loc, engine='xlsxwriter')

outlier.to_excel(fileName, sheet_name='Outliers', index=False)
inv.to_excel(fileName, sheet_name='Inventory Report', index=False)

fileName.save()


# So i dont have to type a bunch
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
