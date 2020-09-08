import pandas as pd

file = pd.read_excel('1_20200604.xls')


df = file[['Дата Партии', 'Group ID', 'ИнгредиентыID', 'Тип ингредиента', 'Загруженный вес']]

date = df['Дата Партии'].reset_index()
date = date.iloc[0, 1]
date = date.date()
ferma_mak = ('000000006', '000000007')
ferma_kar = ('000000014', '000000134', '000000019')
# Делим по фермам
mak = df.loc[df['Group ID'].isin(ferma_mak)]
kar = df.loc[df['Group ID'].isin(ferma_kar)]
#===========================

# Делим по ингридиентам
id_ingr = ('00000026211', '00000026213')
#id_sxpr = ('00000026211',)
id_prixod_vp = ('00000026213',)

prixod_mak = mak.loc[mak['ИнгредиентыID'].isin(id_ingr)]
prixod_mak = prixod_mak.loc[~prixod_mak['Тип ингредиента'].isin(['Субпродукт'])]
prixod_mak = prixod_mak.groupby(['ИнгредиентыID'], sort=True).sum()[['Загруженный вес']].reset_index()
prixod_mak['ИнгредиентыID'].replace('00000026211', 'Зїди СхРП', inplace=True)
prixod_mak['ИнгредиентыID'].replace('00000026213', 'Зїди ВП', inplace=True)
prixod_mak = prixod_mak.rename(columns={prixod_mak.columns[0]: "Компонент", prixod_mak.columns[1]: "Вага"})

prixod_kar = kar.loc[kar['ИнгредиентыID'].isin(id_ingr)]
prixod_kar = prixod_kar.groupby(['ИнгредиентыID'], sort=True).sum()[['Загруженный вес']].reset_index()
prixod_kar['ИнгредиентыID'].replace('00000026213', 'Зїди', inplace=True)
prixod_kar = prixod_kar.rename(columns={prixod_kar.columns[0]: "Компонент", prixod_kar.columns[1]: "Вага"})

writer = pd.ExcelWriter(r"d:\\{0}".format('Звіт_ПВ_' + '123' + '.xlsx', engine='xlsxwriter'))

farm_mak = 'МАКОВИЩЕ'
farm_kar = 'КАРАШИН'
header_prixod = 'ЗВІТ ПО НАДХОДЖЕННЮ ПОВТОРНИХ ВІДХОДІВ ПО ФЕРМАХ  ' '123'
podpis = 'ПІБ     Нагорна Н.     ПІДПИС ____________________ ДАТА   ' + '123'
podpisBilous = 'ПІБ  Білоус Я.М.  ПІДПИС ____________________ ДАТА   ' + '123'
podpisZina = 'ПІБ  Грицак З.  ПІДПИС ____________________ ДАТА   ' + '123'
prixod_mak.to_excel(writer, sheet_name='Прихід', startrow=6, startcol=1, index=False)
prixod_kar.to_excel(writer, sheet_name='Прихід', startrow=21, startcol=1, index=False)

worksheet = writer.sheets['Прихід']
worksheet.write(1, 1, header_prixod)
worksheet.write(3, 1, farm_mak)
worksheet.write(12, 1, podpis)
worksheet.write(15, 1, podpisBilous)
worksheet.write(19, 1, farm_kar)
worksheet.write(26, 1, podpis)
worksheet.write(28, 1, podpisZina)
worksheet.set_column('B:B', 24)
worksheet.set_column('C:C', 25)
writer.save()
writer.close()



#prixod_vp_mak = mak.loc[mak['ИнгредиентыID'].isin(id_sxpr)]
#prixod_vp_mak = prixod_vp_mak.loc[~prixod_vp_mak['Тип ингредиента'].isin(['Субпродукт'])]


#print(prixod_vp_mak)

#prixod_sxrp_kar = kar.loc[kar['ИнгредиентыID'].isin(id_ingr)]
#print(prixod_sxrp_kar)
#prixod_vp_kar = kar.loc[kar['ИнгредиентыID'].isin(id_ingr)]
#print(prixod_vp_kar)
#prixod_sxrp = prixod_sxrp_mak.groupby(['Group ID'], sort=True).sum()[['Загруженный вес']].reset_index()


#prixod_sxrp.rename({6.0: "Зъди СХ"})
#prixod_sxrp = prixod_sxrp.rename(row={prixod_sxrp.row[0]:"Зъди СХ"})


#prixod_vp = prixod_vp_mak.groupby(['ИнгредиентыID'], sort=True).sum()[['Загруженный вес']].reset_index()
#full_mak = prixod_vp.merge(prixod_sxrp, how='outer')
#full_mak = full_mak.rename(columns={full_mak.columns[0]: "Компонент", full_mak.columns[1]: "Вага"})
#full_mak.set_value(0, 0, "Зъди СХ")
#full_mak.at[0, "Компонент", ] = "Зъди СХ"
#prixod_sxrp_kar = prixod_sxrp_kar.groupby(['Group ID'], sort=True).sum()[['Загруженный вес']].reset_index()
#prixod_vp_kar = prixod_vp_kar.groupby(['Group ID'], sort=True).sum()[['Загруженный вес']].reset_index()

#writer = pd.ExcelWriter(r"d:\\{0}".format('Звіт_ПВ_Маковище_' + '123' + '.xlsx', engine='xlsxwriter'))
#prixod = pd.DataFrame(prixod)
# prixod.to_excel(writer, sheet_name='Прихід', index=False)
#header_prixod = 'ЗВІТ ПО НАДХОДЖЕННЮ ПОВТОРНИХ ВІДХОДІВ ПО  ФЕРМІ МАКОВИЩЕ ЗА  ' '123'
#podpis = 'ПІБ     Нагорна Н.     ПІДПИС ____________________ ДАТА   ' + '123'
#podpisBilous = 'ПІБ  Білоус Я.М.  ПІДПИС ____________________ ДАТА   ' + '123'
#full_mak.to_excel(writer, sheet_name='Прихід', startrow=3, startcol=1, index=False)
#prixod_vp.to_excel(writer, sheet_name='Прихід', startrow=6, startcol=1, index=False)

#worksheet = writer.sheets['Прихід']
#worksheet.write(1, 1, header_prixod)
#worksheet.write(30, 1, podpis)
#worksheet.write(33, 1, podpisBilous)
#worksheet.set_column('B:B', 24)
#worksheet.set_column('C:C', 25)
#writer.save()
#writer.close()