# Matching-certificates
CERT = ['P_cULus','C_cULus','B_cULus','P_CE','C_CE','G_CE','P_RCM','C_RCM','B_RCM',
        'P_UR','C_UR','B_UR','P_WEEE','C_WEEE','B_WEEE',
        'P_CCC','C_CCC','B_CCC','P_ROHS','C_ROHS','B_ROHS',
        'P_RoHS','C_RoHS','B_RoHS','P_MAROCCO','C_MAROCCO','B_MAROCCO',
        'P_EAC','C_EAC','B_EAC','P_CSA','C_CSA','B_CSA',
        'P_KCC','C_KCC','B_KCC']

field_name = ['cULus','CE','RCM','UR','WEEE','CCC',
             'ROHS','RoHS','MAROCCO','EAC','CSA','KCC']

field_value = ['Culus listed cropped.PCX','Culus listed cropped.PCX','Culus listed cropped.PCX',
              'CE.PCX','CE.PCX','CE.PCX','RCM-TICK.JPG','RCM-TICK.JPG','RCM-TICK.JPG','UR.PCX','UR.PCX','UR.PCX',
              'WEEE.BMP','WEEE.BMP','WEEE.BMP','CCC_only.JPG','CCC_only.JPG','CCC_only.JPG',
              'ROHS25.BMP','ROHS25.BMP','ROHS25.BMP','ROHS_E.BMP','ROHS_E.BMP','ROHS_E.BMP',
              'ROROC.JPG','ROROC.JPG','ROROC.JPG','EAC.BMP','EAC.BMP','EAC.BMP','CSA.PCX',
              'CSA.PCX','CSA.PCX','KCC.TIF','KCC.TIF','KCC.TIF']

def whatelse():
    print('Do you want to add more values ? y/n')

end_date = '12/31/9999'

print('---------------------------------------------------------')
## wprowadzanie numerow katalogowych
PN = []
while True:
    Material = input('Insert the catalog number (only one or a list):')
    if Material == '':
        q = input('Do you want to add any more catalog numbers? y/n')
        if q.lower() == 'y':
            continue
        else:
            break
    PN.append(Material)
print('Download: {0}'.format(PN))

print('---------------------------------------------------------')
## wprowadzanie certyfikatow na odpowiednia etykiete
certs = [("culus", 0), ("ce", 1), ("rcm", 2), ("ur", 3), ("weee", 4), ("ccc", 5), ("rohs25", 6), ("rohse", 7),
         ("roroc", 8), ("eac", 9), ("csa", 10), ("kcc", 11)]

def add_cert(cert_param):
    for el in certs:
        if el[0] == cert_param:
            whatelse()
            cert = field_name[el[1]]
    return cert


values = []
while True:

    print('Is this product label, carton label, bag label? p/c/b')
    label = input()
    if label.lower() == 'p':                                                                # NAMEPLATE (product label)
        print('What do you want to add?')
        label = label.upper() + str('_')
        cert = input()
        cert = add_cert(cert)
        more = input()
        if more.lower() == 'y':
            values.append(label + cert)
        else:
            break
    if label.lower() == 'c':                                                                # CARTON LABEL
        print('What do you want to add?')
        label = label.upper() + str('_')
        cert = input()
        cert = add_cert(cert)
        more = input()
        if more.lower() == 'y':
            values.append(label + cert)
        else:
            break
    if label.lower() == 'b':                                                                # BAG LABEL
        print('What do you want to add?')
        label = label.upper() + str('_')
        cert = input()
        cert = add_cert(cert)
        more = input()
        if more.lower() == 'y':
            values.append(label + cert)
        else:
            break
    else:
        continue

values.append(label.replace('o','') + cert)
print('Values: {0}'.format(values))

print('---------------------------------------------------------')
## wprowadzanie daty

from datetime import date
date.today()

start_date = []
while True:
    your_date = input('Start date: today? y/n ')
    if your_date.lower() == 'y':
        start_date = date.today().strftime('%m/%d/%Y')
        break
    elif your_date.lower() == 'n':
        month = int(input('Month ( 1 / 12 ) : '))
        day = int(input('Day ( 1 / 31 ) : '))
        year = int(input('Year (2019 <= ): '))
        your_date = "{0}/{1}/{2}".format(month, day, year)
        if month <= 12 and day <= 31 and year >= 2019:
            start_date = your_date
            break
        else:
            continue
    else:
        continue

print('---------------------------------------------------------')
## data koncowa - na stale
print('End date:' + str(end_date))

print('---------------------------------------------------------')
## wprowadzanie nazwy wartosci pola
print('You have ' + str(len(PN)) + ' catalog number/s and ' + str(len(values)) + ' certification/s !')
print(PN)
print(values)
print(start_date)
print(end_date)


x = len(values)
y = []
for i in range(x):
    a = CERT.index(values[i])
    b = field_value[a]
    y.append(b)
print('Download: {0}'.format(y))

print('---------------------------------------------------------')
## wprowadzanie nazwy pliku

name = input('Insert file name: ')

print('---------------------------------------------------------')
## zapisywanie wartosci do pliku Excel

import openpyxl
wb = openpyxl.load_workbook(name + ".xlsx")
sheet = wb.get_sheet_by_name("Sheet1")
q = int(len(PN))
w = int(len(values))
x = 10

print('---------------------------------------------------------')
for i in range(0, q):
        sheet['A' + str(x)] = str(PN[i])
        x = x + w

x = 10
for i in range(0, q):
    x = 10
    for s in range(0, w):
        sheet['B' + str(x + (i * w))] = str(values[s])
        x += 1

x = 10
for i in range(0, (w * q)):
        sheet['C' + str(x)] = str(start_date)
        x += 1

x = 10
for i in range(0, (w * q)):
        sheet['D' + str(x)] = str(end_date)
        x += 1

x = 10
for i in range(0, q):
    x = 10
    for s in range(0, w):
        sheet['E' + str(x + (i * w))] = str(y[s])
        x += 1

wb.save(name + '.xlsx')


