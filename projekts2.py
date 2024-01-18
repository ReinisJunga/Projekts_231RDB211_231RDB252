from openpyxl import Workbook, load_workbook
import sys

temp = False
temp2 = True

while temp == False:
    filiale=input("Izvēlieties filiāli (Rīga, Mārupe, Ādaži): ")
    if filiale != "Rīga" and filiale != "Mārupe" and filiale != "Ādaži":
        print("Nepareizi norādīta filiāle")
        print("")
    else:
        temp = True

dienas = ["Pirmdiena", "Otrdiena", "Trešdiena", "Ceturtdiena", "Piektdiena"]
temp = False

while temp == False:
    diena=input("Izvēlieties dienu: ")
    for i in dienas:
        if diena == i:
            temp = True
    if temp == False:
        print("Nepareizi norādīta diena")

wb=load_workbook("Klasu grafiks 2023_2024.xlsx")
ws1=wb['Rīga']
ws2=wb['Mārupe']
ws3=wb['Ādaži']

brivie_laiki = []

brivais_laiks_pirmdiena=False
brivais_laiks_otrdiena=False
brivais_laiks_tresdiena=False
brivais_laiks_ceturtdiena=False
brivais_laiks_piektdiena=False

if filiale == 'Rīga':
    for row in range(3,13):
        brivie_laiki_pirmdiena = ws1['d' + str(row)].value
        brivs1 = ws1['b' + str(row)].value

        brivie_laiki_otrdiena = ws1['i' + str(row)].value
        brivs2 = ws1['g' + str(row)].value

        brivie_laiki_tresdiena = ws1['n' + str(row)].value
        brivs3 = ws1['l' + str(row)].value

        brivie_laiki_ceturtdiena = ws1['s' + str(row)].value
        brivs4 = ws1['q' + str(row)].value

        brivie_laiki_piektdiena = ws1['x' + str(row)].value
        brivs5 = ws1['v' + str(row)].value

        if diena == 'Pirmdiena' and brivs1 == None:
            brivie_laiki.append(str(brivie_laiki_pirmdiena))
            brivais_laiks_pirmdiena = True

        elif diena == 'Otrdiena' and brivs2 == None:
            brivie_laiki.append(str(brivie_laiki_otrdiena))
            brivais_laiks_otrdiena = True

        elif diena == 'Trešdiena' and brivs3 == None:
            brivie_laiki.append(str(brivie_laiki_tresdiena))
            brivais_laiks_tresdiena = True

        elif diena == 'Ceturtdiena' and brivs4 == None:
            brivie_laiki.append(str(brivie_laiki_ceturtdiena))
            brivais_laiks_ceturtdiena = True

        elif diena == 'Piektdiena' and brivs5 == None:
            brivie_laiki.append(str(brivie_laiki_piektdiena))
            brivais_laiks_piektdiena = True

    if not brivais_laiks_pirmdiena and not brivais_laiks_otrdiena and not brivais_laiks_tresdiena and not brivais_laiks_ceturtdiena and not brivais_laiks_piektdiena:
        print("Nav brīvu laiku")
        temp2 = False

if filiale == 'Mārupe':
    for row in range(3,13):
        brivie_laiki_pirmdiena = ws2['d' + str(row)].value
        brivs1 = ws2['b' + str(row)].value
        brivie_laiki_otrdiena = ws2['i' + str(row)].value
        brivs2 = ws2['g' + str(row)].value
        brivie_laiki_tresdiena = ws2['n' + str(row)].value
        brivs3 = ws2['l' + str(row)].value
        brivie_laiki_ceturtdiena = ws2['s' + str(row)].value
        brivs4 = ws2['q' + str(row)].value
        brivie_laiki_piektdiena = ws2['x' + str(row)].value
        brivs5 = ws2['v' + str(row)].value

        if diena == 'Pirmdiena' and brivs1 == None:
            brivie_laiki.append(str(brivie_laiki_pirmdiena))
            brivais_laiks_pirmdiena = True

        elif diena == 'Otrdiena' and brivs2 == None:
            brivie_laiki.append(str(brivie_laiki_otrdiena))
            brivais_laiks_otrdiena = True

        elif diena == 'Trešdiena' and brivs3 == None:
            brivie_laiki.append(str(brivie_laiki_tresdiena))
            brivais_laiks_tresdiena = True

        elif diena == 'Ceturtdiena' and brivs4 == None:
            brivie_laiki.append(str(brivie_laiki_ceturtdiena))
            brivais_laiks_ceturtdiena = True

        elif diena == 'Piektdiena' and brivs5 == None:
            brivie_laiki.append(str(brivie_laiki_piektdiena))
            brivais_laiks_piektdiena = True

    if not brivais_laiks_pirmdiena and not brivais_laiks_otrdiena and not brivais_laiks_tresdiena and not brivais_laiks_ceturtdiena and not brivais_laiks_piektdiena:
        print("Nav brīvu laiku")
        temp2 = False

if filiale == 'Ādaži':
    for row in range(3,13):
        brivie_laiki_pirmdiena = ws3['d' + str(row)].value
        brivs1 = ws3['b' + str(row)].value
        brivie_laiki_otrdiena = ws3['i' + str(row)].value
        brivs2 = ws3['g' + str(row)].value
        brivie_laiki_tresdiena = ws3['n' + str(row)].value
        brivs3 = ws3['l' + str(row)].value
        brivie_laiki_ceturtdiena = ws3['s' + str(row)].value
        brivs4 = ws3['q' + str(row)].value
        brivie_laiki_piektdiena = ws3['x' + str(row)].value
        brivs5 = ws3['v' + str(row)].value

        if diena == 'Pirmdiena' and brivs1 == None:
            brivie_laiki.append(str(brivie_laiki_pirmdiena))
            brivais_laiks_pirmdiena = True

        elif diena == 'Otrdiena' and brivs2 == None:
            brivie_laiki.append(str(brivie_laiki_otrdiena))
            brivais_laiks_otrdiena = True

        elif diena == 'Trešdiena' and brivs3 == None:
            brivie_laiki.append(str(brivie_laiki_tresdiena))
            brivais_laiks_tresdiena = True

        elif diena == 'Ceturtdiena' and brivs4 == None:
            brivie_laiki.append(str(brivie_laiki_ceturtdiena))
            brivais_laiks_ceturtdiena = True

        elif diena == 'Piektdiena' and brivs5 == None:
            brivie_laiki.append(str(brivie_laiki_piektdiena))
            brivais_laiks_piektdiena = True

    if not brivais_laiks_pirmdiena and not brivais_laiks_otrdiena and not brivais_laiks_tresdiena and not brivais_laiks_ceturtdiena and not brivais_laiks_piektdiena:
        print("Nav brīvu laiku")
        temp2 = False

if temp2 != False:
    print("Pieejami brīvi laiki:")
    print("")
    for i in brivie_laiki:
        print(" " + i)
else:
    wb.save('Klasu grafiks 2023_2024.xlsx')
    wb.close()
    sys.exit(1)

print("")

temp = False
while temp == False:
    print("Vai vēlaties pieteikt jaunu nodarbību? (Jā/Nē)")
    temp1 = input()
    if temp1 != "Jā" and temp1 != "Nē":
        print("Ievades kļūda, mēģiniet vēlreiz")
        print("")
    else:
        temp = True

if temp1 == "Nē":
    wb.save('Klasu grafiks 2023_2024.xlsx')
    wb.close()
    sys.exit(1)
temp = False
while temp == False:
    print("Izvēlaties laiku, kurā ievietot jaunu nodarbību: ")
    laiks = input()
    for i in brivie_laiki:
        if laiks == i:
            temp = True
    if temp == False:
        print("Ievadīts nederīgs laiks")
        print("")

print("Ievadiet pasniedzēja vārdu un uzvārdu: ") 
vards = input()

temp = False
while temp == False:
    print("Kāda veida nodarbība būs? (individuālā/grupu)")
    veids = input()
    if veids != "individuālā" and veids != "grupu":
        print("Ievadīts nederīgs veids")
        print("")
    else:
        temp = True

if veids == "individuālā":
    print("Ievadiet audzēkņa vārdu: ")
    audzekna_vards = input()

if veids == "grupu":
    audzekna_vards = []
    temp = False
    while temp == False:
        print ("Ievadiet audzēkņu skaitu (nevar būt vairāk par 6): ")
        try:
            skaits = int(input())
        except ValueError:
            print("Nederīga vērtība")
            continue
        if isinstance(skaits, int) and skaits <= 6:
            temp = True
        else:
            print("Nederīgs skaits")
    print("Ievadiet audzēkņu vārdus: ")
    while skaits > 0:
        audzekna_vards.append(input())
        skaits = skaits - 1
    audzekna_vards = ", ".join(audzekna_vards)



if filiale == 'Rīga':
    for row in range(3,13):
        if diena == 'Pirmdiena' and str(laiks) == str(ws1['d' + str(row)].value):
            ws1['b' + str(row)].value = vards
            ws1['c' + str(row)].value = veids
            ws1['e' + str(row)].value = audzekna_vards
            break

        elif diena == 'Otrdiena' and str(laiks) == str(ws1['i' + str(row)].value):
            ws1['g' + str(row)].value = vards
            ws1['h' + str(row)].value = veids
            ws1['j' + str(row)].value = audzekna_vards
            break

        elif diena == 'Trešdiena' and str(laiks) == str(ws1['n' + str(row)].value):
            ws1['l' + str(row)].value = vards
            ws1['m' + str(row)].value = veids
            ws1['o' + str(row)].value = audzekna_vards
            break

        elif diena == 'Ceturtdiena' and str(laiks) == str(ws1['s' + str(row)].value):
            ws1['q' + str(row)].value = vards
            ws1['r' + str(row)].value = veids
            ws1['t' + str(row)].value = audzekna_vards
            break

        elif diena == 'Piektdiena' and str(laiks) == str(ws1['x' + str(row)].value):
            ws1['v' + str(row)].value = vards
            ws1['w' + str(row)].value = veids
            ws1['y' + str(row)].value = audzekna_vards
            break

if filiale == 'Mārupe':
    for row in range(3,13):
        if diena == 'Pirmdiena' and str(laiks) == str(ws2['d' + str(row)].value):
            ws2['b' + str(row)].value = vards
            ws2['c' + str(row)].value = veids
            ws2['e' + str(row)].value = audzekna_vards
            break

        elif diena == 'Otrdiena' and str(laiks) == str(ws2['i' + str(row)].value):
            ws2['g' + str(row)].value = vards
            ws2['h' + str(row)].value = veids
            ws2['j' + str(row)].value = audzekna_vards
            break

        elif diena == 'Trešdiena' and str(laiks) == str(ws2['n' + str(row)].value):
            ws2['l' + str(row)].value = vards
            ws2['m' + str(row)].value = veids
            ws2['o' + str(row)].value = audzekna_vards
            break

        elif diena == 'Ceturtdiena' and str(laiks) == str(ws2['s' + str(row)].value):
            ws2['q' + str(row)].value = vards
            ws2['r' + str(row)].value = veids
            ws2['t' + str(row)].value = audzekna_vards
            break

        elif diena == 'Piektdiena' and str(laiks) == str(ws2['x' + str(row)].value):
            ws2['v' + str(row)].value = vards
            ws2['w' + str(row)].value = veids
            ws2['y' + str(row)].value = audzekna_vards
            break


if filiale == 'Ādaži':
    for row in range(3,13):
        if diena == 'Pirmdiena' and str(laiks) == str(ws3['d' + str(row)].value):
            ws3['b' + str(row)].value = vards
            ws3['c' + str(row)].value = veids
            ws3['e' + str(row)].value = audzekna_vards
            break

        elif diena == 'Otrdiena' and str(laiks) == str(ws3['i' + str(row)].value):
            ws3['g' + str(row)].value = vards
            ws3['h' + str(row)].value = veids
            ws3['j' + str(row)].value = audzekna_vards
            break

        elif diena == 'Trešdiena' and str(laiks) == str(ws3['n' + str(row)].value):
            ws3['l' + str(row)].value = vards
            ws3['m' + str(row)].value = veids
            ws3['o' + str(row)].value = audzekna_vards
            break

        elif diena == 'Ceturtdiena' and str(laiks) == str(ws3['s' + str(row)].value):
            ws3['q' + str(row)].value = vards
            ws3['r' + str(row)].value = veids
            ws3['t' + str(row)].value = audzekna_vards
            break

        elif diena == 'Piektdiena' and str(laiks) == str(ws3['x' + str(row)].value):
            ws3['v' + str(row)].value = vards
            ws3['w' + str(row)].value = veids
            ws3['y' + str(row)].value = audzekna_vards
            break

wb.save('Klasu grafiks 2023_2024.xlsx')
wb.close()