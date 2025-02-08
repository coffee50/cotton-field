import openpyxl as xl
import random
wb = xl.Workbook()
sheet = wb.active
workerAmount = random.randint(1,100)
for _field in range(workerAmount):
    colChoice = random.choice('abcdefghij')
    rowChoice = random.randint(1, 10)
    sheet[str(colChoice) + str(rowChoice)] = 'nigger'
cottonAmount = 0
rows = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j']
for placedWorkerCol in range(100):
    placedWorkerCol += 1
    for i in range(10):
        if sheet[rows[i] + str(placedWorkerCol)].value == 'nigger':
            cottonAmount += random.randint(1, 6)
print('Cotton collected by workers:', cottonAmount, 'kg')
if cottonAmount >= 550:
    print('Good work.')
elif 420 <= cottonAmount < 550:
    print('Let it be.')
elif 270 <= cottonAmount < 420:
    print('Bad!! You need to work more!')
elif 150 <= cottonAmount < 270:
    print('Pathetic niggers! Work much more!!!')
elif 60 <= cottonAmount < 150:
    print('Fuck these niggers! Take the whips!')
elif cottonAmount < 60:
    print('Fuck these stupid niggers! Kill all niggers!!!')
wb.save('cotton-field.xlsx')
