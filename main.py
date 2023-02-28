from os import listdir
import re
import xlwings as xw

pattern = re.compile('_(\d+)_')

try:
    files = sorted(
        filter(
            lambda file: file.endswith('.xlsx') and file != 'Накладные.xlsx',
            listdir()
        ),
        key=lambda file: int(pattern.search(file).groups()[0])
    )
except Exception as e:
    print('Files reading error.', e)

try:
    with xw.App() as app:
        new_wb = app.books.add()
        for file in files:
            app.books.open(file).sheets[-1].copy(
                after=new_wb.sheets[-1],
                name=pattern.search(file).groups()[0]
            )
        new_wb.sheets[0].delete()    
        new_wb.save(path='Накладные.xlsx')
except Exception as e:
    print('Main cycle error.', e)

input('Enter to exit.')
