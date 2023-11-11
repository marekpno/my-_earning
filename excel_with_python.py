import xlwings as xw
import pandas as pd

'''
data = [['Marek', 32, 'male', 'human'],
        ['Jas', 5, 'male', 'human'],
        ['Monika', 32, 'female', 'human'],
        ['Sinus', 2, 'male', 'dog']]

df = pd.DataFrame(data=data,
                  columns=['Imię', 'wiek', 'płeć', 'gatunek'],
                  index=[100, 101, 102, 104])
xw.view(df)
'''
book = xw.Book()
book.name
book.sheets
print(book.name)
print(book.sheets)
sheet1 = book.sheets[0]
sheet1 = book.sheets["sheet1"]
sheet1.range("A1")
print(sheet1.range("A1"))

sheet1.range("A1").value = [[1,2],[3,4]]
sheet1.range("A4").value = "Siema Eniu"

a = sheet1.range("A1:B2").value
print(a)

b = sheet1.range("A1:B2")[0, 0]
print(b)

c = sheet1.range("A1:B2")[:, 1]
print(c)
#another method
d = sheet1["A1:B2"]
print(d)

e = sheet1[:2,:2]
print(e)


#f = sheet1.range[10:4,11:6]
#print(f)