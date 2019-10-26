import xlwings as xw
import pandas as pd

df = pd.read_excel('xlwingsProject.xlsx', sheet_name="Sheet1",index_col=0)
print(df)
print(df["sourceLanguageText"])
print(df["sourceLanguageText"].values)
print(df.loc["key3"])
print(df.loc["key3"].values)

app = xw.App(visible=False)
PIDlist = xw.apps.keys()
print("PID:\n{}".format(PIDlist))

wb = xw.Book('xlwingsProject.xlsx')  # connect to an existing file in the current working directory

PIDlist = xw.apps.keys()
print("PID:\n{}".format(PIDlist))

sh = wb.sheets['Sheet1']

# automagically find the size of the data on a sheet using expand
tablerange = sh.range("A1").expand("table")
print(tablerange.value)

range = sh.range("A1:C6")

print(range.value)
wb.close()

df.set_value('key3','sourceLanguageText','Benson text string 3')
print(df)
with pd.ExcelWriter('xlwingsProject.xlsx') as writer:
    df.to_excel(writer)


TestData = [
    ["key1","string text 1","type1"],
    ["key2","string text 2","type2"],
    ["key3","string text 3","type3"],
    ["key4","string text 4","type4"],
    ["key5","string text 5","type5"]
]

testdf = pd.DataFrame(TestData,columns=['identifierName','sourceLanguageText','stringType'])
testdf = testdf.set_index('identifierName')
print(testdf)

with pd.ExcelWriter('xlwingsProject.xlsx') as writer:
    testdf.to_excel(writer)