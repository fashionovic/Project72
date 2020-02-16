# -*- coding: utf-8 -*-
import xlrd
import json
import os

import shutil



#os.remove("info_.json")

wb = xlrd.open_workbook("calc.xlsx")
#print(wb.sheetnames())

sh = wb.sheet_by_index(0)
for rownum in range(sh.nrows):
    opaque_row = sh.row_values(rownum)
   # print(opaque_row)
    print("**************")

#with open ("sample_input.json") as f:
 #   data = json.load(f)
#print (data)
#print (json.dumps(data, indent=0))

for colnum in range(sh.ncols):
    #opaque_column = {"opaque_column"+str(colnum+1) : sh.col_values(colnum)}
    opaque_column = sh.col_values(colnum)

shutil.copyfile("excel2py.json", "excel2py1.json")
with open("excel2py1.json", "r") as g:
    inputa = json.load(g)
    print(inputa["opaque_column2"])
    g.close()

with open("excel2py1.json", "a") as g:
    inputa = json.load(g)
    for colnum in range(sh.ncols):
        opaque_column = sh.col_values(colnum)
        print(opaque_column)
        inputa["opaque_column2"]=opaque_column[0]
    g.close()











#with open("info_.json", "a") as h:
        #information = inputa
        #h.write(json.dumps(information))
        #for colnum in range(sh.ncols):
            #opaque_column = {"opaque_column" + str(colnum + 2): sh.col_values(colnum)}
            #a=(str(opaque_column).replace("{", "").replace("}", ""))
            #h.write((json.dumps(a)))


        #h.close()











#first_column = sh.col_values(0)
#print(first_column)

#cel_c4 = sh.cell(3, 2).value
#print(cel_c4)
