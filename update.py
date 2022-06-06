import xlrd
import xlutils.copy
from createdb import *
import sys

def update(bptree,key,column,newV):
    mybptree = bptree
    result = mybptree.search(key,key)
    if (len(result) == 0):
        print("没有找到更新的元组")
        sys.exit()
    rawName = result[0].value
    sheetName = rawName.split(',')[0].split('/')[-1]
    row = rawName.split(',')[1]

    origin = xlrd.open_workbook(path+sheetName+'.xls',formatting_info=True)
    newFile = xlutils.copy.copy(origin)
    sheet = newFile.get_sheet(0)

    sheet.write(int(row),column,newV)
    newFile.save(path+sheetName+'-update.xls')

if __name__ == '__main__':
    update(mybptree2,'agent005',1,'xyz')