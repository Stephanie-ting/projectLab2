import xlrd
import xlutils.copy
from createdb import *
import sys

def delete(key,bptree):
    mybptree = bptree
    print("原始B+树-------------------------\n")
    mybptree.show()
    result = mybptree.search(key,key)
    if(len(result)==0):
        print("没有要删除的元组")
        sys.exit()

    rawName = result[0].value
    sheetName = rawName.split(',')[0].split('/')[-1]
    row=rawName.split(',')[1]

    origin = xlrd.open_workbook(path+sheetName+'.xls',formatting_info=True)
    newFile = xlutils.copy.copy(origin)
    sheet = newFile.get_sheet(0)
    colNum = origin.sheet_by_name('sheet1').ncols

    for k in range(len(result)):
        for col in range(colNum):
            sheet.write(int(row),col,'')

    for kv in result:
        mybptree.delete(kv)
    print("delete后的B+树-------------------------\n")
    mybptree.show()
    newFile.save(path+sheetName+'-delete.xls')

if __name__ == '__main__':
    delete('agent002',mybptree2)