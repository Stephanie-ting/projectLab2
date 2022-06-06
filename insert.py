import xlrd
import xlutils.copy
from createdb import *

def insert(SheetName,insertList,bptree):
    mybptree = bptree
    print("原始B+树-------------------------\n")
    mybptree.show()

    origin = xlrd.open_workbook(path+SheetName+'.xls',formatting_info=True)
    newFile = xlutils.copy.copy(origin)
    sheet = newFile.get_sheet(0)
    rowNum = origin.sheet_by_name('sheet1').nrows
    colNum = origin.sheet_by_name('sheet1').ncols
    for i in range(rowNum,rowNum+len(insertList)):
        for j in range(colNum):
            sheet.write(i,j,insertList[i-rowNum][j])
    newFile.save(path+SheetName+'-insert.xls')

    newNodeSet=[]
    for i in range(len(insertList)):
        newNodeSet.append(KeyValue(insertList[i][0],path+'AgentInforSheet' + ',' + str(i+rowNum) + ',0'))

    for kv in newNodeSet:
        mybptree.insert(kv)
    print("insert后的B+树-------------------------\n")
    mybptree.show()

if __name__ == '__main__':
    SheetName = 'AgentInforSheet'
    insertList = [['agent007', 'yy'],
                  ['agent008', 'zz'],
                  ['agent009', 'xx']]
    insert(SheetName,insertList,mybptree2)