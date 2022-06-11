import xlrd
from createdb import *
import sys

# 在建立了索引的基础上，查找编号为id的乘客
def searchWithIndex_passenger_by_id(id,bptree):
    xls = xlwt.Workbook()  # 查询结果表
    searchInfo = xls.add_sheet('sheet1')

    time = 0
    SheetName = "PassengerInforSheet"
    mybptree = bptree
    result = mybptree.search(id, id) # 在索引的基础上进行查询
    returnResult = []  # 存放查询结果

    if (len(result) > 0): # 查找到了
        rawName = result[0].value
        sheetName = rawName.split(',')[0].split('/')[-1]  # 获取文件名：取第一个值
        file = xlrd.open_workbook(path + sheetName + '.xls', formatting_info=True)
        sheet = file.sheet_by_index(0)  # 获得第一个sheet的对象
        time += 1
        value = result[0].value
        row = int(value.split(',')[1])
        returnResult.append(sheet.row_values(row))

        for j in range(6):
            searchInfo.write(0, j, returnResult[0][j])

    else: # 没有找到
        print("未查找到相关结果")
        sys.exit()

    print("在B+树索引上进行查询，访问内存的次数为：", time)
    xls.save(path + SheetName + "-search-pass-id.xls") # 将查询结果保存到文件中
    return returnResult


# 在建立了索引的基础上，按age的特定谓词查找乘客
def searchWithIndex_passenger_by_age(age):
    xls = xlwt.Workbook()  # 查询结果表
    searchInfo = xls.add_sheet('sheet1')

    time = 0
    SheetName = "PassengerInforSheet"

    mybptree = Bptree(4, 4)
    nodeList1 = []
    for i in range(11):  # 行数
        key = row1[i][5]  # 乘客的age作为索引
        value = path + 'PassengerInforSheet' + ',' + str(i) + ',0'
        if i > 0:
            nodeList1.append(KeyValue(key, value))
    for kv in nodeList1:
        mybptree.insert(kv)

    result = mybptree.search(0, age) # 在age索引的基础上进行查询

    returnResult = []  # 存放查询结果

    if (len(result) > 0): # 查找到符合条件的元组
        rawName = result[0].value
        sheetName = rawName.split(',')[0].split('/')[-1]  # 获取文件名：取第一个值
        file = xlrd.open_workbook(path + sheetName + '.xls', formatting_info=True)
        sheet = file.sheet_by_index(0)  # 获得第一个sheet的对象

        for i in range(len(result)):
            time += 1
            value = result[i].value
            row = int(value.split(',')[1])
            returnResult.append(sheet.row_values(row))

            for j in range(6):
                searchInfo.write(i, j, returnResult[i][j])

    else: # 没有查到符合条件的元组
        print("未查找到相关结果")
        sys.exit()

    print("在B+树索引上进行查询，访问内存的次数为：", time)
    xls.save(path + SheetName + "-search-pass-age.xls") # 将查询结果保存到文件中
    return returnResult


# 在建立了索引的基础上，查找编号为id的旅行社
def searchWithIndex_agent_by_id(id,bptree):
    xls = xlwt.Workbook()  # 查询结果表
    searchInfo = xls.add_sheet('sheet1')

    time = 0
    SheetName = "AgentInforSheet"
    mybptree = bptree
    result = mybptree.search(id, id) # 在索引的基础上进行查询
    returnResult = []  # 存放查询结果

    if (len(result) > 0):  # 查找到符合条件的元组
        rawName = result[0].value
        sheetName = rawName.split(',')[0].split('/')[-1]  # 获取文件名：取第一个值
        file = xlrd.open_workbook(path + sheetName + '.xls', formatting_info=True)
        sheet = file.sheet_by_index(0)  # 获得第一个sheet的对象
        time += 1
        value = result[0].value
        row = int(value.split(',')[1])
        returnResult.append(sheet.row_values(row))

        for j in range(2):
            searchInfo.write(0, j, returnResult[0][j])

    else: # 没有查找到符合条件的元组
        print("未查找到相关结果")
        sys.exit()

    print("在B+树索引上进行查询，访问内存的次数为：", time)
    xls.save(path + SheetName + "-search-agent-id.xls") # 将查询结果保存到文件中
    return returnResult

# 在没有建立索引的情况下，查找编号为id的乘客
def search_passenger_by_id(id):
    time = 0
    column = 0
    SheetName = "PassengerInforSheet"
    file = xlrd.open_workbook(path + SheetName + '.xls', formatting_info=True)
    sheet = file.sheet_by_index(0)  # 获得第一个sheet的对象

    data = []
    for each_row in range(sheet.nrows): # 循环保存每一行的数据
        data.append(sheet.row_values(each_row))

    # 顺序查找（无索引）
    result = []
    for i in range(1,len(data)):
        time+=1
        if(data[i][column]== id):
            result.append(data[i])
    print("无索引查找，访问内存的次数为：", time)
    return result

def search_passenger_by_age(age):
    time = 0
    column = 5
    SheetName = "PassengerInforSheet"
    file = xlrd.open_workbook(path + SheetName + '.xls', formatting_info=True)
    sheet = file.sheet_by_index(0)  # 获得第一个sheet的对象

    data = []
    for each_row in range(sheet.nrows): # 循环保存每一行的数据
        data.append(sheet.row_values(each_row))

    # 顺序查找（无索引）
    result = []
    for i in range(1,len(data)):
        time+=1
        if(int(data[i][column] )<= age):
            result.append(data[i])
    print("无索引查找，访问内存的次数为：", time)
    return result

# 在没有建立索引的情况下，查找编号为id的旅行社
def search_agent_by_id(id):
    time = 0
    column = 0
    SheetName = "AgentInforSheet"
    file = xlrd.open_workbook(path + SheetName + '.xls', formatting_info=True)
    sheet = file.sheet_by_index(0)  # 获得第一个sheet的对象

    data = []
    for each_row in range(sheet.nrows): # 循环保存每一行的数据
        data.append(sheet.row_values(each_row))

    # 顺序查找（无索引）
    result = []
    for i in range(1,len(data)):
        time+=1
        if(data[i][column] == id):
            result.append(data[i])
    print("无索引查找，访问内存的次数为：", time)
    return result

if __name__ == '__main__':
    #针对乘客信息表的查询
    print("开始有索引的查询------------------")
    result = searchWithIndex_passenger_by_id('passenger004', mybptree1)
    print(" 乘客p_id 为 passenger004 的查询结果为：")
    print(result)
    print("开始无索引的查询------------------")
    result0 = search_passenger_by_id('passenger004')
    print(" 乘客p_id 为 passenger004 的查询结果为：")
    print(result0,'\n')

    print("开始有索引的查询------------------")
    result1 = searchWithIndex_passenger_by_age(20)
    print(" age <= 20 的乘客的查询结果为：")
    print(result1)
    print("开始无索引的查询------------------")
    result2 = search_passenger_by_age(20)
    print(" age <= 20 的乘客的查询结果为：")
    print(result2,'\n')


    #针对旅行社信息表的查询
    print("开始有索引的查询------------------")
    result4 = searchWithIndex_agent_by_id('agent004',mybptree2)
    print("编号为 agent004 的旅行社的查询结果为：")
    print(result4)
    print("开始无索引的查询------------------")
    result5 = search_agent_by_id('agent004')
    print("编号为 agent004 的旅行社的查询结果为：")
    print(result5)