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

    if (len(result) > 0): # 查找成功
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

    else: # 查找失败
        print("未查找到相关结果")
        sys.exit()

    print("有索引访问内存地址的次数为：", time)
    xls.save(path + SheetName + "-search.xls") # 将查询结果保存到文件中
    return returnResult


# 在建立了索引的基础上，查找所属旅行社编号为agentId的乘客
# 由于原始的乘客信息表是以乘客编号为索引建立的，因此这里需要新建一个以旅行社编号为索引的B+树
# 【这个查询的正确结果应该可以有多个的情况，但是我写的这个始终只能查询到一个结果】

def searchWithIndex_passenger_by_age(age):
    xls = xlwt.Workbook()  # 查询结果表
    searchInfo = xls.add_sheet('sheet1')

    time = 0
    SheetName = "PassengerInforSheet"

    mybptree = Bptree(4, 4)
    nodeList1 = []
    for i in range(11):  # 行数
        key = row1[i][5]  # 乘客的编号作为索引
        value = path + 'PassengerInforSheet' + ',' + str(i) + ',0'
        if i > 0:
            nodeList1.append(KeyValue(key, value))
    for kv in nodeList1:
        mybptree.insert(kv)

    result = mybptree.search(0, age) # 在索引的基础上进行查询

    returnResult = []  # 存放查询结果

    if (len(result) > 0): # 查找成功
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

    else: # 查找失败
        print("未查找到相关结果")

    print("有索引访问内存地址的次数为：", time)
    xls.save(path + SheetName + "-search-pass-age.xls") # 将查询结果保存到文件中
    return returnResult

def searchWithIndex_passenger_by_agentId(agentId):
    xls = xlwt.Workbook()  # 查询结果表
    searchInfo = xls.add_sheet('sheet1')

    time = 0
    SheetName = "PassengerInforSheet"

    mybptree = Bptree(4, 4)
    nodeList1 = []
    for i in range(11):  # 行数
        key = row1[i][1]  # 乘客的编号作为索引
        value = path + 'PassengerInforSheet' + ',' + str(i) + ',0'
        if i > 0:
            nodeList1.append(KeyValue(key, value))
    for kv in nodeList1:
        mybptree.insert(kv)
    result = mybptree.search(agentId, agentId) # 在索引的基础上进行查询

    returnResult = []  # 存放查询结果

    if (len(result) > 0): # 查找成功
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
    else: # 查找失败
        print("未查找到相关结果")

    print("有索引访问内存地址的次数为：", time)
    xls.save(path + SheetName + "-search-pass-agentId.xls") # 将查询结果保存到文件中
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

    if (len(result) > 0): # 查找成功
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

    else: # 查找失败
        print("未查找到相关结果")

    print("有索引访问内存地址的次数为：", time)
    xls.save(path + SheetName + "-search-agent-id.xls") # 将查询结果保存到文件中
    return returnResult

# 在没有建立索引的情况下，查找编号为id的乘客
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
        if(int(data[i][column] )< age):
            result.append(data[i])
    print("无索引（遍历）访问内存地址的次数为：", time)
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
    print("无索引（遍历）访问内存地址的次数为：", time)
    return result

if __name__ == '__main__':
    #针对乘客信息表的查询
    print("开始有索引的查询------------------")
    result1 = searchWithIndex_passenger_by_age(25)
    print(" age < 25的 的乘客的查询结果为：")
    print(result1)
    print("开始无索引的查询------------------")
    result2 = search_passenger_by_age(25)
    print(" age < 25的 的乘客的查询结果为：")
    print(result2,'\n')

    print("开始有索引的查询------------------")
    result3 = searchWithIndex_passenger_by_agentId('agent005')
    print(" agent005 的 的乘客的查询结果为：")
    print(result3,'\n')

    #针对旅行社信息表的查询
    print("开始有索引的查询------------------")
    result4 = searchWithIndex_agent_by_id('agent004',mybptree2)
    print("编号为 agent004 的旅行社的查询结果为：")
    print(result4)
    print("开始无索引的查询------------------")
    result5 = search_agent_by_id('agent004')
    print("编号为 agent004 的旅行社的查询结果为：")
    print(result5)