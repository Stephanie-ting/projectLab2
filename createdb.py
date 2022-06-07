# create db
import xlwt
from btree import *

path = 'D:/数据库/project2/projectLab2/sheet/'
xls1 = xlwt.Workbook()  # 巡航信息表
xls2 = xlwt.Workbook()  # 邮轮信息表

passengerInfo = xls1.add_sheet('sheet1')
agentInfor = xls2.add_sheet('sheet1')

fileName = ['PassengerInforSheet', 'AgentInforSheet']
mybptree1 = Bptree(4, 4)
mybptree2 = Bptree(4, 4)

treeSet = [mybptree1, mybptree2]
'''''创建表格，并添加数据项'''


row1 = [['p_id', 't_id', 'name', 'gender', 'phone_number','age'],
        ['passenger001', 'agent001', 'Sam Smith', 'Male', '36347889',19],
        ['passenger002', 'agent001', 'Charlie Puth', 'Male', '36347888',18],
        ['passenger003', 'agent001', 'Mary Darling', 'Female', '21113446',23],
        ['passenger004', 'agent001', 'Harry Potter', 'Male', '16688888',32],
        ['passenger005', 'agent001', 'Tom Lierd', 'Male', '26712136',20],
        ['passenger006', 'agent005', 'Peter', 'Male', '29991999',27],
        ['passenger007', 'agent005', 'Sara', 'Female', '18902200',45],
        ['passenger008', 'agent005', 'Lee', 'Female', '20001100',47],
        ['passenger009', 'agent009', 'Wu', 'Female', '13332333',30],
        ['passenger010', 'agent010', 'Xu', 'Male', '18888888',39]]

row2 = [['t_id', 'name'],
        ['agent001', 'tuniu'],
        ['agent002', 'xiecheng'],
        ['agent003', 'qinglv'],
        ['agent004', 'tianya'],
        ['agent005', 'xingyun'],
        ['agent006', 'yuanyang']]

# 构造一个空的B+树
# isExists = os.path.exists(path + 'PassengerInforSheet')
# if not isExists: os.makedirs(path + 'PassengerInforSheet')
# isExists = os.path.exists(path + 'AgentInforSheet')
# if not isExists: os.makedirs(path + 'AgentInforSheet')


# 乘客信息
nodeList1 = []
for i in range(11):  # 行数
    key = row1[i][0]  # 乘客信息的 p_id 作为索引
    value = path + 'PassengerInforSheet' + ',' + str(i) + ',0'
    if i > 0:
        nodeList1.append(KeyValue(key,value))
    for j in range(6):  # 列数
        passengerInfo.write(i, j, row1[i][j])
for kv in nodeList1:
    mybptree1.insert(kv)
# mybptree1.show()


# 旅行社信息
nodeList2 = []
for i in range(7):  # 行数
    key = row2[i][0]  #旅行社信息的 t_id 作为索引
    value = path + 'AgentInforSheet' + ',' + str(i) + ',0'
    if i > 0:
        nodeList2.append(KeyValue(key,value))
    for j in range(2):  # 列数
        agentInfor.write(i, j, row2[i][j])
for kv in nodeList2:
    mybptree2.insert(kv)
# mybptree2.show()


xls1.save(path + 'PassengerInforSheet.xls')
xls2.save(path + 'AgentInforSheet.xls')


