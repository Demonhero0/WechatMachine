import itchat
import time
import xlrd
import xlwt
from xlutils.copy import copy

send = 1
state = 0

def read():
    global state
    file = xlrd.open_workbook('17级2班擦黑板.xls')#formatting_info=True
    table = file.sheet_by_index(0)
    data = []
    for i in range(1,55):
        number = int(table.cell_value(i,0))
        name = table.cell_value(i,1)
        turn = int(table.cell_value(i,2))
        date = []
        for j in range(table.ncols - 3):
            date.append(int(table.cell_value(i,j + 3)))
        temp = {'number':number,'name':name,'turn':turn,'date':date}
        data.append(temp)
    #for i in range(0,53):
    #    if data[i]['turn'] > data[i + 1]['turn']:
    #        state = i + 1
    return data


def save(data,sheet=0,row=2,col=1):
    xls = xlrd.open_workbook('17级2班擦黑板.xls')
    xlsc = copy(xls)
    sheet1 = xlsc.get_sheet(0)
    for i in range(54):
        j = 0
        for t in data[i].values():
            if j == 3:
                for n in range(len(t)):
                    sheet1.write(i + 1, j + n, t[n])
            else:
                sheet1.write(i + 1, j, t)
                j += 1
    xlsc.save('17级2班擦黑板.xls')

def getChatroom(name):
    for room in itchat.get_chatrooms():
        if room['NickName'] == name:
            return room['UserName']


@itchat.msg_register(itchat.content.TEXT,isGroupChat=True)
def LabourBot(msg):
    global send
    global state
    exchange = []

    data = None
    hour = time.localtime(time.time())[3]
    if hour >= 0 and send == 1:
        data = read()
        students = [data[state]['name'],data[state + 1]['name']]
        #print(students)
        message = students[0] + '和' + students[1] + '同学，明天将轮到你们擦黑板。请回复“我会好好擦黑板”，否则将视作缺勤。'
        #print(message)
        itchat.send_msg(msg=message,toUserName=getChatroom('test'))
        for stu in data:
            if stu.get('name',None) in students:
                stu['turn'] += 1
        print(data[state],data[state + 1])

        send = 0
        state = (state + 2) % 54
    if hour < 0:
        send = 1
    #print(msg['Text'])

    if msg['Text'] == '我会好好擦黑板':
        data = read()
        userName = msg.get('ActualUserName',None)
        students = [data[state - 2]['name'],data[state - 1]['name']]
        print(students)

        student = itchat.search_friends(userName=userName)
        if student.get('RemarkName',None) in students:
            for stu in data:
                if stu['name'] == student.get('RemarkName',None):
                    stu['date'] = [int(time.strftime("%m%d%Y", time.localtime()))]
        print(data[state - 2],data[state - 1])

    if '$' in msg['Text']:
        flag = 0
        data = read()
        names = msg['Text'].split('$')          # 0-本人 1-他人
        print(names)

        userName = msg.get('ActualUserName',None)
        students = [data[state - 2]['name'],data[state - 1]['name']]
        print(students)

        student = itchat.search_friends(userName=userName)
        for stu in data:
            if stu['name'] == names[1] and stu['name'] == names[0]:
                flag = 1


        if flag == 0:
            itchat.send_msg(msg="请输入正确名字",toUserName=getChatroom('test'))

        if student.get('RemarkName',None) in students and student.get('RemarkName',None) == names[0] and flag == 1:
            exchange = names
            print(exchange)
    if data is not None:
        save(data)


#    if(msg['ActualUserName']==&&msg['Text']=='我会好好擦黑板')
itchat.auto_login(enableCmdQR=True, hotReload = True)
#print(itchat.search_chatrooms(name="数院2017级二班"))
itchat.run()
