import itchat
import time
import xlrd
import xlwt

send = 1
state = 0

def read():
    file = xlrd.open_workbook('17级2班擦黑板.xlsx')#formatting_info=True
    table = file.sheet_by_index(0)
    data = []
    for i in range(1,55):
        number = int(table.cell_value(i,0))
        name = table.cell_value(i,1)
        turn = int(table.cell_value(i,2))
        temp = {'number':number,'name':name,'date':None,'turn':turn}
        data.append(temp)
    return data

data = read()

def save(data,sheet=0,row=2,col=1):
    xls = xlrd.open_workbook('17级2班擦黑板.xlsx')
    xlsc = copy(xls)
    sheet1 = xlsc.get_sheet(sheet)
    (r,c) = data.shape
    for i in range(r):
        for j in range(c):
            sheet1.write(i + row - 1, j + col - 1, float(data[i][j]))
    xlsc.save(path)

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
        students = [data[state]['name'],data[state+1]['name']]
        #print(students)
        message = students[0] + '和' + students[1] + '同学，明天将轮到你们擦黑板。请回复“我会好好擦黑板”，否则将视作缺勤。'
        #print(message)
        itchat.send_msg(msg=message,toUserName=getChatroom('test'))
        for stu in data:
            if stu.get('name',None) in students:
                stu['turn'] += 1
        print(data[state],data[state+1])

        send = 0
        state = (state + 2) % 54
    if hour < 0:
        send = 1
    #print(msg['Text'])

    if msg['Text'] == '我会好好擦黑板':
        data = read()
        userName = msg.get('ActualUserName',None)
        students = [data[state-2]['name'],data[state-1]['name']]
        print(students)

        student = itchat.search_friends(userName=userName)
        if student.get('RemarkName',None) in students:
            for stu in data:
                if stu['name'] == student.get('RemarkName',None):
                    stu['date'] = time.strftime("%m%d%Y", time.localtime())
        print(data[state-2],data[state-1])

    if '$' in msg['Text']:
        flag = 0
        data = read()
        names = msg['Text'].split('$')          # 0-本人 1-他人
        print(names)

        userName = msg.get('ActualUserName',None)
        students = [data[state-2]['name'],data[state-1]['name']]
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



#    if(msg['ActualUserName']==&&msg['Text']=='我会好好擦黑板')
itchat.auto_login(hotReload = True)
#print(itchat.search_chatrooms(name="数院2017级二班"))
itchat.run()
