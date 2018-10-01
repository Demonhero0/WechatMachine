import itchat
import time
import xlrd
import xlwt

send = 1
state = 0

def read():
    file = xlrd.open_workbook('二班名单.xlsx')#formatting_info=True
    table = file.sheet_by_index(0)
    data = []
    for i in range(54):
        number = int(table.cell_value(i,0))
        name = table.cell_value(i,1)
        temp = {'number':number,'name':name,'date':None}
        data.append(temp)
    return data

data = read()

def save(data,path,sheet,row,col):
    xls = xlrd.open_workbook(path)
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
    hour = time.localtime(time.time())[3]
    if hour >= 19 and send == 1:
        message = data[state]['name'] + '和' + data[state + 1]['name'] + '同学，明天将轮到你们擦黑板。请回复“我会好好擦黑板”，否则将视作缺勤。'
        itchat.send_msg(msg=message,toUserName=getChatroom('求求你们都别学了'))
        send = 0
        state = (state + 2) % 53
    if hour < 19:
        send = 1
    print(msg['Text'])
    
#    if(msg['ActualUserName']==&&msg['Text']=='我会好好擦黑板')
itchat.auto_login(enableCmdQR = True, hotReload = True)
#print(itchat.search_chatrooms(name="数院2017级二班"))
itchat.run()
