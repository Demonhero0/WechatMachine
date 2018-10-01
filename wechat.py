import itchat
import time
import xlrd

def read():
    file = xlrd.open_workbook('二班名单.xlsx')#formatting_info=True
    table = file.sheet_by_index(0)
    data = []
    for i in range(1,55):
        number = int(table.cell_value(i,0))
        name = table.cell_value(i,1)
        turn = table.cell_value(i,2)
        temp = {'number':number,'name':name,'date':None,'turn':turn}
        data.append(temp)
    return data

def getChatroom(name):
    for room in itchat.get_chatrooms():
        if room['NickName'] == name:
            return room['UserName']

send = 1

@itchat.msg_register(itchat.content.TEXT,isGroupChat=True,isFriendChat=True)
def LabourBot(msg):
    '''
    hour = time.localtime(time.time())[3]
    if hour >= 19 and send:
        itchat.send_msg(msg='Text Message',toUserName=getChatroom('求求你们都别学了'))
        send = 0
    if hour < 19:
        send = 1
    print(msg['Text'])
    '''
    if msg['Text'] == '我会好好擦黑板':
        userName = msg.get('ActualUserName',None)

        student = itchat.search_friends(userName=userName)
        for stu in read():
            if stu['name'] == student.get('RemarkName',None):
                stu['date'] = time.strftime("%a-%m-%d-%Y", time.localtime()) 



#    if(msg['ActualUserName']==&&msg['Text']=='我会好好擦黑板')
itchat.auto_login(hotReload = True)
#print(itchat.search_chatrooms(name="数院2017级二班"))

itchat.run()
