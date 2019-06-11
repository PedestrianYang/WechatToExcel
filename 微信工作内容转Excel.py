#coding=utf8


import itchat
from itchat.content import TEXT, SHARING, PICTURE
import xlrd
import xlwt
from xlutils.copy import copy
import datetime
import os.path
import pandas as pd
path = '/Users/iyunshu/Documents/'


class Wirter(object):
    def __init__(self):
        #生成Excel表格
        now = datetime.datetime.now()
        self.filename = path + '微信平台工作日志_' + now.strftime( '%Y-%m-%d') + ".xls"
        if not os.path.exists(self.filename):
            self.f = xlwt.Workbook() #创建工作簿
            self.creatExcel()

        oldWb = xlrd.open_workbook(self.filename, formatting_info=True)
        self.newWb = copy(oldWb)
        self.newSheet = self.newWb.get_sheet(0)


    def creatExcel(self):
        self.sheet = self.f.add_sheet('sheet1',cell_overwrite_ok=True)

        print(self.sheet.get_rows)
        row0 = ['发送人','工作动态','发送内容', '时间' , '字数']

        for i in range(0,len(row0)):
            self.sheet.write(0, i, row0[i], self.set_style('Times New Roman',220,True))

        self.f.save(self.filename)


    def set_style(self, name,height,bold=False):
        style = xlwt.XFStyle()  # 初始化样式

        font = xlwt.Font()  # 为样式创建字体
        font.name = name # 'Times New Roman'
        font.bold = bold
        font.color_index = 4
        font.height = height
        style.font = font
        return style

    def writeData(self, name, content):
        now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        strlen = len(content)

        name = name[0:2] + '办事处'
        rowContent = [name,'工作动态',content, now, strlen]

        name_col = self.newSheet.col(0)
        name_col.width = 400 * 7

        content_col = self.newSheet.col(2)
        content_col.width = 400 * strlen

        name_col = self.newSheet.col(3)
        name_col.width = 400 * 15

        rowindex = len(self.newSheet.get_rows())

        for i in range(0,len(rowContent)):
            self.newSheet.write(rowindex, i, rowContent[i], self.set_style('Times New Roman',220,True))
        self.newWb.save(self.filename)



    def filterAndMerge(self):
        df = pd.read_excel(self.filename)
        df=pd.DataFrame(df,columns=['发送人','工作动态','发送内容', '时间' , '字数'])
        df = df.sort_values(by='发送人')#排序
        df=df.reset_index(drop=True)#重建索引
        df.to_excel(self.filename, index=False)
        print('文件保存目录：--->',self.filename)





@itchat.msg_register([TEXT, SHARING, PICTURE], isGroupChat=True)
def group_reply_text(msg):
    print(msg['User']['NickName'])
    msg_gourpName = msg['User']['NickName']
    if msg_gourpName == groupName:
        goup_username = msg['FromUserName']
        goup_nickname = msg['ActualNickName'] #群昵称
        goup_Content = msg['Content']
        if not goup_username == myUserName:
            print('['+ goup_nickname + ']发来消息：' + goup_Content)
            writer.writeData(goup_nickname, goup_Content)
        writer.filterAndMerge()




if __name__ == '__main__':
    writer = Wirter()

    itchat.auto_login()
    # 获取自己的UserName
    myUserName = itchat.get_friends(update=True)[0]['UserName']
    # groupName = '经开区城市精细化管理工作群'
    groupName = '哈哈哈'
    itchat.run()