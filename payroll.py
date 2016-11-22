#-*- coding: utf8 -*-
#配置smtp信息，发送工资条邮件
#v1.1 指定sheet包含“部”，并要求工资条标题栏是整个xls文件都统一标准
#v1.2 通过首列的“序号”，以及“序号”所在行是否包含“邮箱”来判断是否为工资条sheet，且工资条标题栏要单个sheet统一标准即可

from email import encoders
from email.header import Header
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr
import smtplib
import xlrd,threading,queue
import os,sys,time,string,re
import configparser

def _format_addr(s): #格式化邮件信息
    name, addr = parseaddr(s)
    return formataddr((Header(name,'utf-8').encode(),addr))

#@conf   邮件发送配置
#@lock   锁
#@queue  队列数据需包含2个元素(sendto,msg)
#@errAccount  无法发送的邮件账号，需要传入变量名
class Sender(threading.Thread): #发送邮件--线程类对象
    def __init__(self,conf,lock,queue,errAccount):
        self.conf=conf
        self.lock=lock
        self.q=queue
        self.errAccount=errAccount
        super(Sender,self).__init__()
    def run(self):
        server = smtplib.SMTP(self.conf['smtp'], 25) # SMTP协议默认端口是25
##        server.set_debuglevel(1)
        
        try: # 检测登录是否OK?
            server.login(self.conf['from'], self.conf['pwd'])
        except Exception as e:
            self.lock.acquire()
            print(e)
            self.lock.release()
        else:
            while True:
                data=self.q.get()
                sendto=data[0]
                msg=data[1]
                
                try:# 检测邮件发送是否OK?
                    server.sendmail(self.conf['from'], sendto, msg.as_string())
                except Exception as e:
                    errAccount.append(sendto)
                else:
                    self.lock.acquire()
                    print('%-2s%-30s%-25s%s'%("√",sendto,time.time(),self.name))
                    self.lock.release()
                self.q.task_done() #告诉队列取数后的操作已完毕。
                
            server.quit()
            # print("%s is empty"%self.name)
            
        
        

def Ldump(*txt):
    lock.acquire()
    print(*txt)
    lock.release()


#@conf 配置
#@td   td=td[0]:email ,td[1]:name,td[2]=td.data  
def Msg_encode(conf,td):
    content=html_head+"<table>"
    content+=td[2]
    content+="</table>"+html_end
    msg = MIMEText(content, 'html', 'utf-8')
    msg['From'] = _format_addr('财务 <%s>' % conf['from'])
    msg['To'] = _format_addr('%s <%s>' %(td[1],td[0]))
    msg['Subject'] = Header(conf['subject'], 'utf-8').encode()
    return td[0],msg
    

def htmlFile(d):
    fname=r"payroll.html"
    if os.path.exists(fname):
        os.remove(fname) 
    with open(fname,'w') as f:
        content=html_head+"<table>"
        for k,v in d.items():
            for x in v:
                check='''<caption class='msg'>TO:%s-%s<input type="checkbox" style="vertical-align:middle;" ></caption>'''%(x[0],k)
                content+=check+x[2]+"</table><br/><br/><table>"
        content+="</table>"+ html_end
        f.write(content)
    os.startfile(fname)
    


#@rowIndex   工资条标题栏起始行号
#@colName  姓名的列号
#@colMail  邮箱的列号
#@lab     标题栏数据--list（首行，次行）
#@d_lab   标题栏数据--dict（合并的列数，行数）
#@d       本sheet工资条数据
def th_encode(sh):
    cv=sh.col_values(0)
    if "序号" in cv:
        rowIndex=cv.index('序号') #工资条标题栏的行号
        rv=sh.row_values(rowIndex) #rv 获取“序号”所在行的values
        d_lab={} #{"姓名"：[0,1],...} 姓名占n+0列、n+1行(n=1)
        lab=([],[]) #lab= (['姓名', '基本工资', '岗位工资', '绩效工资', ..., '银行发放'], [])
        if "邮箱" in rv:
            colName=rv.index("姓名")  #  工资条标题栏：姓名的列号
            colMail=rv.index("邮箱") #  工资条标题栏：邮箱的列号
            # -- 标题栏首行数据 --

            rv.reverse()
            t=0
            for x in rv:
                if x:
                    d_lab[x]=([t,0])
                    lab[0].append(x)
                    t=0          
                else:
                    t+=1
            lab[0].reverse()
            lab[0].remove("邮箱") #删除 mail
            lab[0].remove("序号") #删除 序号

            # -- 标题栏次行数据 --
            # print(sh.cell_value(rowIndex+1,1))
            if not(sh.cell_value(rowIndex+1,0)): #如果不为空，标题栏为2行
                rv2=sh.row_values(rowIndex+1)
                # rv2.pop(colMail)
                # rv2.pop(0)
                for i,x in enumerate(rv2):
                    if x:
                        lab[1].append(x)
                    else: # -- 统计占两行的标题栏 --
                        v=sh.cell_value(rowIndex,i)
                        if v in lab[0]:
                            d_lab[v][1]=1

            # -- 生成标题栏的html --
            th_html="<tr>"
            for x in lab[0]:
                th_html+="<th "
                if d_lab[x][0]:
                    th_html+="colspan="+str(d_lab[x][0]+1)
                if d_lab[x][1]:
                    th_html+="rowspan="+str(d_lab[x][1]+1)
                th_html+=">"+x+"</th>"
            th_html+="</tr>"
            if lab[1]:
                th_html+="<tr>"
                for x in lab[1]:
                    th_html+="<th>"+x+"</th>"
                th_html+="</tr>"
            return colMail,colName,th_html
        else:
            return False
    else:
        return False
    
#@td  员工工资条数据html--list
#@sh   sheet对象
#@i    邮箱的列号
#@j    姓名的列号
#@d    td数据存储容器
def td_encode(sh,i,j,th):
    d=[]  #{"sheetname":[["mail","name","td.data"],……],……}
    nrows=sh.nrows
    for n in range(nrows):
        rv=sh.row_values(n)
        mail=rv[i]
        data=[] #data=["test@123.com","name","td.data"]
        if isinstance(mail,str) and re.match(r'^(\w+[\-\.]?\w+)@(\w+\-?\w+)(\.\w+)$',mail.strip()):
            data.append(mail)  #append  email
            data.append(rv[j])  #append  name
            td="<tr>"
            rv.pop(i)  #pop email
            rv.pop(0)  #pop 序号
            for y in rv:
                td+="<td>"+("%s"%y)+"</td>"
            td+="</tr>"
            tab=th+td
            data.append(tab)
            d.append(data)
    return d
    

#@colName  姓名的列号
#@colMail  邮箱的列号
#@th_sign 标题栏分析状态（0：未进行 1：完成）--
#@th_html 标题栏html的table格式
#@td_dhtml  员工工资条数据html格式--list
def readXLS(fname):
    # th_sign=0
    bk=xlrd.open_workbook(fname)
    shname=bk.sheet_names()
    for s in shname:
        sh=bk.sheet_by_name(s)
        th=th_encode(sh)
        if th:
            d[s]=td_encode(sh,*th)
            print("%3s%1s%s"%("√","",sh.name))            
def setGlobal():
    global conf,html_head,html_end,q,lock,errAccount
    conf={}
    html_head='''<html>
            <head>
            <meta charset="GBK">
            <title>工资条预览</title>
            <style type="text/css">
            #mainbox {margin:5 auto;}
            .msg{text-align:right;}
            table {border-collapse:collapse;width:88%;margin:0 auto;}
            table,tr,th,td {
                border:1px solid #000;
                text-align:center;
                }
            th{background-color:#eee}

            </style>
            </head>
            <body>
            <div id="mainbox">'''
    html_end='''</div>
            </body>
            </html>'''
    q=queue.Queue()
    lock=threading.Lock()
    errAccount=[]

def SetConf():
    fname="payroll.cfg"
    print('        *--*--*--*--*--*--*--*--*--*--*--*--*')
    print('        |                                   |')
    print('    ----|       工资条发送程序              |----')
    print('        |                                   |')
    print('        *--*--*--*--*--*--*--*--*--*--*--*--*')
    conf={}
    cf = configparser.ConfigParser()
    if os.path.exists(fname):#确认配置文件存在，则读取配置
        cf.read(fname)
        conf['from']=cf.get('smtpset','from')
        conf['pwd']=cf.get('smtpset','pwd')
        # conf['pwd']=cryptcode(cf.get('smtpset','pwd'),1)
        conf['smtp']=cf.get('smtpset','smtp')
    else:
        conf['from']=input("请设置发送邮箱：")
        conf['pwd']=input("请输入邮箱密码：")
        conf['smtp']=input("请设置SMTP：")    
        cf.add_section("smtpset")#增加section 
        cf.set("smtpset", "from", conf["from"])#增加option 
        cf.set("smtpset", "pwd", conf["pwd"])
        # cf.set("smtpset", "pwd", cryptcode(conf["pwd"])) 
        cf.set("smtpset", "smtp", conf["smtp"])  
        with open(fname, "w") as f: 
            cf.write(f)#写入配置文件文件中   
#  -- 设置邮件主题（某年某月工资条） --
    sub=time.strftime('%Y%m',time.localtime())
    month=int(sub[4:6])-1
    if month:
        conf['subject']='%s年%02d月工资条'%(sub[:4],month)
    else:
        conf['subject']='%s年%02d月工资条'%(int(sub[:4])-1,12)

    conf['fxls']=''
    prein="请输入"
    while conf['fxls']=='': 
        conf['fxls']=input("%s%s"%(prein,'工资表格文件：')).lower()
        if not(re.match(r'.*\.xls(x?)$',conf['fxls'])):
            conf['fxls']=''
        elif not(os.path.exists(conf["fxls"])):
            prein=("%s文件不存在!\n请重新输入"%conf["fxls"])
            conf["fxls"]=""
            
            
    print("*"*50)
    print("%s%-10s%s"%("*"*8,'配置信息如下：',"*"*8))
    print("%-10s%s"%('发送邮箱：',conf['from']))
    print("%-10s%s"%('SMTP服务器：',conf['smtp']))
    print("%-10s%s"%('工资条表格文件：',conf['fxls']))
    print("%-10s%s"%('邮件标题：',conf['subject']))
    print("*"*50)
    return conf
def loginTest(conf):
    pass
    # server.login

def iSelect():
    conf['cmd']=''
    inbox='浏览工资条(B)/'
    print("*"*50)
    while not(conf['cmd']): 
        conf['cmd']=input(inbox+'发送邮件(S)/退出(Q):').strip().lower()
        if conf['cmd']=='b':
            cmdBrow()
            time.sleep(2)
            inbox='重新浏览工资条(B)/'
            print(".\n"*2)
            conf['cmd']=''
        elif conf['cmd']=='q':
            sys.exit()
        elif conf['cmd']=="s":
            cmdSend()
        else:
            conf['cmd']=''

# -- 浏览数据 --
def cmdBrow():
    htmlFile(d)

# -- 发送工资条 --            
def cmdSend():
    # print("Send mail is begining……")
    print(".\n"*2)
    print("开始发送电子邮件……")
    # --工资条邮件生成 --
    for k,v in d.items():
        for x in v:
            data=Msg_encode(conf,x)
            q.put(data)
    # -- 工资条邮件多线程发送 --
    Consumer=[Sender(conf,lock,q,errAccount) for i in range(4)]
    start=time.time() #计时开始
    for c in Consumer:
        c.daemon = True
        c.start()
    # -- 单线程测试 --
##    Consumer=Sender(conf,lock,q,errAccount)
##    Consumer.daemon=True
##    Consumer.start()
    q.join()

    # -- 发送完毕，输出结果 --
    # print("\nSend completed!  Total spending time:：%s"%(time.time()-start))
    # print("\nThe following is a list of error accounts：")
    print("\n发送完毕!  总计用时：%s"%(time.time()-start))
    print("\n以下邮件账号发送失败：")
    print("*"*25)
    for x in errAccount:
        print("%-2s%-25s"%("×",x))
    print("*"*25)

if __name__ == '__main__':
    setGlobal()
    if not(conf):
        conf=SetConf()
    pay_label="" #标题栏数据,默认为空
    i_height=1  #标题栏行高，默认为1
##    fname=r"d:\code\js\123.xls"
    fname=conf["fxls"]
    th=[]
    d={}  #@工资数据
    # print("Data analysis is begining……")
    print("数据分析开始……")
    readXLS(fname)
    # print("Data analysis is completed！")
    print("数据分析完毕。")
    iSelect()
    
       


    




 

