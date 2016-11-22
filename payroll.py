# Version 2.0.3 by yjs 2016/11/21
# 工资条批量发送工具
from email import encoders
from email.header import Header
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr
import smtplib
import xlrd,threading,queue
import os,sys,time,string,re
import configparser
import base64,webbrowser
import tkinter
import tkinter.messagebox
import tkinter.filedialog

# ----------------------------tkinter-------------------------------------------- #
def main():
    global root,xMenu
    SetGlobal()
    root = tkinter.Tk()
    root.title('工资条批量发送工具')
    #root.geometry("500x300") #geometry: 设置窗口大小 [width]x[heght] 的字符串

    xMenu=tkinter.Menu(root)
    xMenu.add_command(label='账号',command=ToConf)
    xMenu.add_command(label='文件',command=ToFile)
    xMenu.add_command(label='退出',command=root.destroy)
    
    root.config(menu=xMenu) #加载菜单
    frame = tkinter.Frame(root)
    frame.grid()
    file_widget(frame)
    smtp_widget(frame)
    command_widget(frame)
    msg_widget(frame)
    ReadCfg()
    ToFile() if conf else ToConf()
    root.mainloop()

#清空Entry控件内的文本
def delen(*en_arg):
    [en.delete(0,len(en.get())) for en in en_arg]

#设置控件的状态：0=disabled，1=normal
def setState(s=0,*arg): 
    st='normal' if s else 'disabled'   
    [x.config(state=st) for x in arg]
        

def ToConf():
    root.geometry('300x200')
    root.maxsize(350,250)
    xMenu.activate(2)
    xMenu.entryconfig(2,{'state':'normal'})
    xMenu.entryconfig(1,{'state':'disabled'})
    file_frame.grid_remove()   
    cmd_frame.grid_remove()
    show_lab.config(wraplength=250,fg='black')
    smtp_frame.grid()
    show_lab['text']=''
    if conf.get('smtp'):
        delen(en_user,en_pwd)
        en_user.insert(0,conf['user'])
        en_pwd.insert(0,conf['pwd'])
        setState(0,en_user,en_pwd)
        btn_smtp.config(text='修改',command=chsmtp,fg='blue')
        
                       
            
def ToFile():
    root.geometry('400x400')
    root.maxsize(500,500)

    #xMenu.entrycget(2,'label') 获取主界面的label标签
    xMenu.entryconfig(2,{'state':'disabled'})
    xMenu.entryconfig(1,{'state':'normal'})
    smtp_frame.grid_remove()
    cmd_frame.grid() if en_file['state']=='disabled' else cmd_frame.grid_remove()
    delen(en_subject)
    en_subject.insert(0,'%4d/%2d'%gettime())
    show_lab.config(text='',wraplength=400,fg='black')
    file_frame.grid()
    
def gettime():
    t=time.localtime()
    y=t.tm_year
    m=t.tm_mon-1
    if m==0:
        return y-1,12
    else:
        return y,m
    
def file_widget(frame):
    global file_frame,en_file,en_subject,btn_fok,btn_browser 
    file_frame=tkinter.Frame(frame)
    file_frame.grid()
    lab_subject=tkinter.Label(file_frame,text='月份：')
    lab_subject.grid(row=0,column=0)
    en_subject=tkinter.Entry(file_frame,width=10)
    en_subject.grid(row=0,column=1,sticky='w')
    lab_file=tkinter.Label(file_frame,text='工资表:')
    lab_file.grid(row=1,column=0)
    en_file=tkinter.Entry(file_frame,width=50)
    en_file.grid(row=1,column=1,columnspan=4)
    btn_browser=tkinter.Button(file_frame,text='选择',width=10,command=getfile)
    btn_browser.grid(row=1,column=4,sticky='e')
    btn_fok=tkinter.Button(file_frame,text='确定',width=10,command=fok_onclick)
    btn_fok.grid(row=2,column=2)

    
def smtp_widget(frame):
    global en_user,en_pwd,btn_smtp,smtp_frame    
    smtp_frame=tkinter.Frame(frame)
    smtp_frame.grid()
    tkinter.Label(smtp_frame,text="发件人账号：").grid(sticky='w')
    conf_frame=tkinter.Frame(smtp_frame)
    conf_frame.grid()
    tkinter.Label(conf_frame, text="User:").grid(row=1,column=0,sticky='e')
    en_user=tkinter.Entry(conf_frame,width=25)
    en_user.grid(row=1,column=1)

    tkinter.Label(conf_frame, text="Password:").grid(row=2,column=0,sticky='e')
    en_pwd=tkinter.Entry(conf_frame, show='*',width=25)
    en_pwd.grid(row=2,column=1,sticky='w')

    #点击按钮运行命令getmsg获取输入框输入信息
    btn_smtp=tkinter.Button(smtp_frame,text='登录测试',width=10, command=getsmtp)
    btn_smtp.grid()

def command_widget(frame):
    global cmd_frame,btn_preview,btn_send,btn_quit
    cmd_frame=tkinter.Frame(frame)
    btn_view=tkinter.Button(cmd_frame,text='预览工资条',width=10,command=viewFile)
    btn_view.grid(row=0,column=0)
    btn_send=tkinter.Button(cmd_frame,text='发送工资条',width=10,command=sureSend)
    btn_send.grid(row=0,column=1)



def msg_widget(frame):
    global msg_frame,show_lab
    msg_frame=tkinter.Frame(frame)
    msg_frame.grid()
    show_lab=tkinter.Label(msg_frame,wraplength=500,justify='left')
    show_lab.grid()
    
    
def getfile(): #打开文件对话框获取文件名
    fxls=tkinter.filedialog.askopenfilename()
    en_file.delete(0,len(en_file.get()))
    en_file.insert(0,fxls)
    setState(1,btn_fok)

def fok_onclick():
    strmsg=''
    if btn_fok['text']=='修改':
        if conf.get('fxls'):conf.pop('fxls')
        btn_fok.config(text='确定',fg='black')
        setState(1,en_file,btn_browser,en_subject)
        show_lab['text']=''
        cmd_frame.grid_remove()
    else:
        if en_subject.get() and en_file.get():
            x=en_file.get().lower()
            y=en_subject.get()
            rx= re.match(r'.*\.xlsx?$',x)
            ry= re.match(r'(^\d{4})\/(\d{1,2}$)',y)
            if ry and 1<=int(ry.group(2))<=12:
                conf['subject']='%s年%02s月工资条'%(ry.group(1),ry.group(2))
                if rx:
                    if os.path.exists(x): 
                        try:
                            with xlrd.open_workbook(x) as bk:
                                pass
                        except Exception as e:
                            show_lab['text']='ErrorCode= %s\n该文件无法打开，请重新输入:'%e
                        else:
                            conf['fxls']=x
                            try:
                                Analysis(conf['fxls'])
                            except Exception as e:
                                show_lab['text']='Error: {}'.format(e)
                            else:
                                msg_frame.grid_remove()
                                cmd_frame.grid()
                                msg_frame.grid()
                                show_lab.config(wraplength=400)
                                btn_fok.config(text='修改',fg='blue')
                                setState(0,en_file,btn_browser,en_subject)
                                show_lab['text']='文件已分析完毕，请选择“预览”或“发送”'
                    else:
                        show_lab['text']='文件不存在！请重新输入:'
                else:
                    show_lab['text']="仅支持excel文件，请重新输入:" 
            else:
                show_lab['text']='年月格式不正确！\n'
        else:
            show_lab['text']='有项目未输入！\n'
        

def getsmtp():
    user=en_user.get().lower()
    pwd=en_pwd.get()
    r= re.match(r'(^[a-z]\w*[\-\_\.]?\w*)@(\w+\-?\w+\.[a-z]{2,6}$)',user)
    if r:
        smtp=r'smtp.'+r.group(2)
        s=TestSMTP(smtp,user,pwd)
        if s!=1:
            if s==2:
                en_user.delete(0,len(en_user.get()))
                x='%s\n服务器无法连接'%smtp
            else:
                x='%s\n账号无法登录'%user
            show_lab.config(text='%s\n %s，请重新设置！'%(show_lab['text'],x),fg='red')
            en_pwd.delete(0,len(en_pwd.get()))
            smtp,user,pwd='','',''
        else:
            conf['smtp'],conf['user'],conf['pwd']=smtp,user,pwd
            show_lab.config(text='OK！',fg='black')  
            en_user['state']=en_pwd['state']='readonly'
            btn_smtp.config(text='修改',command=chsmtp,fg='blue')
            WriteCfg(smtp,user,pwd)
            if tkinter.messagebox.askokcancel(title='登录成功',message='设置完毕。\n跳转到文件设置？'): ToFile()
                
                
    else:
        show_lab.config(text='%s \n账号格式不正确，请重新输入！'%user,fg='red')
        en_user.select_range(0,len(en_user.get()))
    
    
def chsmtp():
    en_user['state']=en_pwd['state']='normal'
    en_user.delete(0,len(en_user.get()))
    en_pwd.delete(0,len(en_pwd.get()))
    btn_smtp.config(text='登录测试',fg='black',command=getsmtp)
    show_lab['text']=''

def sureSend():
    if conf.get('smtp'):
        strmsg='邮件标题：%s\n发件人邮箱：%s\n工资表文件：%s\n'\
                %(conf['subject'],conf['user'],conf['fxls'])
        s=tkinter.messagebox.askokcancel(title="发送确认？",message=strmsg)
        if s==True:
            sendMail()
    else:
        s=tkinter.messagebox.showerror(title='错误',message='发件人邮箱账号未设置，请设置。')
        if s=='ok':
            ToConf()                                    


def viewFile():
    To_do(th_html,td_data,0)
    
def sendMail():
    To_do(th_html,td_data,1)
    
# ------------------------code--------------------------------------------------------#
def _format_addr(s): #格式化邮件信息
    name, addr = parseaddr(s)
    return formataddr((Header(name,'utf-8').encode(),addr))


    
class Sender(threading.Thread): #发送邮件--线程类对象
    def __init__(self):
        super(Sender,self).__init__()
    def run(self):
        global conf,q,errAccount
        with smtplib.SMTP(conf['smtp']) as s: # SMTP协议默认端口是25
            s.login(conf['user'], conf['pwd'])
            while True:
                msg=q.get()                
                try:# 检测邮件发送是否OK?
                    s.sendmail(conf['user'], msg['mail'], msg['msg'].as_string())
                except Exception as e:
                    lock.acquire()
                    errAccount.append(msg['mail'])
                    lock.release()
                else:
                    lock.acquire()
                    #print('%-2s%-30s%-25s%s'%("√",msg['mail'],time.time(),self.name))
                    lock.release()
                q.task_done() #告诉队列取数后的操作已完毕。

def Msg_encode(conf,th_html,td):
    html=html_head + "<table>"+ th_html + td['html'] + "</table>" + html_end
    msg = MIMEText(html, 'html', 'utf-8')
    msg['From'] = _format_addr('财务 <%s>' % conf['user'])
    msg['To'] = _format_addr('%s <%s>' %(td['mail'],td['name']))
    msg['Subject'] = Header(conf['subject'], 'utf-8').encode()
    return {'mail':td['mail'],'msg':msg}




def Analysis_Index(bk):
    '''
       分析--查找标题起始行号、列号、标题栏占用行数
           --查找“邮箱”列号
    '''
    '''
       @i['mail']  “邮箱”列号 
       @i['col']   标题起始列号
       @i['row']   标题起始行号
       @i['merge']   标题栏占用行数
       @i['title']   标题栏出现的首张Sheet
       
    '''

    i={}
    i['mail'],i['col'],i['row'],i['merge'],i['title']=None,None,None,0,0
    

    sheets=bk.sheets()
    s_mail=s_xuhao=None
    
    for sh in sheets:
        lrows=list(range(sh.nrows))
        ncols=sh.ncols
        name=sh.name
        i['title']+=1
        #print(name)
        for r in lrows:
            row_data=sh.row_values(r)
            #print( '###',r)
            if not i['mail']:                     
                sx=[k for k,v in enumerate(row_data) if isinstance(v,str) and re.match(r'(^\w*[\-\_\.]?\w*)@(\w+\-?\w+\.[a-z]{2,6}$)',v.lower())]
                if sx:
                    i['mail']=sx[0]
                    i['title']-=1
                    #print('r=',r,'title=',i['title'],'mail=',chr(65+i['mail']),'col=',chr(65+i['col']),'row=',i['row'],'merge=',i['merge'])
                    lrows.clear()
                    sheets.clear()
                sy=[k for k,v in enumerate(row_data) if isinstance(v,str) and re.match(r'\s*序\s*号\s*',v)]                    
                if sy:
                    i['col']=sy[0]
                    i['row']=r
                    j=0
                    col_data=sh.col_values(i['col'],i['row']+1)   #,ncols)
                    #print('data=',col_data)
                    for v in col_data:
                        if v=='':
                            j+=1
                        else:
                            break
                    #print('i=',i)
                    i['merge']=j
                    #print('r=',r,'mail=',i['mail'],'col=',i['col'],'row=',i['row'],'merge=',i['merge'])
##    print(i)
    return i

def Analysis_Title(sh,i):
    '''
       生成标题栏html代码
       
       ##仅支持标题栏为1-2行
       ##请保证一张sheet里只有一个邮箱列，一个姓名列，其它列均为数值类型

       @th_html    标题栏html代码
       
    '''
    row1_data=list(reversed(sh.row_values(i['row'])))
    j=1
    t1={} #第1行数据占行列数
    t2={} #第2行数据占行列数
    for v in row1_data:
        if v=='':
            j+=1
        else:
            t1[v]=[1,j]
            j=1
    if i['merge']:
        row2_data=reversed(sh.row_values(i['row']+1))
        for k,v in enumerate(row2_data):
            if v=='':
                t1[row1_data[k]][0]+=1
            else:
                t2[v]=[1,1]

                
# -- 标题栏html代码 --
    if t1:
        th_html="<tr>"
        for k,v in enumerate(sh.row_values(i['row'])):
            if v and k not in (i['col'],i['mail']):
                th_html+="<th rowspan=%s colspan=%s>%s</th>"%(*t1[v],v)
        th_html+="</tr>"
        if t2:
            th_html+="<tr>"
            for v in sh.row_values(i['row']+1):
                if v:
                    th_html+="<th rowspan=%s colspan=%s>%s</th>"%(*t2[v],v)
            th_html+="</tr>"
    #print(th_html)
    return th_html
                
def Analysis_Data(bk,i):
    '''
      生成数据list[dict{},……]   格式如： [{'dept':'xxx部',account':'jansen_yung@126.com','name':'黑白印象','td': td_html},{……}]
      @td_data 工资条数据   
    '''
    td_data=[]
    i_name=0
    sheets=bk.sheets()
    for sh in sheets:
        if i['mail']<=sh.ncols:
##                print(help(sh))
            col_data=sh.col_values(i['mail'],i['row']+i['title'])
            s= [re.match(r'(^\w*[\-\_\.]?\w*)@(\w+\-?\w+\.[a-z]{2,6}$)',v.lower()) for v in col_data if isinstance(v,str)]
##                print('s=',any(s),sh.name,s)
            if not any(s):
##                    print('name=',sh.name)
                continue
            else:
##                    print('####',sh.name)
                
##                    print('tg',td)
                m=[]
                for r in range(i['row']+i['title'],sh.nrows):
                    row_data=sh.row_values(r)
                    if isinstance(row_data[i['mail']],str) and re.match(r'(^\w*[\-\_\.]?\w*)@(\w+\-?\w+\.[a-z]{2,6}$)',row_data[i['mail']].lower()):
                        if not i_name:
                            for k,v in enumerate(row_data):
                                if isinstance(v,str) and k!=i['mail']:
                                    i_name=k
                                    break   
                        html='<tr>'
                        for k,v in enumerate(row_data):
                            if k not in (i['mail'],i['col']):
                                x=v if isinstance(v,str) else '%.2f'%v #工资条数值保留2位小数
                                #print(type(v),x)
                                html+="<td>%s</td>"%x
                        html+="</tr>"
                        m.extend([{'dept':sh.name,'mail':row_data[i['mail']],'name':row_data[i_name],'html':html}])
                        html=""
                td_data.extend(m)
                                   
        else:
            continue
    #print('td:',td_data)

    return td_data

def View_Html(th_html,td_data):
    '''
      预览文件按原表格的sheet顺序排列
 
    '''
    global html_head,html_end
    
    html_fname="payroll.html"
    html_content=''
    dept=""
    ix=0
    for td in td_data:
        if not td['dept']==dept:
            dept=td['dept']
            if ix!=0:
                html_content+='</td></tr></table>'
            html_content+='<br/><table class="box"><caption class="cap"><span>&nbsp;%s&nbsp;</span></caption>'%dept            
        html_content+='<table><caption class="msg">%s<input type="checkbox" style="vertical-align:middle;" ></caption>'\
                       %(td['dept']+'-'+td['name']+':'+td['mail'])+th_html
##      print('td:',v['html'],'name:',v['name'],'mail:',v['mail'])
        html_content+=td['html']
        html_content+="</table>"
        if dept:
            ix=1
             
    html=html_head + html_content + html_end
##    fpath=os.getcwd()+"\\"+html_fname
    try:
        with open(html_fname,'w') as f:
            f.write(html)
        webbrowser.open(html_fname)
    except Exception as e:
        show_lab['text']='预览文件无法生成，请检查是否对该文件夹(%s)有写权限！'%(os.getcwd()+"\\"+html_fname)

        

               
def SetGlobal():
    global conf,q,lock,errAccount,time_begin,th_html,td_data,html_head,html_end
    html_head='''
<html>
<head>
<meta charset="GBK">
<style type="text/css">
#mainbox {margin:5 auto;}
table {border-collapse:collapse;width:88%;margin:0 auto;}
table,tr,th,td {
    border:1px solid #000;
    text-align:center;
    font-size:12px;
    }
th {
    background-color:#ccc;
    color:Black;
    }
.msg {text-align:right;}
.box {border:0px solid blue;}
.cap {
    color: #aa2211;
    font-size:2em;
    font-weight:bold;
    }
.cap span {
    background-color:OrangeRed;
    color:White;
    }

</style>
</head>
<body>
<div id="mainbox">
'''
    html_end='''
</div>
</body>
</html>
'''
    conf={}
    q=queue.Queue()
    lock=threading.Lock()
    errAccount=[]
    
def Analysis(fname,mode=0):
    global th_html,td_data
    with xlrd.open_workbook(fname) as bk:
        i=Analysis_Index(bk)
        try:
            sh=bk.sheet_by_index(i['title'])
            th_html=Analysis_Title(sh,i)
            td_data=Analysis_Data(bk,i)
        except Exception as e:
            string='''
表格需要符合以下要求：

标题栏:
1. A列标题必须为“序号”
2. 支持占用1-2行
3. 在每一列上不能出现空数据的单元格（可以有合并单元格）

工资行:
1. 有且仅有两列文本数据（邮箱列，姓名列）
2. 其它列数据必须是数值型

'''
            raise Exception(string)
        
def To_do(th_html,td_data,mode=0,thread_num=4):    
    if mode==0:           
        View_Html(th_html,td_data)
    if mode==1:
        show_lab['text']='正在发送……'
        # --工资条邮件生成 --
        for td in td_data:
            data=Msg_encode(conf,th_html,td)
            q.put(data)
        # -- 工资条邮件多线程发送 --
        Consumer=[Sender() for i in range(thread_num)]
        time_begin=time.time() #计时开始
        for c in Consumer:
            c.daemon = True
            c.start()
        q.join()

        # -- 发送完毕，输出结果 --
        strx="\n发送完毕!  总计用时：%4.3f秒"%(time.time()-time_begin)
        strx+="\n以下邮件账号发送失败：\n"
        for x in errAccount:
            strx+="%-2s%-25s\n"%("×",x)
        show_lab['text']+=strx
        setState(0,btn_send)
            

    
       
def TestSMTP(smtp,user,pwd): #返回结果 1-OK,2-服务器连接失败，3-账号登录失败
    try:
        with smtplib.SMTP(smtp) as s:
            try:
                s.login(user,pwd)
                return 1
            except Exception as e:
                show_lab['text']="[Error]……authentication failed!  Please reset!\n"
                return 3
    except Exception as e:
        show_lab['text']="[Error]……Server can't connect!  Please reset!\n"
        return 2
        

def X_64code(string,mode=0): #0 为解码，其他为编码
    if not mode: #0 解码
        code = base64.decodestring(string.encode('ascii')).decode('ascii')
    else:  #其它  编码
        code = base64.encodebytes(string.encode('ascii')).decode('ascii')
    return code
    
def ReadCfg(fname="payConfig.ini"): #获取或设置相关配置信息
    cf = configparser.ConfigParser()
    if os.path.exists(fname):#确认配置文件存在，则读取配置
        cf.read(fname) 
        user=X_64code(cf.get('smtpset','user'))
        pwd=X_64code(cf.get('smtpset','pwd'))
        smtp=X_64code(cf.get('smtpset','smtp'))
        if not TestSMTP(smtp,user,pwd):
            ToConf()
        else:
            conf.update(zip(['smtp','user','pwd'],[smtp,user,pwd]))
    else:
        ToConf()
    

def WriteCfg(smtp,user,pwd,fname="payConfig.ini"):  
    cf = configparser.ConfigParser()
    cf.add_section("smtpset")#增加section 
    cf.set("smtpset", "user", X_64code(user,1))#增加option 
    cf.set("smtpset", "pwd", X_64code(pwd,1))
    cf.set("smtpset", "smtp", X_64code(smtp,1))
    try:
        with open(fname, "w") as f: 
            cf.write(f)#写入配置文件文件中
    except Exception as e:
       show_lab['text']+='\n[msg]……配置文件保存失败！'



if __name__=='__main__':
    main()

