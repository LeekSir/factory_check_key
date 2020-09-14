import sys
import xlrd
import xlwt
import xlutils.copy
import tkinter.messagebox
import tkinter as tk
import serial.tools.list_ports
#from serial import Serial
import os
import time
import logging
import logging.handlers
import datetime
from tkinter import *
import tkinter.filedialog
from  tkinter  import ttk
import serial.tools.list_ports  

 
window = Tk()
window.title('8710_MAC查重工具')

screenwidth = window.winfo_screenwidth()  
screenheight = window.winfo_screenheight()  
size = '%dx%d+%d+%d' % (850, 750, (screenwidth - 1000)/2, (screenheight - 700)/2)  
print(size) 
window.geometry(size) 

#函数
def find_new_file(dir):
    '''查找目录下最新的文件'''
    file_lists = os.listdir(dir)
    file_lists.sort(key=lambda fn: os.path.getmtime(dir + "\\" + fn)
                    if not os.path.isdir(dir + "\\" + fn) else 0)
    print('最新的文件为： ' + file_lists[-1])
    file = os.path.join(dir, file_lists[-1])
    print('完整路径：', file)
    return file

def go_com(*args):   #处理事件，*args表示可变参数，返回COM口
	return comboxlist1.get() #返回选中的值
	

def go_baud_rate(*args):   #处理事件，*args表示可变参数，返回波特率
	return comboxlist2.get() #打印选中的值

	
def __change_color():
	msg.config(bg='gray', font=('times', 40, 'italic'))
	b1['text'] = '下 一 次 测 试'
# 定义按键两个触发事件时的函数"开始测试"和"退出"

def func(event):
	start_test()
	print("You hit return.")


def start_test(*args): 
	try:
		#端口，GNU / Linux上的/ dev / ttyUSB0 等 或 Windows上的 COM3 等
		portx=go_com()
		#波特率，标准值之一：50,75,110,134,150,200,300,600,1200,1800,2400,4800,9600,19200,38400,57600,115200
		bps=go_baud_rate()
		logger.info("Port: " + portx + ", Baud: " + bps + "bps")
		print("Port: " + portx + ", Baud: " + bps + "bps")
		#超时设置,None：永远等待操作，0为立即返回请求结果，其他值为等待超时时间(单位为秒）
		timex=5
		# 打开串口，并得到串口对象
		ser=serial.Serial(portx,bps,timeout=timex)

		# 写数据
		result=ser.write("AT+TXMAC=0\r\n".encode("gbk"))

		logger.info("写总字节数: " + str(result))
		print("写总字节数: " + str(result))
		
		#循环接收数据，此为死循环，可用线程实现
		i=True
		""" 读取数据 """
		out = ''
		while True:
			while ser.inWaiting() > 0:
				out += ser.read(1).decode()  # 一个一个的读取
			if len(out) >= 21:
				dev_name = out[9:21].upper()
				logger.info("dev_name: " + dev_name)
				print("dev_name: " + dev_name)
				break
		
		ser.close()#关闭串口
		#time.sleep(2)
	except Exception as e:
		logger.error('The port open error!!!')
		tkinter.messagebox.showerror('错误','串口异常！')
		print("---异常---：",e)

	
	
	try:
		#打开文件，获取excel文件的workbook（工作簿）对象
		


		
		excel = xlrd.open_workbook(var_xls.get(),encoding_override="utf-8")
		wt = xlutils.copy.copy(excel)
		#获取sheet对象
		print ("hello")
		sheet = excel.sheets()[0]
		sheet1 = wt.get_sheet(0)
		
		sheet_row_mount = sheet.nrows#4行
		sheet_col_mount = sheet.ncols#6列
		
		#excel = xlwt.open_workbook(excel_path,encoding_override="utf-8")
		cmp_flag = False
		dev_name1 = 'T' + dev_name
		j=0
		for x in range(sheet_row_mount):#4
			y = 0
			while y < sheet_col_mount:#6
				if dev_name1 == sheet.cell_value(x,y) and (not cmp_flag):
					j = j + 1
					logger.info("KEY_重复写入")
					print("KEY_重复写入")
					break
				y += 1
		x=0
		y=0
		i=0
		for x in range(sheet_row_mount):#4
			y = 0
			while y < sheet_col_mount:#6
				if dev_name == sheet.cell_value(x,y) and (not cmp_flag):
					i = i + 1
					sheet1.write(x, y, dev_name1)
					wt.save(var_xls.get())
					print(i)
					break
				y += 1
		
		if i == 1:
			logger.info("KEY_存在，且为第一次烧录")
			print("KEY_存在，且为第一次烧录")
			cmp_flag = True
		if i+j == 0:
			logger.info("KEY_在系统中不存在")
			print("KEY_在系统中不存在")
		
		#b1['text'] = '测试结束'
		if(cmp_flag ):
			logger.info("测试成功,KEY_不重复")
			print("测试成功,KEY_不重复")
			var.set("测试已PASS！")
			msg.config(bg='lightgreen', font=('times', 40, 'italic'))
		else:
			if(i+j > 0):
				logger.info("测试失败,KEY_重复")
				print("测试失败,KEY_重复")
				var.set("测试Fail，重号")
				msg.config(bg='red', font=('times', 40, 'italic'))
			else:
				logger.info("测试失败,KEY_在MAC表中不存在")
				print("测试失败,KEY_在MAC表中不存在")
				var.set("测试Fail，空片")
				msg.config(bg='red', font=('times', 40, 'italic'))
		window.after(3000, __change_color)	
		b1['text'] = '测 试 结 束'
	#查KEY是否重复
	except Exception as e:
		logger.error('The excel open error!!!')
		tkinter.messagebox.showerror('错误','表格打开异常！')
		print("---异常---：",e)
	
	
def clear_test():   # 在文本框内容最后接着插入输入内容
	var.set("测试信息显示")
	msg.config(bg='white', font=('times', 40, 'italic'))
	"""
	if("测试PASS!!!!!!" == var.get()):
		var.set("下一次测试!")
		msg.config(bg='yellow', font=('times', 30, 'italic'))
	if("测试ERR!!!!!!" == var.get()):
		var.set("请重新测试!")
		msg.config(bg='blue', font=('times', 30, 'italic'))
	"""



#主程序
i =0
while i < 4:
	Label(window, text="").grid(row=i,column=10,padx=5, pady=5,sticky="n" + "s" + "w" + "e" )
	i = i + 1

i =0
while i < 10:
	Label(window, text="").grid(row=0,column=i,padx=5, pady=5,sticky="n" + "s" + "w" + "e" )
	i = i + 1

Label(window, text="端口号").grid(row=5,column=14,padx=5, pady=5)
Label(window, text="波特率").grid(row=7,column=14,padx=5, pady=5)


logger = logging.getLogger('mylogger')

   
port_list = list(serial.tools.list_ports.comports())  

#端口号
comvalue1=tk.StringVar()#窗体自带的文本，新建一个值
comboxlist1=ttk.Combobox(window,textvariable=comvalue1) #初始化
port = ""
if len(port_list) <= 0:  
	print ("The Serial port can't find!")  
	logger.error("The Serial port can't find!")
	tkinter.messagebox.showwarning('警告','无串口！！！')
else: 
	for i in list(port_list):
		port += str(i).split(' ')[0] + ' '
		print(port)
		#comboxlist1.set(port)
comboxlist1["values"] = str(port).split(' ')
#comboxlist1["values"]=("COM1","COM2","COM3","COM4","COM5","COM6","COM7","COM8","COM9","COM10","COM11","COM18")
comboxlist1.current(0)  #选择第一个
comboxlist1.bind("<<ComboboxSelected>>",go_com)  #绑定事件,(下拉列表框被选中时，绑定go()函数)
comboxlist1.grid(row=5, column=15, sticky=W+E+N+S)
#波特率
comvalue2=tk.StringVar()#窗体自带的文本，新建一个值
comboxlist2=ttk.Combobox(window,textvariable=comvalue2) #初始化
comboxlist2["values"]=("9600","19200","38400","57600","115200","921600")
comboxlist2.current(0)  #选择第一个
comboxlist2.bind("<<ComboboxSelected>>",go_baud_rate)  #绑定事件,(下拉列表框被选中时，绑定go()函数)
comboxlist2.grid(row=7, column=15, sticky=W+E+N+S)

#创建开始测试按键和退出按键并分别触发两种情况
b1 = tk.Button(window, text='开 始 测 试', width=15,
               height=3, font=('microsoft yahei', '15'),command=lambda : start_test(""))
#b1.side(BOTTOM)
Label(window, text="").grid(row=10,column=10,padx=5, pady=5,sticky="n" + "s" + "w" + "e" )
b1.grid(row=11, column=15)


#测试结果显示
Label(window, text="").grid(row=12,column=10,padx=5, pady=5,sticky="n" + "s" + "w" + "e" )
var = StringVar()
msg = Message(window, textvariable=var, width=500)
var.set("测试信息显示")

msg.config(bg='white', font=('times', 40, 'italic'))
msg.grid(row=13, column=15)

Label(window, text="").grid(row=14,column=10,padx=5, pady=5,sticky="n" + "s" + "w" + "e" )
b2 = tk.Button(window, text='清 除 信 息', width=15,
               height=3, font=('microsoft yahei', '15'), command=clear_test)
b2.grid(row=15, column=15)

#回车按键绑定开始函数
window.bind('<Return>', func)
#window.bind('<space>', func)


logger.setLevel(logging.INFO)
rf_handler = logging.handlers.TimedRotatingFileHandler('.\\log\\all.log', when='midnight', interval=1, backupCount=7, atTime=datetime.time(0, 0, 0, 0))
rf_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))


logger.addHandler(rf_handler)
#logger.addHandler(f_handler)
#logger.info(sys.stdout)
#sys.stdout = logger("E:\\产测工具\\all.log")
#filename = ""

def xz():
    filename = tkinter.filedialog.askopenfilename()
    print(filename)
    if filename != '':
        var_xls.set(filename.split('/')[-1])
    else:
        var_xls.set("您没有选择任何文件!")
"""
def xz():
	#import os
	filelist=[]

	for root, dirs, files in os.walk(".", topdown=False):
		for name in files:
			str=os.path.join(root, name)
			if str.split('.')[-1]=='xls':
				filelist.append(str)
	var_xls.set(filelist[0])
"""
#b2 = tk.Button(window, text='清除信息', width=10,
#               height=2, command=clear_test)
#b2.grid(row=5, column=1)
var_xls = StringVar()
lb = Message(window, textvariable=var_xls, width=300)
var_xls.set("请选择key文件")
lb.config(bg='white', font=('times', 18, 'italic'))
lb.grid(row=9, column=15)
lb.config(bg='yellow', font=('times', 24, 'italic'))

#自动选择key文件

filelist2=[]
"""
for root, dirs, files in os.walk(".", topdown=False):
	for name in files:
		str=os.path.join(root, name)
		if str.split('.')[-1]=='xls':
			filelist2.append(str)
			#var_xls.set(str[2::])
var_xls.set(filelist2[0])
"""
#手动选择key文件
btn = Button(window,text="选择“key”文件", width=15,
               height=1, font=('microsoft yahei', '15'), command=xz)
btn.grid(row=9, column=16)

#window.after(2000,change_color,0)
mainloop()
 
 

	

