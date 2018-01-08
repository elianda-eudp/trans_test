# -*- coding: utf-8 -*-
"""
Created on Fri Sep 22 13:48:55 2017

@author: Administrator
"""
import  paramiko
import  os
import  chardet
import fabric
from fabric.api import hosts, env,run,execute,task
from fabric.api import *
from fabric.colors import *
from fabric.context_managers import *
import  xdrlib 
import xlrd
import xlwt
import xlsxwriter
import datetime
#import sys
#reload(sys)
#sys.setdefaultencoding('utf8')

filename = u'E:\邮储理财项目组工作\新技术\Python\pythonsrc\\tran_test\测试用例.xlsx'

all_dict={}

def open_excel(file= 'file.xlsx'):
	try:
         print file
         data = xlrd.open_workbook(file)
         return data
	except Exception,e:
         print str(e)
		
		
#根据索引获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_index：表的索引
def excel_table_byindex(file= filename,colnameindex=0,by_index=0):
    #data = open_excel(file)
    #print filename
    data = xlrd.open_workbook(file)
    table = data.sheets()[by_index]
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    print nrows,ncols
    headlist=[]
 
    for colnum in range(2,ncols):
        #print colnum
        #print chardet.detect(table.row(0)[colnum].value)
        #print type(table.row(0)[colnum].value)
        low_cell = str(table.row(0)[colnum].value.encode('UTF-8'))
        if low_cell=='_tx_code' :
            table.row(1)[colnum].ctype=1
            trade_code=str(table.row(1)[colnum].value)
            print table.cell_value(1,colnum).encode('UTF-8')
        headlist.append(low_cell)
    rowslist=[]     
    for rownum in range(1,nrows):
        rowlist=[]
        for colnum in range(2,ncols):
            #print colnum
            table.row(rownum)[colnum].ctype=1
            low_cell = str(table.row(rownum)[colnum].value)	
            rowlist.append(low_cell)
        rowslist.append(rowlist)
        table.row(rownum)[0].ctype=1
        table.row(rownum)[1].ctype=1
        str1=str(table.row(rownum)[0].value.encode('utf-8'))
        str2=str(table.row(rownum)[1].value.encode('utf-8'))
        print str1
        print chardet.detect(str1)
        print str2
        all_dict[str(str1)]=str2
        #写ini文件文件
        file_name=u'E:\邮储理财项目组工作\新技术\Python\pythonsrc\\tran_test\ini\\' + trade_code + '_' + table.row(rownum)[0].value + '.ini'
        print trade_code.encode('UTF-8')
        print file_name.encode('UTF-8')
        f = open(file_name,'w')
        f.write( u'[filed]\n')
        for head_filed in headlist:
            #print chardet.detect(head_filed)
            tmp_str=head_filed + '\n'
            f.write(tmp_str)
        f.write( u'[filed_end]\n')
        f.write( u'[values]\n')
        for body_list in rowslist:
            body_filed = '|'.join(body_list)
            if body_list[-1]=='':
                f.write(body_filed + u'|\n')
            else:
                f.write(body_filed + u'\n')
        f.write( u'[values_end]\n')
        f.close()

        

        


 
def main():
	excel_table_byindex()

if __name__=="__main__":
	main()

t = paramiko.Transport(("10.136.1.3",22))
t.connect(username = "financialsys", password ="financialsys")
sftp = paramiko.SFTPClient.from_transport(t)  #使用t的设置方式连接远程主机
#remotepath='/app/financialsys/tb_d_ifperiod.dat'
localpath=u'E:\邮储理财项目组工作\新技术\Python\pythonsrc\\tran_test\ini'  
'''因为\t为特殊字符，所有需要转义'''
remote_dir = u'/home/financialsys/public/yy'
print localpath.encode('UTF-8')
files=os.listdir(localpath) 
#sftp.get(remotepath, localpath)  #下载文件
for f in files:
    localfile=localpath+'\\'+f
    remotefile=remote_dir+'/'+f
    print localfile.encode('UTF-8')
    print remotefile.encode('UTF-8')
    #sftp.get(remotefile,localfile) 
    sftp.put(localfile,remotefile) #上传文件
t.close()


env.user="financialsys"
env.hosts=[u"10.136.1.3"]
env.password="financialsys"
env.warn_only = True


#@hosts(['10.136.1.3'])
@task
def cmd():
    with cd('/home/financialsys/bin/'):
        return run('python monitor.py '+remotefile )

match_dict={}
no_match_dict={}
for f in files:
    remotefile=str(remote_dir+'/'+f).encode('UTF-8')
    results = execute(cmd )
    #tran_out_str=str(results[u'10.136.1.3']).decode('GB2312').encode('UTF-8')    
    #print tran_out_str
    tran_out_str= results[u'10.136.1.3'].split('\r\n')
    #print tran_out_str
    #print all_dict["text:u'1'"]
    #print all_dict[tran_out_str[1].decode('GB2312').encode('UTF-8')]   #预期结果
    #print type(tran_out_str)
    #print tran_out_str

    dict_index = tran_out_str[1].decode('GB2312').encode('utf-8').split('|')[0]
    #dict_index = tran_out_str[1].decode('ascii').encode('UTF-8')
    dict_value = tran_out_str[1].decode('GB2312').encode('utf-8').split('|')[1]
    
    #print chardet.detect(tran_out_str[1])
    print chardet.detect(dict_index)
    #print chardet.detect(dict_value)
    print str(dict_index[6:])
    if all_dict[str(dict_index[6:])] == dict_value:
        match_dict[str(dict_index)]=dict_value
    else:
        no_match_dict[str(dict_index)]=dict_value

#print   len(match_dict),len(no_match_dict)                           
fabric.network.disconnect_all()


        
def get_format(wd, option={}):
    return wd.add_format(option)

# 设置居中
def get_format_center(wd,num=1):
    return wd.add_format({u'align': u'center',u'valign': u'vcenter',u'border':num})
def set_border_(wd, num=1):
    return wd.add_format({}).set_border(num)

#写数据
def _write_center(worksheet, cl, data, wd):
    return worksheet.write(cl, data, get_format_center(wd))
workbook = xlsxwriter.Workbook(u"测试报告.xlsx")
worksheet = workbook.add_worksheet(u"测试总况")
worksheet2 = workbook.add_worksheet(u"测试详情")

def init(worksheet):
    # 设置列行的宽高
    worksheet.set_column(u"A:A", 15)
    worksheet.set_column(u"B:B", 20)
    worksheet.set_column(u"C:C", 20)
    worksheet.set_column(u"D:D", 20)
    worksheet.set_column(u"E:E", 20)
    worksheet.set_column(u"F:F", 20)

    worksheet.set_row(1, 30)
    worksheet.set_row(2, 30)
    worksheet.set_row(3, 30)
    worksheet.set_row(4, 30)
    worksheet.set_row(5, 30)

    # worksheet.set_row(0, 200)

    define_format_H1 = get_format(workbook, {u'bold': True, u'font_size': 18})
    define_format_H2 = get_format(workbook, {u'bold': True, u'font_size': 14})
    define_format_H1.set_border(1)

    define_format_H2.set_border(1)
    define_format_H1.set_align(u"center")
    define_format_H2.set_align(u"center")
    define_format_H2.set_bg_color(u"blue")
    define_format_H2.set_color(u"#ffffff")
    # Create a new Chart object.

    worksheet.merge_range(u'A1:F1', u'测试报告总概况', define_format_H1)
    worksheet.merge_range(u'A2:F2', u'测试概括', define_format_H2)
    worksheet.merge_range(u'A3:A6', u'这里放图片', get_format_center(workbook))

    _write_center(worksheet, u"B3", u'项目名称', workbook)
    _write_center(worksheet, u"B4", u'接口版本', workbook)
    _write_center(worksheet, u"B5", u'脚本语言', workbook)
    _write_center(worksheet, u"B6", u'测试网络', workbook)


    data = {u"test_name": u"理财联机交易", u"test_version": u"v1.0.1", u"test_pl": u"python", u"test_net": u"合肥"}
    _write_center(worksheet, u"C3", data[u'test_name'], workbook)
    _write_center(worksheet, u"C4", data[u'test_version'], workbook)
    _write_center(worksheet, u"C5", data[u'test_pl'], workbook)
    _write_center(worksheet, u"C6", data[u'test_net'], workbook)

    _write_center(worksheet, u"D3", u"案例总数", workbook)
    _write_center(worksheet, u"D4", u"通过总数", workbook)
    _write_center(worksheet, u"D5", u"失败总数", workbook)
    _write_center(worksheet, u"D6", u"测试日期", workbook)


    now = datetime.datetime.now() 
    data1 = {u"test_sum": len(match_dict)+len(no_match_dict), u"test_success": len(match_dict), u"test_failed": len(no_match_dict), u"test_date":now.strftime('%Y-%m-%d %H:%M:%S') }
    _write_center(worksheet, u"E3", data1[u'test_sum'], workbook)
    _write_center(worksheet, u"E4", data1[u'test_success'], workbook)
    _write_center(worksheet, u"E5", data1[u'test_failed'], workbook)
    _write_center(worksheet, u"E6", data1[u'test_date'], workbook)

    _write_center(worksheet, u"F3", u"通过率", workbook)

    tongguolv=float(data1[u'test_failed'])/float(data1[u'test_sum'])
    worksheet.merge_range(u'F4:F6', tongguolv, get_format_center(workbook))

    pie(workbook, worksheet)

 # 生成饼形图
def pie(workbook, worksheet):
    chart1 = workbook.add_chart({u'type': u'pie'})
    chart1.add_series({
    u'name':       u'测试统计',
    u'categories': u'=测试总况!$D$4:$D$5',
    u'values':    u'=测试总况!$E$4:$E$5',
    })
    chart1.set_title({u'name': u'测试统计'})
    chart1.set_style(10)
    worksheet.insert_chart(u'A9', chart1, {u'x_offset': 25, u'y_offset': 10})

def test_detail(worksheet):

    # 设置列行的宽高
    worksheet.set_column(u"A:A", 30)
    worksheet.set_column(u"B:B", 20)
    worksheet.set_column(u"C:C", 20)
    worksheet.set_column(u"D:D", 20)
    worksheet.set_column(u"E:E", 20)
    worksheet.set_column(u"F:F", 20)
    worksheet.set_column(u"G:G", 20)
    worksheet.set_column(u"H:H", 20)

    worksheet.set_row(1, 30)
    worksheet.set_row(2, 30)
    worksheet.set_row(3, 30)
    worksheet.set_row(4, 30)
    worksheet.set_row(5, 30)
    worksheet.set_row(6, 30)
    worksheet.set_row(7, 30)



    worksheet.merge_range(u'A1:H1', u'测试详情', get_format(workbook, {u'bold': True, u'font_size': 18 ,u'align': u'center',u'valign': u'vcenter',u'bg_color': u'blue', u'font_color': u'#ffffff'}))
    _write_center(worksheet, u"A2", u'用例ID', workbook)
    _write_center(worksheet, u"B2", u'预期值', workbook)
    _write_center(worksheet, u"C2", u'实际值', workbook)
    _write_center(worksheet, u"D2", u'测试结果', workbook)

    
    for item in no_match_dict.iterkeys():
    #    print item
        #print chardet.detect(str(item).encode('utf-8'))
        #print chardet.detect(all_dict[item])
        #print chardet.detect(no_match_dict[item])
        print chardet.detect(item)
        print chardet.detect(all_dict[item[6:]])
        print chardet.detect(no_match_dict[item])
        #_write_center(worksheet, u"A" + item[6:], item, workbook)
        _write_center(worksheet, u"B" + item[6:], all_dict[item[6:]], workbook)
        #_write_center(worksheet, u"C" + item[6:], str(no_match_dict[item]).decode('utf-8').encode('gbk'), workbook)
        _write_center(worksheet, u"D" + item[6:], u'结果与预期不匹配', workbook)

    #for item in match_dict.iterkeys():
        #print item
        #print chardet.detect(item)
        #print chardet.detect(all_dict[item[6:]])
        #print chardet.detect(match_dict[item])
     #   _write_center(worksheet, u"A" + item[6:], item, workbook)
     #   _write_center(worksheet, u"B" + item[6:], all_dict[item[6:]], workbook)
     #   _write_center(worksheet, u"C" + item[6:], match_dict[item], workbook)
     #   _write_center(worksheet, u"D" + item[6:], u'结果与预期匹配', workbook)

init(worksheet)
test_detail(worksheet2)

workbook.close() 

