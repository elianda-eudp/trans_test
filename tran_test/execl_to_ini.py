# -*- coding: utf-8 -*-
"""
Created on Sat Sep 30 10:29:02 2017

@author: Administrator
"""
 
import xlrd
import  chardet


filename = u'E:\邮储理财项目组工作\新技术\Python\pythonsrc\\tran_test\测试用例.xlsx'
		
#根据索引获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_index：表的索引
def excel_table_byindex(file= filename,colnameindex=0,by_index=0):
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
            print table.cell_value(1,colnum)
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
    
        #写ini文件文件
        file_name=u'E:\邮储理财项目组工作\新技术\Python\pythonsrc\\tran_test\ini\\' + trade_code + '_' + table.row(rownum)[0].value + '.ini'
        print trade_code
        print file_name
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
        


if __name__=="__main__":
	excel_table_byindex()
