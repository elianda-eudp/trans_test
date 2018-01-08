# -*- coding: utf-8 -*-
"""
Created on Mon Oct 16 18:01:06 2017

@author: Administrator
"""
import xlsxwriter

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
workbook = xlsxwriter.Workbook(u'report.xlsx')
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


    data = {u"test_name": u"智商", u"test_version": u"v2.0.8", u"test_pl": u"android", u"test_net": u"wifi"}
    _write_center(worksheet, u"C3", data[u'test_name'], workbook)
    _write_center(worksheet, u"C4", data[u'test_version'], workbook)
    _write_center(worksheet, u"C5", data[u'test_pl'], workbook)
    _write_center(worksheet, u"C6", data[u'test_net'], workbook)

    _write_center(worksheet, u"D3", u"接口总数", workbook)
    _write_center(worksheet, u"D4", u"通过总数", workbook)
    _write_center(worksheet, u"D5", u"失败总数", workbook)
    _write_center(worksheet, u"D6", u"测试日期", workbook)



    data1 = {u"test_sum": 100, u"test_success": 80, u"test_failed": 20, u"test_date": u"2018-10-10 12:10"}
    _write_center(worksheet, u"E3", data1[u'test_sum'], workbook)
    _write_center(worksheet, u"E4", data1[u'test_success'], workbook)
    _write_center(worksheet, u"E5", data1[u'test_failed'], workbook)
    _write_center(worksheet, u"E6", data1[u'test_date'], workbook)

    _write_center(worksheet, u"F3", u"分数", workbook)


    worksheet.merge_range(u'F4:F6', u'60', get_format_center(workbook))

    pie(workbook, worksheet)

 # 生成饼形图
def pie(workbook, worksheet):
    chart1 = workbook.add_chart({u'type': u'pie'})
    chart1.add_series({
    u'name':       u'接口测试统计',
    u'categories': u'=测试总况!$D$4:$D$5',
    u'values':    u'=测试总况!$E$4:$E$5',
    })
    chart1.set_title({u'name': u'接口测试统计'})
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
    _write_center(worksheet, u"B2", u'接口名称', workbook)
    _write_center(worksheet, u"C2", u'接口协议', workbook)
    _write_center(worksheet, u"D2", u'URL', workbook)
    _write_center(worksheet, u"E2", u'参数', workbook)
    _write_center(worksheet, u"F2", u'预期值', workbook)
    _write_center(worksheet, u"G2", u'实际值', workbook)
    _write_center(worksheet, u"H2", u'测试结果', workbook)

    data = {u"info": [{u"t_id": u"1001", u"t_name": u"登陆", u"t_method": u"post", u"t_url": u"http://XXX?login", u"t_param": u"{user_name:test,pwd:111111}",
                      u"t_hope": u"{code:1,msg:登陆成功}", u"t_actual": u"{code:0,msg:密码错误}", u"t_result": u"失败"}, {u"t_id": u"1002", u"t_name": u"商品列表", u"t_method": u"get", u"t_url": u"http://XXX?getFoodList", u"t_param": u"{}",
                      u"t_hope": u"{code:1,msg:成功,info:[{name:123,detal:dfadfa,img:product/1.png},{name:456,detal:dfadfa,img:product/1.png}]}", u"t_actual": u"{code:1,msg:成功,info:[{name:123,detal:dfadfa,img:product/1.png},{name:456,detal:dfadfa,img:product/1.png}]}", u"t_result": u"成功"}],
            u"test_sum": 100,u"test_success": 20, u"test_failed": 80}
    temp = 4
    for item in data[u"info"]:
        _write_center(worksheet, u"A"+str(temp), item[u"t_id"], workbook)
        _write_center(worksheet, u"B"+str(temp), item[u"t_name"], workbook)
        _write_center(worksheet, u"C"+str(temp), item[u"t_method"], workbook)
        _write_center(worksheet, u"D"+str(temp), item[u"t_url"], workbook)
        _write_center(worksheet, u"E"+str(temp), item[u"t_param"], workbook)
        _write_center(worksheet, u"F"+str(temp), item[u"t_hope"], workbook)
        _write_center(worksheet, u"G"+str(temp), item[u"t_actual"], workbook)
        _write_center(worksheet, u"H"+str(temp), item[u"t_result"], workbook)
        temp = temp -1

init(worksheet)
test_detail(worksheet2)

workbook.close()