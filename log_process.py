#coding=utf-8

'''
Description: 烤机测试日志解析脚本
Anthor: James.Li
Date: 2017-01-05 13:19
Using: 
    python <日志路径>/ <输出路径>/<输出文件名>.xls
Require: 
    sudo apt-get install python-xlwt python-xlrd

COPYRIGHT © 2015 LIBRATONE 版权所有

'''

import os.path
import sys
import xlrd
import xlwt
import linecache
import string

'''
# 样式说明
# Text values for colour indices. "grey" is a synonym of "gray".
# The names are those given by Microsoft Excel 2003 to the colours
# in the default palette. There is no great correspondence with
# any W3C name-to-RGB mapping.
_colour_map_text = """\
aqua 0x31
black 0x08启动次数:  79

blue 0x0C
blue_gray 0x36
bright_green 0x0B
brown 0x3C
coral 0x1D
cyan_ega 0x0F
dark_blue 0x12
dark_blue_ega 0x12
dark_green 0x3A
dark_green_ega 0x11
dark_purple 0x1C
dark_red 0x10
dark_red_ega 0x10
dark_teal 0x38
dark_yellow 0x13
gold 0x33
gray_ega 0x17
gray25 0x16
gray40 0x37
gray50 0x17
gray80 0x3F
green 0x11
ice_blue 0x1F
indigo 0x3E
ivory 0x1A1226\&80/
lavender 0x2E
light_blue 0x30
light_green 0x2A
light_orange 0x34
light_turquoise 0x29
light_yellow 0x2B
lime 0x32
magenta_ega 0x0E
ocean_blue 0x1E
olive_ega 0x13
olive_green 0x3B
orange 0x35
pale_blue 0x2C
periwinkle 0x18
pink 0x0E
plum 0x3D
purple_ega 0x14
red 0x0A
rose 0x2D
sea_green 0x39
silver_ega 0x16
sky_blue 0x28
tan 0x2F
teal 0x15
teal_ega 0x15
turquoise 0x0F
violet 0x14
white 0x09
yellow 0x0D"""
'''

class LBRPyExcel:
    def __init__(self, path):
        # 保存路径信息
        self.path = path
        # 创建工作簿
        self.workbook = xlwt.Workbook(encoding='utf-8')
        # 创建工作表
        self.sheets = []
        self.sheet1 = self.workbook.add_sheet(u'概览', cell_overwrite_ok=True)
        self.sheets.append(self.sheet1)
        # 创建表头样式
        self.head_style = xlwt.XFStyle()
        self.head_style = xlwt.easyxf('pattern: pattern solid, fore_colour black; font: colour white, bold on;');
        # 设置居中对齐
        self.head_style.alignment.horz = xlwt.Alignment.HORZ_CENTER
        # 设置表头字体
        self.head_style.font.name = 'WenQuanYi Zen Hei'
        # 设置加粗
        self.head_style.font.bold = True
        
        # 创建内容样式
        self.content_style = xlwt.XFStyle()
        self.content_style.alignment.horz = xlwt.Alignment.HORZ_CENTER
        self.content_style.font.name = 'WenQuanYi Zen Hei'
        self.content_style.font.bold = False

        # 设置警告项背景颜色为红色, 字体加粗
        self.warning_style = xlwt.easyxf('pattern: pattern solid, fore_colour red; font: colour black, bold on;');
        self.warning_style.font.name = 'WenQuanYi Zen Hei'
        # 设置警告项左对齐
        self.warning_style.alignment.horz = xlwt.Alignment.HORZ_CENTER
        
        # 设置错误项背景颜色为红色, 字体加粗
        self.err_style = xlwt.easyxf('pattern: pattern solid, fore_colour red; font: colour black, bold on;');
        self.err_style.font.name = 'WenQuanYi Zen Hei'
        # 设置错误项左对齐
        self.err_style.alignment.horz = xlwt.Alignment.HORZ_LEFT
        
        # 设置错误项背景颜色为红色, 字体加粗
        self.pro_style = xlwt.easyxf('pattern: pattern solid, fore_colour yellow; font: colour black, bold on;');
        self.pro_style.font.name = 'WenQuanYi Zen Hei'
        # 设置错误项左对齐
        self.pro_style.alignment.horz = xlwt.Alignment.HORZ_CENTER        
        pass
    
    def prepare(self):
        # 依次写入表头
        self.f_index = 0
        self.sheet1.write_merge(0, 0, 0, 6, u'提示: 此表格由日志分析程序自动生成', self.pro_style)
        self.f_index += 1
        self.sheet1.write(self.f_index, 0, u'设备编号', self.head_style)
        self.sheet1.write(self.f_index, 1, u'启动次数', self.head_style)
        self.sheet1.write(self.f_index, 2, u'LIB算法(℃)', self.head_style)
        self.sheet1.write(self.f_index, 3, u'Amlogic算法(℃)', self.head_style)
        self.sheet1.write(self.f_index, 4, u'启动失败次数', self.head_style)
        self.sheet1.write(self.f_index, 5, u'运行时崩溃次数', self.head_style)
        
        self.sheet1.col(0).width = 512 * (len(u'设备编号') + 1)
        self.sheet1.col(1).width = 512 * (len(u'启动次数') + 1)
        self.sheet1.col(2).width = 512 * (len(u'LIB算法(℃)') + 1)
        self.sheet1.col(3).width = 512 * (len(u'Amlogic算法(℃)') + 1)
        self.sheet1.col(4).width = 512 * (len(u'启动失败次数') + 1)
        self.sheet1.col(5).width = 512 * (len(u'运行时崩溃次数') + 1)
        
        # 让出第一行, 准备填充数据
        self.f_index += 1
    
    def process(self):
        # 如果路径是文件夹, 分别遍历所有文件, 一一进行处理
        if(os.path.isdir(self.path)):
            for root, dirs, files in os.walk(self.path):
                for name in files:
                    f_name = os.path.join(root, name)
                    print f_name
                    if(os.path.isdir(f_name)):
                        for _root, _dirs, _files in os.walk(f_name):
                            for _name in _files:
                                _f_name = s.path.join(_root, _name)
                                print _f_name
                                self.process_single_file(_f_name)
                    else:
                        self.process_single_file(f_name)
                    #print f_name
                    #sys.exit(0)
        else:
            f_name = self.path
            self.process_single_file(f_name)
        # 保存工作簿数据到文件
        self.workbook.save(sys.argv[2])

    def process_single_file(self, f_name):
        if not os.path.isfile(f_name):
            print '路径不是文件'
            sys.exit(-1)
            
        # 统计启动次数
        boot_times = 0
        # 最大结温(LY计算)
        max_ly_temp = 0
        # 最大结温(Amlogic计算)
        max_cpu_temp = 0
        # 最近一次启动的行数, 用于探测是否出现启动异常
        last_boot = 0
        # 打开文件
        file_obj = open(f_name)
        # 当前行号
        line_num = 0
        err_lines = []
        err_contents = []
        crash_lines = []
        crash_contents = []
        ddr_lines = []
        ddr_contents = []
        pll_lines = []
        pll_contents = []
        for line in file_obj:
            if line.find('test start') == 0:
                temp = linecache.getline(f_name, line_num - 1).strip('\n')
                #print temp
                try:
                    temp_temp = int(temp)
                    if temp_temp < 200:
                        if temp_temp > max_cpu_temp:
                            max_cpu_temp = temp_temp
                except:
                    pass
            # 如果开头是'ly'则认定为烤机的输出
            if line.find('ly:') == 0:
                # 对数据项目进行分割, 例如:'ly:123|temp:200'会分割为:['ly:123', 'temp:200']
                datas = line.split('|')
                # 将list[]中的每一项拆成两部分, 并忽略出错的情形
                try:
                    # 统计ly的数据中的最大值
                    temp_temp = int(datas[0].split(':')[1])
                    if temp_temp < 200:
                        if temp_temp > max_ly_temp:
                            max_ly_temp = int(datas[0].split(':')[1])
                except:
                    pass   
                # 同样统计temp中数据最大值
                try:
                    temp_temp = int(datas[1].split(':')[1])
                    if temp_temp < 300:
                        if temp_temp > max_cpu_temp:
                            max_cpu_temp = int(datas[1].split(':')[1])
                except:
                    pass
            if line.find('test fail') >= 0:
                print 'boot time: ', boot_times
                ddr_lines.append(line_num)
                ddr_contents.append(line)
                pass
            if line.find('pll_times') == 0:
                pll_lines.append(line_num)
                pll_contents.append(line)
            # 由于每次启动的头两行可能出现乱码, 因此以启动后的5行左右内容为基准
            if line.find('DDR mode') >= 0:# and len(line) == 23:
                # 统计启动次数
                # 统计距离上一次启动时的日志行数差
                diff = line_num - last_boot
                # 如果两次复位间隔日志出现的行号在10到1000之间, 说明复位存在异常(根据一次正常的启动过程参考)
                if boot_times > 0:
                    if diff > 10 and diff < 800:
                        print '问题出现在第%d次启动' % (boot_times)
                        # 统计错误的行号和内容, 注意内容通过其它库索引
                        err_lines.append(line_num)
                        # print "[diff:%d]%d:%s(%s)" % (diff, (line_num - 1), line.strip('\r\n'), linecache.getline(f_name, line_num - 9).strip('\n'))
                        err_contents.append(linecache.getline(f_name, line_num - 9).strip('\n'))
                    # 如果两次复位间隔在1000到2000之间, 则说明烤机过程中出现崩溃(根据一次正常的完整测试参考)
                    elif diff > 1800 and diff < 2900:
                        # print 'line: ', diff
                        crash_lines.append(line_num)
                        crash_contents.append(linecache.getline(f_name, line_num - 8).strip('\n'))
                # 记录最后一次启动的行号
                last_boot = line_num
                boot_times = boot_times + 1
                #print boot_times
            # 记录当前行号
            line_num += 1
            # old_line = line
        # 用于计算表项宽度
        line_width = 0
        # 用于设置表格的风格
        write_style = self.content_style
        # 如果该设备日志中的崩溃统计行号数大于0, 则认定存在崩溃问题, 进行样式重置
        if len(crash_lines) > 0:
            write_style = self.warning_style
            # 终端输出错误提示
            print '文件:%s中有%d次可能的崩溃异常' % (os.path.basename(f_name), len(crash_lines))
            err_idx = 1
            # 新建工作表, 准备填写错误行号和对应的错误内容
            # sheet = self.workbook.add_sheet(u'%s 崩溃详情' % os.path.basename(f_name), cell_overwrite_ok=True)
            f_name_temp = filter(lambda x: x in string.printable, f_name)
            sheet = self.workbook.add_sheet(u'%s|C' % f_name_temp.replace('/', '|').strip('.').replace('_', '|')[:28], cell_overwrite_ok=True)
            # 填写详情表头
            sheet.write(0, 0, u'运行崩溃行号', self.head_style)
            sheet.write(0, 1, u'运行崩溃内容', self.head_style)
            # 填写详情信息
            for (line, each) in zip(crash_lines, crash_contents):
                sheet.write(err_idx, 0, line - 8, self.err_style)
                # 这里做一个特殊处理, 对于有些日志的行, 存在特殊字符将导致表格保存时失败, 因此filter掉特殊字符
                sheet.write(err_idx, 1, filter(lambda x: x in string.printable, each), self.err_style)
                # 这里算出最多错误内容的宽度, 后续进行设置宽度时使用
                if len(each) > line_width:
                    line_width = len(each)
                err_idx = err_idx + 1
            # 根据内容最多的行, 设置此列的宽度(自动)
            sheet.col(4).width = 256 * (line_width + 1)
            # 添加新工作表到当前工作簿
            self.sheets.append(sheet)
        # 与崩溃信息处理基本相同
        if len(err_lines) > 0:
            write_style = self.warning_style
            print '文件:%s中有%d次可能的启动异常' % (os.path.basename(f_name), len(err_lines))
            err_idx = 1
            f_name_temp = filter(lambda x: x in string.printable, f_name)
            sheet = self.workbook.add_sheet(u'%s|B' % f_name_temp.replace('/', '|').strip('.').replace('_', '|')[:28], cell_overwrite_ok=True)
            sheet.write(0, 0, u'启动失败行号', self.head_style)
            sheet.write(0, 1, u'启动失败内容', self.head_style)             
            for (line, each) in zip(err_lines, err_contents):
                # print line, filter(lambda x: x in string.printable, each)           
                sheet.write(err_idx, 0, line - 9, self.err_style)
                sheet.write(err_idx, 1, filter(lambda x: x in string.printable, each), self.err_style)
                if len(each) > line_width:
                    line_width = len(each)
                err_idx = err_idx + 1
            sheet.col(5).width = 256 * (line_width + 1)
            self.sheets.append(sheet)
        if len(ddr_contents) > 0:
            write_style = self.warning_style
            print '文件:%s中有%d次DDR Failed' % (os.path.basename(f_name), len(ddr_contents))
            err_idx = 1
            f_name_temp = filter(lambda x: x in string.printable, f_name)
            sheet = self.workbook.add_sheet(u'%s|P' % f_name_temp.replace('/', '|').strip('.').replace('_', '|')[:28], cell_overwrite_ok=True)
            sheet.write(0, 0, u'DDR测试失败行号', self.head_style)
            sheet.write(0, 1, u'DDR测试失败内容', self.head_style)             
            for (line, each) in zip(ddr_lines, ddr_contents):
                # print line, filter(lambda x: x in string.printable, each)           
                sheet.write(err_idx, 0, line)
                sheet.write(err_idx, 1, filter(lambda x: x in string.printable, each), self.err_style)
                if len(each) > line_width:
                    line_width = len(each)
                err_idx = err_idx + 1
            sheet.col(6).width = 256 * (line_width + 1)
            self.sheets.append(sheet)  
        if len(pll_contents) > 0:
            write_style = self.warning_style
            print '文件:%s中有%d次PLL失锁' % (os.path.basename(f_name), len(pll_contents))
            err_idx = 1
            f_name_temp = filter(lambda x: x in string.printable, f_name)
            sheet = self.workbook.add_sheet(u'%s|P' % f_name_temp.replace('/', '|').strip('.').replace('_', '|')[:28], cell_overwrite_ok=True)
            sheet.write(0, 0, u'PLL失锁行号', self.head_style)
            sheet.write(0, 1, u'PLL失锁失败内容', self.head_style)             
            for (line, each) in zip(pll_lines, pll_contents):
                # print line, filter(lambda x: x in string.printable, each)           
                sheet.write(err_idx, 0, line)
                sheet.write(err_idx, 1, filter(lambda x: x in string.printable, each), self.err_style)
                if len(each) > line_width:
                    line_width = len(each)
                err_idx = err_idx + 1
            sheet.col(7).width = 256 * (line_width + 1)
            self.sheets.append(sheet)           
        # 对此前统计的信息进行建表
        # 通过日志文件名解析设备编号
        # self.sheet1.write(self.f_index, 0, os.path.basename(f_name).split('-')[1], self.content_style)
        self.sheet1.write(self.f_index, 0, f_name, self.content_style)
        # 保存启动次数
        self.sheet1.write(self.f_index, 1, boot_times, self.content_style)
        # 保存lib最大结温
        self.sheet1.write(self.f_index, 2, max_ly_temp, self.content_style)
        # 保存amlogic最大结温
        self.sheet1.write(self.f_index, 3, max_cpu_temp, self.content_style)
        # 对于有运行时崩溃的表项, 应用警告样式
        if len(err_lines) > 0:
            self.sheet1.write(self.f_index, 4, len(err_lines), write_style)
        else:
            self.sheet1.write(self.f_index, 4, len(err_lines), self.content_style)
        #self.sheet1.write(self.f_index, 4, len(err_lines), write_style)
        #self.sheet1.write(self.f_index, 5, len(crash_lines), write_style)
        # 对于启动异常的表项, 应用警告样式
        if len(crash_lines) > 0:
            self.sheet1.write(self.f_index, 5, len(crash_lines), write_style)
        else:
            self.sheet1.write(self.f_index, 5, len(crash_lines), self.content_style)
        # 检出错误次数
        self.sheet1.write(1, 6, u'DDR测试错误次数', self.head_style)
        self.sheet1.col(6).width = 512 * (len(u'DDR测试错误次数') + 1)         
        if len(ddr_lines) > 0:        
            self.sheet1.write(self.f_index, 6, len(ddr_lines), self.warning_style)
        else:
            self.sheet1.write(self.f_index, 6, len(ddr_lines), self.content_style)
            
        self.sheet1.write(1, 7, u'PLL失锁次数', self.head_style)
        self.sheet1.col(6).width = 512 * (len(u'PLL失锁次数') + 1)         
        if len(pll_lines) > 0:        
            self.sheet1.write(self.f_index, 7, len(pll_lines), self.warning_style)
        else:
            self.sheet1.write(self.f_index, 7, len(pll_lines), self.content_style)

        err_lines = []
        err_contents = []
        crash_lines = []
        crash_contents = []
        ddr_lines = []
        ddr_contents = []
        pll_lines = []
        pll_contents = []
        file_obj.close()
        self.f_index = self.f_index + 1
                
if __name__ == '__main__':
    if len(sys.argv) < 3:
        print '请给出日志存储路径在第一个参数(%d)' % len(sys.argv)
        sys.exit(-1)
    log = LBRPyExcel(sys.argv[1])
    log.prepare()
    log.process()
