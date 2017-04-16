#!/usr/bin/python
#coding=utf-8

import os.path
import sys
import xlrd
import xlwt
import linecache
import string
import pylab as pl
import matplotlib
import numpy as np


class Xls2Plot:
    def __init__(self, path):
        # 保存路径信息
        self.path = path
        self.boars = {}
        pass
    
    def xml_prepare(self, xml_name):
        # 保存输出xml文件名
        self.xml_name = xml_name
        # 创建工作簿
        self.workbook = xlwt.Workbook(encoding='utf-8')
        # 创建工作表
        self.sheets = []
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
        
        # 依次写入表头
        self.sheet_summery = self.workbook.add_sheet(os.path.basename('summery'), cell_overwrite_ok=True)
        self.f_index = 0
        self.sheet_temp.write_merge(0, 0, 0, 6, u'提示: 此表格由日志分析程序自动生成', self.pro_style)
        self.f_index += 1
        self.sheet_temp.write(self.f_index, 0, u'boardNum', self.head_style)
        self.sheet_temp.write(self.f_index, 1, u'freq', self.head_style)
        self.sheet_temp.write(self.f_index, 2, u'samples', self.head_style)
        self.sheet_temp.write(self.f_index, 3, u'cycles', self.head_style)
        self.sheet_temp.write(self.f_index, 4, u'startTemp', self.head_style)
        self.sheet_temp.write(self.f_index, 5, u'stableTemp', self.head_style)
        self.sheet_temp.write(self.f_index, 6, u'Note', self.head_style)
    
        self.sheet_temp.col(0).width = 512 * (len(u'boardNum') + 1)
        self.sheet_temp.col(1).width = 512 * (len(u'freq') + 1)
        self.sheet_temp.col(2).width = 512 * (len(u'samples') + 1)
        self.sheet_temp.col(3).width = 512 * (len(u'cycles') + 1)
        self.sheet_temp.col(4).width = 512 * (len(u'startTemp') + 1)
        self.sheet_temp.col(5).width = 512 * (len(u'stableTemp') + 1)    
        self.sheet_temp.col(6).width = 512 * (len(u'Note') + 1)
        pass
        
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
                        self.process_single_file(os.path.basename(f_name))
                    #print f_name
                    #sys.exit(0)
        else:
            f_name = self.path
            self.process_single_file(os.path.basename(f_name))
            
        self.show_curves()
        #self.save_xml()

    def process_single_sheet(self, sheet, colnameindex=0):
        name = sheet.name
        nrows = sheet.nrows          
        print 'sheet:name=%s,nrows=%d' % (name, nrows) 
        colnames =  sheet.row_values(colnameindex)
        all_cases = {}

        #parse sheet
        for i in range(colnameindex + 1, nrows):
            row = sheet.row_values(i) 
            app = {}
            if row:
                for i in range(len(colnames)):
                    app[colnames[i]] = row[i]
                    
                if app.get('case') == 'memtest': 
                    continue
                
                case = all_cases.get(app.get('case'))
                cycle_key = 'cycle:' + str(int(app.get('cycles')))
                if case:
                    cycle = case.get(cycle_key)
                    if cycle:
                        cycle.append(app.get('temp'))
                    else:
                        cycle = []
                        cycle.append(app.get('temp'))
                        case[cycle_key] = cycle
                    pass
                else:
                    case = {}
                    cycle = []
                    cycle.append(app.get('temp'))
                    case[cycle_key] = cycle
                    all_cases[app.get('case')] = case
                    pass
                        
        #calculate average for all cycles of every case  
        for key in all_cases:
            case = all_cases[key]
            if len(case) > 0:
                len_tmp = 0
                max_len = 0
                for cycles_key in case:
                    len_tmp = len(case[cycles_key])
                    if len_tmp > max_len:
                        max_len = len_tmp  
                    print 'cycle:', cycles_key, len_tmp    
                average = []
                for i in range(max_len):
                    total = 0
                    cnt = 0
                    cycle = []
                    for cycles_key in case:
                        cycle = case[cycles_key]
                        if len(cycle) > i:
                            total += cycle[i]
                            cnt += 1
                    if cnt > 0:      
                        average.append(int(total/cnt))
                    
                case['average'] = average                
        
            
        if len(all_cases) > 0:
            self.boars[name] = all_cases
        
    def process_single_file(self, f_name):
        print 'f_name: ', f_name
        if not os.path.isfile(f_name):
            print '路径不是文件'
            sys.exit(-1)
 
        if f_name.find('.') >= 0:    
            temp = f_name.split('.')
            if temp[len(temp) - 1] != "xls":
                return 
           
        workbook = xlrd.open_workbook(os.path.basename(f_name))
        
        for i in range(workbook._all_sheets_count):
            sheet = workbook.sheet_by_index(i)
            self.process_single_sheet(sheet, 1)
            pass
        
    def show_curves(self):
        for board_name in self.boars:
            board_cases = {}
            kaoji_cases = {}
            all_cases = {}
            print 'boards iter: key=', board_name
            all_cases = self.boars[board_name]
            case = {}
            for cases_key in all_cases:
                case = all_cases[cases_key]
                cycle = []
                if len(case) > 0 and cases_key.find('kj') < 0:
                    self.save_one_boards(case, board_name.split('.')[0]+ '-' + cases_key)
                    pass
                for cycles_key in case:
                    cycle = case[cycles_key]
                    temp_average = 0
                    for i in range(100, len(cycle)-40):
                        temp_average += cycle[i]
                    temp_average = temp_average // (len(cycle)-140)
                    stable = ''
                    stable += cases_key
                    stable += '-'
                    stable += str(int(temp_average))
                    if cycles_key == 'average':  
                        if cases_key.find('kj') < 0:
                            #self.sheet_temp.write(self.f_index, 1, freq_tmp, self.content_style);
                            board_cases[stable] = cycle
                        else:
                            kaoji_cases[stable] = cycle
                        pass
                    pass
                pass
            if len(board_cases) > 0:
                self.save_one_boards(board_cases, board_name.split('.')[0]+'-average')
                self.save_one_boards(kaoji_cases, board_name.split('.')[0]+'-kaoji')
            pass
        pass
    
        
    def save_one_boards(self, dict_data, f_name):
        disp_font = matplotlib.font_manager.FontProperties(fname='/usr/share/fonts/truetype/ubuntu-font-family/Ubuntu-B.ttf')
        maxes_val = []
        maxes_cnt = []
        min_val = []
        pl.figure(figsize=(9, 6))
        #main_fig = pl.subplot(211)
        if len(dict_data) > 0:
            print 'have %d series...' % len(dict_data)
            new_data = []
            # 每组数据
            for key in dict_data:
                temp_temp_data = dict_data[key]
                maxes_val.append(max(temp_temp_data))
                maxes_cnt.append(len(temp_temp_data))
                min_val.append(min(temp_temp_data))
                new_data.append(temp_temp_data[len(temp_temp_data) / 2])
                pl.plot(range(0, len(temp_temp_data), 1), temp_temp_data, label=key)
                #p = np.polyfit(range(0, 300, 1),temp_temp_data[0:300], 9)
                #y_av = np.polyval(p, range(0, 300, 1))
                #pl.plot(range(0, 300, 1), y_av, 'r', label=key)
                #p = np.polyfit(range(0, len(temp_temp_data)-300, 1),temp_temp_data[300:], 9)
                #y_av = np.polyval(p, range(0, len(temp_temp_data) - 300, 1))
                #pl.plot(range(300, len(temp_temp_data), 1), y_av, 'r')                
                pl.legend()
                

        pl.title(f_name, fontproperties=disp_font)
        pl.xlabel(u'samples(time)', fontproperties=disp_font)
        pl.ylabel(u'temperature(C)', fontproperties=disp_font)
        if len(maxes_cnt):
            pl.xlim(0.0, max(maxes_cnt) + 40)
        if len(maxes_val):
            pl.ylim(min(min_val) - 20, max(maxes_val) + 10)
            
        pl.grid(False)
        ax = pl.axes()
        ax.yaxis.grid()
        #ax.xaxis.grid()
        #pl.show()  
        curve_name = f_name + '.png'
        pl.savefig(curve_name)
        pass
    
    def save_xml(self):
        self.workbook.save(sys.xml_name)
        pass
                            
if __name__ == '__main__':
    if len(sys.argv) < 2:
        print 'xml filename sys.argc=(%d)' % len(sys.argv)
        sys.exit(-1)   
        
    #print sys.argv[0], sys.argv[1], sys.argv[2]
    plot = Xls2Plot(sys.argv[1])
    #plot.xml_prepare(sys.argv[2])
    plot.process()
