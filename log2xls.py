#!/usr/bin/python
#coding=utf-8

import os.path
import sys
import xlrd
import xlwt
import linecache
import string
import SheetBaseClass


class LBRPyExcel:
    def __init__(self, path):
        # 保存路径信息
        self.path = path
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
        pass
    
    def process(self):
        # 如果路径是文件夹, 分别遍历所有文件, 一一进行处理
        if(os.path.isdir(self.path)):
            for root, dirs, files in os.walk(self.path):
                for name in files:
                    f_name = os.path.join(root, name)
                    if(os.path.isdir(f_name)):
                        for _root, _dirs, _files in os.walk(f_name):
                            for _name in _files:
                                _f_name = s.path.join(_root, _name)
                                self.process_single_file(_f_name)
                    else:
                        self.process_single_file(f_name)
        else:
            f_name = self.path
            self.process_single_file(f_name)
        # 保存工作簿数据到文件
        self.workbook.save(sys.argv[2])

    def process_single_file(self, f_name):
        if not os.path.isfile(f_name):
            print '路径不是文件'
        if f_name.find('xls') > 0:
            return
        if f_name.find('png') > 0:
            return
        if f_name.find('zip') > 0:
            return              
        print 'f_name: ', f_name
        sheet_temp = SheetBaseClass.SheetBaseClass(self.workbook, os.path.basename(f_name)) #self.workbook.add_sheet(os.path.basename(f_name), cell_overwrite_ok=True)
        # 打开文件
        file_obj = open(f_name)
 
        # 依次写入表头
        head = [u'case', u'freq(Hz)', u'coreCnt', u'time(secs)', u'cpuTemp(°C)', u'ModuleTemp(°C)', u'cycles', u'Note']
        
        # 由于每一行都会对cycle赋值， 因此某次循环检测到的cycles start会被normal的行覆盖掉， 因此， cycle永远为0
        cycle = 0
        ubootTemp = 0
        while 1:
            updown = ''
            line = file_obj.readline()
            if not line:
                break
        
            if line.find('kj600000|') == 0:
                break
            #ubootTemp|temp:27|time:0
            if line.find('ubootTemp|') == 0: 
                if ubootTemp == 0:
                    cycle = cycle + 1;
                    if cycle == 3:
                        cycle = 1
                ubootTemp = int(line.split('|')[1].split(':')[1])
                #print 'cycle=', cycle, line
            
            # 提取常温下cycle的曲线
            if line.find('envTemp:') == 0:
                #envTemp:15|temp:28|adc:|1.1.1.1|freq:1536000|time:9.518|mtemp:16|points:1|up
                #==> {"envTemp:15", "temp:28", "adc:", "1.1.1.1", "freq:153600", "time:9.518", "mtemp:91", "points:70", "up"}
                
                datas = line.split('|')
                if len(datas) <= 8:
                    line = line + file_obj.readline()
                    datas = line.split('|')
                    
                if ubootTemp != 0:
                    sheet_temp.addSheetHead(head)
                    content = [datas[0].split(':')[1], None, None, 0, ubootTemp, None, None, 'uboot']
                    sheet_temp.addSheetRow(content);
                    ubootTemp = 0
                try:
                    # 提取频率数据
                    content= []
                    coreCnt = 0;
                    cores = datas[3].split('.');
                    for i in range(len(cores)):
                        if cores[i] == '1':
                            coreCnt += 1;
                    #print "CORE_CNT:", coreCnt
                    freq_tmp = int(datas[4].split(':')[1])
                    #print "freq_tmp:", freq_tmp
                    # 提取时间戳数据
                    time_tmp = float(datas[5].split(':')[1])
                    #print "time_tmp:", time_tmp
                    temp_tmp = int(datas[1].split(':')[1])
                    #print "temp_tmp:", temp_tmp
                    temp_tmp1 = int(datas[6].split(':')[1])
                    #print "temp_tmp1:", temp_tmp1
                    
                    if datas[8].find('d') >= 0:
                        updown = 'down'
                    elif datas[8].find('u') >= 0:    
                        updown = 'up'
                    
                    #case 
                    content.append(datas[0].split(':')[1])
                    #freq
                    if freq_tmp > 0:
                        content.append(freq_tmp)
                    else:
                        content.append(None)
                    #coreCnt
                    content.append(coreCnt)
                    #time 
                    if time_tmp > 0:
                        content.append(time_tmp)
                    else:
                        content.append(None)                    
                    
                    # 提取结温数据
                    if temp_tmp < 200:
                        content.append(temp_tmp)
                    else:
                        content.append(None)
                        
                    # 提取模块温度
                    if temp_tmp1 < 125:
                        content.append(temp_tmp1);
                    else:
                        content.append(None)
                    if cycle >= 0:
                        content.append(cycle);
                    else:
                        content.append(None)
                        
                    content.append(updown);
                    
                    if len(content) > 0:
                        sheet_temp.addSheetRow(content)
                    #print "freq: %d, core: %d, time: %d, cpuTemp: %d, moduleTemp:%d" % (freq_tmp, coreCnt, time_tmp, temp_tmp, temp_tmp1)
                except:
                    pass
                
if __name__ == '__main__':
    if len(sys.argv) < 3:
        print 'xml filename sys.argc=(%d)' % len(sys.argv)
        sys.exit(-1)   
    print sys.argv[0] 
    print sys.argv[1]
    print sys.argv[2]
    log = LBRPyExcel(sys.argv[1])
    log.process()
