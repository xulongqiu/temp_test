#!/usr/bin/python
#coding=utf-8

import xlsxwriter
import os.path
import sys
import linecache
import string
import pylab as pl
import matplotlib
import numpy as np
from scipy import stats
from ly_mail import mail_sender

class lbtNewLogProcessor:
    def __init__(self, path):
        self.cal_datas = {}
        self.path = path
        self.plot_idx = 1
        self.axss = []
        self.datas = []
        self.f_names = []
        self.global_line_num = 0
        self.isfile = False
        self.process()

    def process(self):
        if(os.path.isdir(self.path)):
            for root, dirs, files in os.walk(self.path):
                for name in files:
                    self.f_name = os.path.join(root, name)
                    if(os.path.isdir(self.f_name)):
                        for _root, _dirs, _files in os.walk(self.f_name):
                            for _name in _files:
                                _f_name = s.path.join(_root, _name)
                                self.process_single_file(_f_name)
                    else:
                        self.process_single_file(self.f_name)
        else:
            self.isfile = True
            print 'isfile'
            self.process_single_file(self.path)
        if not self.isfile:
            maxes_val = []
            maxes_cnt = []
            tmp_idx = 0
            print 'self.datas lenght: ', len(self.datas)
            if len(self.datas):
                disp_font = matplotlib.font_manager.FontProperties(
                    fname='/usr/share/fonts/truetype/wqy/wqy-microhei.ttc'
                )
                new_data = []
                for tmp_data in self.datas:
                    maxes_val.append(max(tmp_data))
                    maxes_cnt.append(len(tmp_data))
                    new_data.append(tmp_data[len(tmp_data) / 2])
                    print 'plot index:', tmp_idx
                    pl.plot(range(len(tmp_data)), 
                            tmp_data, 
                            label=os.path.basename(self.f_names[tmp_idx]), 
                            #marker='.',
                            )
                    tmp_idx = tmp_idx + 1
                print new_data
                pl.title(u'结温数据对比', fontproperties=disp_font)
                pl.xlabel(u'结温采集点', fontproperties=disp_font)
                pl.ylabel(u'结温数据(平均值)', fontproperties=disp_font)
                #pl.xlim(min(max(maxes_cnt)) - 20, max(maxes_cnt))
                #pl.ylim(min(maxes_val) - 20, max(maxes_val) + 20)
                if len(self.datas) < 10:
                    pl.legend(loc='lower right', prop={'size': 10})
                pl.grid(False)
                ax = pl.axes()
                ax.yaxis.grid()
                servers = {'163':'send_163', 'libratone':'send_libratone'}
                #mail = mail_sender()
                reciever = 'eric.xu@libratone.com.cn, james.li@libratone.com.cn'
                subject = u'结温温升数据'
                img_name = 'all.png'
                pl.savefig(img_name)
                #mail.add_image(img_name, img_name)
                body = u'''
                <h1>
                    <p>Hi, all:</p>
                </h1>
                <p>结温图表:</p>
                <img src="cid:%s">
                <table>
                <tr><th>文件</th><th>众数</th><th>极差</th><th>方差</th><th>标准差</th><th>变异系数</th></tr>
                ''' % img_name
                for key, val in self.cal_datas.iteritems():
                    print key, type(float(val['mean']))
                    body += ('<tr><td>%s</td><td>%f</td><td>%d</td><td>%f</td><td>%f</td><td>%f</td></tr>' % 
                             (key, 
                              float(val['mode'][0]), 
                              val['ptp'], 
                              val['var'], 
                              val['std'], 
                              float(val['mean'])
                             ))
                body += '</table>'
                print body
                #mail.add_html(body)
                #mail.set_sender('james.li@libratone.com.cn')
                #mail.set_cc('eric.xu@libratone.com.cn')
                #mail.set_bcc('eric.xu@libratone.com.cn')
                #mail.send(reciever, subject)
                # pl.show()
    def drawHist(self, heights):
        #创建直方图
        #第一个参数为待绘制的定量数据，不同于定性数据，这里并没有事先进行频数统计
        #第二个参数为划分的区间个数
        pl.figure(2)
        print heights
        data_cnt = []
        for item in heights:
            if item not in data_cnt:
                data_cnt.append(item)
        pl.xlim(0.0, len(data_cnt))
        pl.ylim(0.0, max(heights))
        pl.xlabel(u'温度值')
        pl.ylabel(u'出现次数')
        pl.title(u'稳定结温频次直方图')
        pl.hist(heights, len(data_cnt))
        pl.show()

    def process_single_file(self, f_name):
        if f_name.find('data') > 0:
            return
        if f_name.find('png') > 0:
            return
        if f_name.find('zip') > 0:
            return      
        self.f_names.append(os.path.basename(f_name))
        print f_name
        line_num = 0
        temp_datas = []
        temp_data = []
        lens = []
        boot_times = 0
        max_cpu_temp = 0
        
        try:
            file_obj = open(f_name)
        except:
            print '没有输入文件'
        for line in file_obj:
            if line.find('normal') == 0:
                datas = line.split('|')
                self.global_line_num = self.global_line_num + 1
                try:
                    temp_temp = int(datas[1].split(':')[1])
                    temp_data.append(temp_temp)
                    if temp_temp < 300:
                        if temp_temp > max_cpu_temp:
                            max_cpu_temp = temp_temp
                except:
                    pass
            #if line.find('Libratone Temperature') >= 0:
            if line.find('This cycle test of') >= 0:
                boot_times = boot_times + 1
                print 'get one cycle data, len:%d' % len(temp_data)
                if len(temp_data):
                    print 'get one cycle data'
                    temp_datas.append(temp_data)
                    lens.append(len(temp_data))
                    temp_data = []
            # 记录当前行号
            line_num += 1
        if self.isfile:
            print 'isFile'
            disp_font = matplotlib.font_manager.FontProperties(fname='/usr/share/fonts/truetype/wqy/wqy-microhei.ttc')
            maxes_val = []
            maxes_cnt = []
            pl.figure(1)
            #main_fig = pl.subplot(211)
            if len(temp_datas) > 0:
                print 'have %d series...' % len(temp_datas)
                new_data = []
                # 每组数据
                for temp_temp_data in temp_datas:
                    maxes_val.append(max(temp_temp_data))
                    maxes_cnt.append(len(temp_temp_data))
                    new_data.append(temp_temp_data[len(temp_temp_data) / 2])
                    pl.plot(range(0, len(temp_temp_data), 1), temp_temp_data)
            #if len(temp_data):
                #pl.plot(range(20, len(temp_data) + 20, 1), temp_data)
            pl.title(f_name, fontproperties=disp_font)
            pl.xlabel(u'结温采集点', fontproperties=disp_font)
            pl.ylabel(u'结温数据(平均值)', fontproperties=disp_font)
            if len(maxes_cnt):
                pl.xlim(0.0, max(maxes_cnt) + 40)
            else:
                pl.xlim(0.0, len(temp_data) + 40)
            if len(maxes_val):
                pl.ylim(0.0, max(maxes_val) + 10)
            else:
                pl.ylim(0.0, max(temp_data) + 10)
            pl.grid(False)
            ax = pl.axes()
            ax.yaxis.grid()
            pl.show()
            return        
        if len(temp_datas) > 0:
            print 'have %d series...' % len(temp_datas)
            new_data = []
            # 每组数据
            for temp_temp_data in temp_datas:
                new_data.append(temp_temp_data[len(temp_temp_data) / 2])
            cal_data = {}
            cal_data['mode'] = stats.mode(new_data)
            cal_data['ptp'] = np.ptp(new_data)
            cal_data['var'] = np.var(new_data)
            cal_data['std'] = np.std(new_data)
            cal_data['mean'] = np.mean(new_data) / np.std(new_data)
            self.cal_datas[os.path.basename(f_name)] = cal_data
            print f_name + ':'
            print u'\t众数: ', stats.mode(new_data)
            print u'\t极差: ', np.ptp(new_data)
            print u'\t方差: ', np.var(new_data)
            print u'\t标准差: ', np.std(new_data)
            print u'\t变异系数: ', np.mean(new_data) / np.std(new_data)                
            #self.cal_datas.append(, cal_data)            
            '''

            data_cnt = {}
            for item in new_data:
                item_str = str(item)
                if not data_cnt.has_key(item_str):
                    data_cnt[item_str] = 1
                else:
                    data_cnt[item_str] = data_cnt[item_str] + 1
            pl.xlim(min(new_data) - 2, max(new_data) + 2)
            pl.ylim(0.0, max(new_data))
            pl.xlabel(u'温度值', fontproperties=disp_font)
            pl.ylabel(u'出现次数', fontproperties=disp_font)
            pl.title(u'稳定结温频次直方图', fontproperties=disp_font)
            pl.hist(new_data, max(new_data))
            pl.show()    
            '''
        if len(lens):
            len_min = min(lens)
            average_arr = []
            for i in range(len_min):
                sum_res = 0
                for j in range(len(temp_datas)):
                    sum_res = sum_res + temp_datas[j][i]
                    #if(i == 0):
                        #print 'data[%d][%d]: %d' % (j, i, temp_datas[j][i])
                avr = float(sum_res) / float(len(temp_datas))
                average_arr.append(avr)
            self.datas.append(average_arr)
            self.plot_idx = self.plot_idx + 1

if __name__ == '__main__':
    processor = lbtNewLogProcessor(sys.argv[1])
