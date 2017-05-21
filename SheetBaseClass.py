#!/usr/bin/python
#coding=utf-8

import os.path
import sys
import xlrd
import xlwt

class SheetBaseClass:
    def __init__(self, work_book, sheet_name):
        self.rowBase = 0
        self.row = 0
        self.colBase= 0
        self.col = 0
        self.name = sheet_name
        
        # sheet head style
        self.head_style = xlwt.XFStyle()
        self.head_style = xlwt.easyxf('pattern: pattern solid, fore_colour black; font: colour white, bold on;');
        # center aligin
        self.head_style.alignment.horz = xlwt.Alignment.HORZ_CENTER
        # font style
        self.head_style.font.name = 'WenQuanYi Zen Hei'
        self.head_style.font.bold = True
    
        # content style
        self.content_style = xlwt.XFStyle()
        self.content_style.alignment.horz = xlwt.Alignment.HORZ_CENTER
        self.content_style.font.name = 'WenQuanYi Zen Hei'
        self.content_style.font.bold = False
    
        # warning content style
        self.warning_style = xlwt.easyxf('pattern: pattern solid, fore_colour red; font: colour black, bold on;');
        self.warning_style.font.name = 'WenQuanYi Zen Hei'
        self.warning_style.alignment.horz = xlwt.Alignment.HORZ_CENTER
    
        # error content style
        self.err_style = xlwt.easyxf('pattern: pattern solid, fore_colour red; font: colour black, bold on;');
        self.err_style.font.name = 'WenQuanYi Zen Hei'
        self.err_style.alignment.horz = xlwt.Alignment.HORZ_LEFT
    
        # pro content style
        self.pro_style = xlwt.easyxf('pattern: pattern solid, fore_colour yellow; font: colour black, bold on;');
        self.pro_style.font.name = 'WenQuanYi Zen Hei'
        self.pro_style.alignment.horz = xlwt.Alignment.HORZ_CENTER         
        
        #create sheet in work_book
        try:
            self.sheet = work_book.get_sheet(sheet_name)
        except:
            self.sheet = work_book.add_sheet(sheet_name, cell_overwrite_ok=True)
            pass
        
    def getSheet(self):
        return self.sheet
    
    def sheetMerge(self, row_start, row_cnt, col_start, col_cnt, str_content):
        if self.sheet:
            self.sheet.write_merge(row_start, row_start + row_cnt -1, col_start, col_start + col_cnt -1, str_content, self.pro_style)
            pass
        else:
            pass 
        
    def addSheetHead(self, head):
        if self.sheet:
            self.colBase = self.col
            self.sheetMerge(0, 1, self.colBase, len(head), u'提示: 此表格由日志分析程序自动生成')
            self.rowBase = 1
            self.row = 1
            
            for i in range(len(head)):
                self.sheet.write(self.row, self.colBase + i, head[i], self.head_style)
                self.sheet.col(self.colBase + i).width = 300 * (len(head[i]) + 1)
            self.row += 1
            self.rowBase += 1
            self.col += len(head) + 1
            pass
        
        else:
            pass
        pass

    def addSheetRow(self, content):
        sheet = self.sheet
        if sheet:
            for i in range(len(content)):
                sheet.write(self.row, self.colBase + i, content[i], self.content_style)
                width = sheet.col(i + self.colBase).width
                if width < 350 * (len(str(content[i])) + 1):
                    sheet.col(i + self.colBase).width = 350 * (len(str(content[i])) + 1)
            self.row += 1
            self.rowBase += 1
            pass
        else:
            pass
        pass   
    
    def addSheetCol(self, content):
        sheet = self.sheet
        if sheet:
            self.row = self.rowBase
            width = sheet.col(self.colBase).width
            for i in range(len(content)):
                sheet.write(self.row, self.colBase, content[i], self.content_style)
                if width < 300 * (len(str(content[i])) + 1):
                    sheet.col(col_index).width = 300 * (len(str(content[i])) + 1)
                self.row += 1
            self.colBase += 1
            pass
        else:
            pass
        pass   
    
if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Usage %s work_name sheet_name" % (sys.argv[0]))
        sys.exit(-1)
        pass
    
    print(sys.argv[0])
    print(sys.argv[1])  
    print(sys.argv[2])
    
    workbook = xlwt.Workbook(encoding='utf-8')
    sheet = SheetBaseClass(workbook, sys.argv[2])
    head = ['envTemp', 'freq', 'core', 'temp']
    sheet.addSheetHead(head);
    content = ['env15', '1536000', '4', '35']
    sheet.addSheetRow(content);
    sheet.addSheetRow(content);
    sheet.addSheetRow(content);
    sheet.addSheetHead(head);
    #sheet.addSheetRow(content);
    #sheet.addSheetRow(content);
    workbook.save(sys.argv[1])
