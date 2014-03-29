# -*- coding: utf-8 -*-
#
#============================================================
#     File: xls.py
#   Author: Sam [kooex@gmail.com]
#  Version: 0.1
#  Changed: 08/20/2011 17:49:11
#  History:
#============================================================
#
#
from datetime import datetime, date, timedelta
from koogem import Object
from xlwt import easyxf, Workbook, Formula
from xlwt.Utils import *

class Executor(Object):
    book = Workbook()

    def __init__(self, name, db, flag):
        self.db = db
        self.name = name
        self.flag = flag

    def save(self, results):
        r = c = anchor = 2

        sheet = self.book.add_sheet(self.name)
        
        sheet.row(r).height = int(256 * 1.5)
        sheet.row(r).height_mismatch = 1
        sheet.write_merge(r, r, c + 0, c + 4, u"网际通系统业务清单", 
                easyxf(u'''alignment: horizontal center, vertical center;
                           font: name 微软雅黑, bold True;'''))
        r += 1

        sheet.row(r).height = int(256 * 1.5)
        sheet.row(r).height_mismatch = 1

        sheet.col(c + 0).width = 256 * 16
        sheet.col(c + 1).width = 256 * 12
        sheet.col(c + 2).width = 256 * 12
        sheet.col(c + 3).width = 256 * 12
        sheet.col(c + 4).width = 256 * 12

        style = u'''alignment: horizontal right;
                    font: name 微软雅黑, color white;
                    pattern: pattern solid, fore_color aqua;'''
        style0 = easyxf(style + u'borders: top medium, bottom thin, left medium;')
        style5 = easyxf(style + u'borders: top medium, bottom thin;')
        style9 = easyxf(style + u'borders: top medium, bottom thin, right medium;')
        sheet.write(r, c + 0, u'日期', style0)
        sheet.write(r, c + 1, u'替换次数', style5)
        sheet.write(r, c + 2, u'通话次数', style5)
        sheet.write(r, c + 3, u'通话时长', style5)
        sheet.write(r, c + 4, u'金额', style9)

        style = u'''alignment: horizontal right;
                    font: name Tahoma, height 160;'''
        style0 = easyxf(style + u'borders: left medium, right thin, bottom thin;',
                    num_format_str = 'M/D/YYYY')
        style5 = easyxf(style + u'borders: right thin, bottom thin;')
        style9 = easyxf(style + u'borders: right medium, bottom thin;',
                   #num_format_str = u"[$$-409]#,##0.00;-[$$-409]#,##0.00")
                    num_format_str = u"[$¥-804]#,##0.00;-[$¥-804]#,##0.00")
        for db in results:
            r += 1
            sheet.write(r, c + 0, datetime.strptime(db[0], '%Y-%m-%d'), style0)
            sheet.write(r, c + 1, db[1], style5)
            sheet.write(r, c + 2, db[2], style5)
            sheet.write(r, c + 3, db[3], style5)
            sheet.write(r, c + 4, db[4], style9)

        if r > anchor + 1:
            style = 'font: bold True; borders: top medium'
            style0 = easyxf(style)
            style5 = easyxf(style)
            style9 = easyxf(style, num_format_str = u"[$¥-804]#,##0.00;-[$¥-804]#,##0.00")
            r += 1
            sheet.row(r).height = int(256 * 1.5)
            sheet.row(r).height_mismatch = 1
            sheet.write(r, c + 0, 'TOTAL', style0)
            sheet.write(r, c + 1, style = style5)
            sheet.write(r, c + 2, style = style5)
            sheet.write(r, c + 3, style = style5)
            c0 = rowcol_to_cell(anchor + 2, c + 4)
            c1 = rowcol_to_cell(r - 1, c + 4)
            sheet.write(r, c + 4, Formula('SUM(%s:%s)' % (c0, c1)), style9)

    def start(self):
        d = date.today()
        end = date(d.year, d.month, 1)

        d = end - timedelta(days = 1)
        begin = date(d.year, d.month, 1)

        sql = """select DATE_FORMAT(start, :format) as v__, count(*),
         sum(if(duration > 0, 1, 0)), sum(duration),
         sum(if(duration > 3, truncate(((duration + 59) / 60), 0)*0.04, 0)) from records r1
         where start >= :start and start < :stop and %s and reason = 0
         group by v__"""

        sql = sql % ('gcflag = :flag' if self.flag < 10 else '1 = 1')
        self.debug(sql)

        results = self.db.engine.text(sql).execute(
                start = begin, stop = end, format = "%Y-%m-%d", flag = self.flag)
        self.save(results)

class Stopper(object):
    def stop(self):
        Executor.book.save('%s.xls' % date.today().strftime('%Y%m%d'))


