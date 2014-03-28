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

        sheet.col(c + 0).width = 256 * 16
        sheet.col(c + 1).width = 256 * 12
        sheet.col(c + 2).width = 256 * 12
        sheet.col(c + 3).width = 256 * 12
        sheet.col(c + 4).width = 256 * 12

        style = u'''alignment: horizontal right;
                    font: name 微软雅黑, color white;
                    pattern: pattern solid, fore_color aqua;'''
        style0 = easyxf(style + u'borders: top medium, bottom thin, left medium')
        style5 = easyxf(style + u'borders: top medium, bottom thin;')
        style9 = easyxf(style + u'borders: top medium, bottom thin, right medium;')
        sheet.write(r, c + 0, u'日期', style0)
        sheet.write(r, c + 1, u'过网次数', style5)
        sheet.write(r, c + 2, u'拦截次数', style5)
        sheet.write(r, c + 3, u'时长', style5)
        sheet.write(r, c + 4, u'费用', style9)

        style = u'''alignment: horizontal right;
                    font: name Tahoma, height 160;'''
        style0 = easyxf(style + u'borders: left medium;',
                    num_format_str = 'M/D/YYYY')
        style5 = easyxf(style)
        style9 = easyxf(style + u'borders: right medium;',
                   #num_format_str = u"[$$-409]#,##0.00;-[$$-409]#,##0.00")
                    num_format_str = u"[$¥-804]#,##0.00;-[$¥-804]#,##0.00")
        for i, db in enumerate(results):
            r += 1
            sheet.write(r, c + 0, datetime.strptime(db[0], '%Y-%m-%d'), style0)
            sheet.write(r, c + 1, db[1], style5)
            sheet.write(r, c + 2, db[2], style5)
            sheet.write(r, c + 3, db[3], style5)
            sheet.write(r, c + 4, db[4], style9)

        if r > anchor:
            cols = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

            style = 'font: bold True; borders: top medium'
            style0 = easyxf(style)
            style5 = easyxf(style)
            style9 = easyxf(style, num_format_str = u"[$¥-804]#,##0.00;-[$¥-804]#,##0.00")
            m = anchor + 2
            n = r + 1
            r += 1
            sheet.row(r).height = int(256 * 1.5)
            sheet.row(r).height_mismatch = 1
            sheet.write(r, c + 0, 'TOTAL', style0)
            sheet.write(r, c + 1, Formula('SUM({0}{1}:{0}{2})'.format(cols[c + 1], m, n)), style5)
            sheet.write(r, c + 2, Formula('SUM({0}{1}:{0}{2})'.format(cols[c + 2], m, n)), style5)
            sheet.write(r, c + 3, Formula('SUM({0}{1}:{0}{2})'.format(cols[c + 3], m, n)), style5)
            sheet.write(r, c + 4, Formula('SUM({0}{1}:{0}{2})'.format(cols[c + 4], m, n)), style9)

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


