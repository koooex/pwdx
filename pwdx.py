#
#============================================================
#     File: pwdx.py
#   Author: Sam [kooex@gmail.com]
#  Version: 0.1
#  Changed: 08/20/2011 17:49:11
#  History: 
#============================================================
#
# -*- coding: utf-8 -*-
#
from datetime import date, timedelta
from koogem import Object, SpringApp

class Executor(Object):
    def __init__(self, name, db, flag):
        self.db = db
        self.name = name
        self.flag = flag

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

        print '-' * (79 - len(self.name)), self.name
        print '%12s, %12s, %12s, %12s, %12s' % ('DATE', 'TOTAL', 'VALID', 'DURATION', 'FEE')
        for row in self.db.engine.text(sql).execute(start = begin, stop = end, format = "%Y-%m-%d", flag = self.flag):
            print '%12s, %12s, %12s, %12s, %12s' % (row[0], row[1], row[2], row[3], row[4])

class Stopper(object):
    def start(self):
        print '-' * 80

if __name__ == '__main__':
    SpringApp.main(block = 0)


