# -*- coding: utf-8 -*-

import os
import copy
from datetime import datetime
from xltpl.writerx import BookWriter


def write_test():
    pth = os.path.dirname(__file__)
    fname = os.path.join(pth, 'example.xlsx')
    writer = BookWriter(fname)
    writer.jinja_env.globals.update(dir=dir, getattr=getattr)

    now = datetime.now()

    person_info = {'address': u'福建行中书省福宁州傲龙山庄', 'name': u'龙傲天', 'fm': 178, 'date': now}
    person_info2 = {'address': u'Somewhere over the rainbow', 'name': u'Hello Wizard', 'fm': 156, 'date': now}
    rows = [['1', '1', '1', '1', '1', '1', '1', '1', ],
            ['1', '1', '1', '1', '1', '1', '1', '1', ],
            ['1', '1', '1', '1', '1', '1', '1', '1', ],
            ['1', '1', '1', '1', '1', '1', '1', '1', ],
             [1, 1, 1, 1, 1, 1, 1, 1, ],
             [1, 1, 1, 1, 1, 1, 1, 1, ],
             [1, 1, 1, 1, 1, 1, 1, 1, ],
             [1, 1, 1, 1, 1, 1, 1, 1, ],
             [1, 1, 1, 1, 1],
            ]
    person_info['rows'] = rows
    person_info['tpl_idx'] = 0
    person_info2['rows'] = rows
    person_info2['tpl_name'] = 'en'
    person_info3 = copy.copy(person_info2)
    person_info3['sheet_name'] = 'hello sheet'
    person_info4 = copy.copy(person_info2)
    person_info4['tpl_name'] = 'ex'
    person_info4['sheet_name'] = 'cols'
    payloads = [person_info, person_info2, person_info3, person_info4]
    writer.render_book(payloads=payloads)
    fname = os.path.join(pth, 'result00.xlsx')
    writer.save(fname)
    payloads = [person_info3, person_info, person_info2]
    writer.render_book(payloads=payloads)
    writer.render_sheet(person_info2, 'form2', 1)
    fname = os.path.join(pth, 'result01.xlsx')
    writer.save(fname)


if __name__ == "__main__":
    write_test()
