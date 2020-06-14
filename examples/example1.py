# -*- coding: utf-8 -*-

import os
from datetime import datetime
from xltpl.writer import BookWriter
from xltpl.writerx import BookWriter as BookWriterx


def write_test(writer_cls, tpl_fname, result_fname0, result_fname1):
    pth = os.path.dirname(__file__)
    fname = os.path.join(pth, tpl_fname)
    writer = writer_cls(fname)
    writer.jinja_env.globals.update(dir=dir, getattr=getattr)

    now = datetime.now()

    person_info = {'address': u'福建行中书省福宁州傲龙山庄', 'name': u'龙傲天', 'fm': 178, 'date': now}
    person_info2 = {'address': u'Somewhere over the rainbow', 'name': u'Hello Wizard', 'fm': 156, 'date': now}
    rows = [['1', '1', '1', '1', '1', '1', '1', '1', ],
             ['1', '1', '1', '1', '1', '1', '1', '1', ],
             ['1', '1', '1', '1', '1', '1', '1', '1', ],
             ['1', '1', '1', '1', '1', '1', '1', '1', ],
             ['1', '1', '1', '1', '1', '1', '1', '1', ],
             ['1', '1', '1', '1', '1', '1', '1', '1', ],
             ['1', '1', '1', '1', '1', '1', '1', '1', ],
             ['1', '1', '1', '1', '1', '1', '1', '1', ],
            ]
    person_info['rows'] = rows
    person_info2['rows'] = rows
    payload0 = {'tpl_name': 'cn', 'sheet_name': u'表',  'ctx': person_info}
    payload1 = {'tpl_name': 'en', 'sheet_name': u'form', 'ctx': person_info2}
    payload2 = {'tpl_idx': 2, 'ctx': person_info2}
    payloads = [payload0, payload1, payload2]
    writer.render_book2(payloads=payloads)
    fname = os.path.join(pth, result_fname0)
    writer.save(fname)
    payloads = [payload2, payload1, payload0]
    writer.render_book2(payloads=payloads)
    writer.render_sheet(person_info2, 'form2', 1)
    fname = os.path.join(pth, result_fname1)
    writer.save(fname)


if __name__ == "__main__":
    tpl_name = 'example.xls'
    result_fname0 = 'result10.xls'
    result_fname1 = 'result11.xls'
    writer_cls = BookWriter
    write_test(writer_cls, tpl_name, result_fname0, result_fname1)

    tpl_name = 'example.xlsx'
    result_fname0 = 'result10.xlsx'
    result_fname1 = 'result11.xlsx'
    writer_cls = BookWriterx
    write_test(writer_cls, tpl_name, result_fname0, result_fname1)
