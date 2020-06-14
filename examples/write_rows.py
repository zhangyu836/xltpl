# -*- coding: utf-8 -*-

import os
from xltpl.writery import BookWriter, BookWriterx
from collections import OrderedDict

rows = [['1', '1', '1', '1', '1', '1', '1', '1', '1', '1', '1'],
            ['1', '1', '1', '1', '1', '1', '1', '1', ],
            ['1', '1', '1', '1', '1', '1', '1', '1', ],
            ['1', '1', '1', '1', '1', '1', '1', '1', ],
             [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1,],
             [1, 1, 1, 1, 1, 1, 1, 1, ],
             [1, 1, 1, 1, 1, 1, 1, 1, ],
             [1, 1, 1, 1, 1, 1, 1, 1, ],
             [1, 1, 1, 1, 1],
            ]

def write_test(writer_cls, tpl_fname, result_fname):
    pth = os.path.dirname(__file__)
    fname = os.path.join(pth, tpl_fname)
    writer = writer_cls(fname, body_rows=[4])

    writer.write_list(rows)
    writer.write_list(rows)
    writer.write_list(rows, '1')
    writer.write_list(rows, '1')
    sheet_writer = writer.get_sheet_writer('1')
    sheet_writer.write_rows(rows)
    sheet_writer = writer.get_sheet_writer('2')
    sheet_writer.copy_head()
    sheet_writer.write_rows(rows)
    sheet_writer.write_rows(rows)
    sheet_writer.write_rows(rows)
    sheet_writer.copy_foot()

    d = OrderedDict()
    d['1'] = rows
    d['2'] = rows
    d['3'] = rows
    d['8'] = rows
    d['7'] = rows
    d['6'] = rows
    writer.write_dict(d)

    payload0 = {'sheet_name': 'payload0', 'data': rows}
    payload1 = {'sheet_name': 'payload1', 'data': rows}
    payloads = [payload0, payload1]
    writer.write_payloads(payloads)

    fname = os.path.join(pth, result_fname)
    writer.save(fname)


if __name__ == "__main__":
    tpl_fname = 'write_rows.xls'
    result_fname = 'write_rows_result.xls'
    writer_cls = BookWriter
    write_test(writer_cls, tpl_fname, result_fname)
    tpl_fname = 'write_rows.xlsx'
    result_fname = 'write_rows_result.xlsx'
    writer_cls = BookWriterx
    write_test(writer_cls, tpl_fname, result_fname)
