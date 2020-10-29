# -*- coding: utf-8 -*-

import os
from xltpl.writerx import BookWriter as BookWriterx


def write_test(writer_cls, tpl_fname, result_fname):
    pth = os.path.dirname(__file__)
    fname = os.path.join(pth, tpl_fname)
    writer = writer_cls(fname, debug=True)

    rows = [
        ['1800909831', [
            [u'TTT00', ['D017C2870C24', '3C77E6F32EF4', '5056856397', '5056856397', '5056856397'], ['D017C2870C24']],
            [u'TTT01', ['D017C2870C24', '3C77E6F32EF4', '5056856397', '5056856397'], ['D017C2870C24', '5056856397']],
            [u'TTT02', ['D017C2870C24', '3C77E6F32EF4', '5056856397'], ['D017C2870C24', '3C77E6F32EF4']],
            [u'TTT03', ['D017C2870C24', '3C77E6F32EF4'], ['D017C2870C24', '3C77E6F32EF4', '5056856397', '5056856397']],
            [u'TTT04', ['D017C2870C24'], ['D017C2870C24', '3C77E6F32EF4', '5056856397', '5056856397', '5056856397']]],
         ],
        ['1900012985', [
            [u'TTT01', ['D017C2870C24', '3C77E6F32EF4', '5056856397', '5056856397'], ['D017C2870C24', '5056856397']],
            [u'TTT02', ['D017C2870C24', '3C77E6F32EF4', '5056856397'], ['D017C2870C24', '3C77E6F32EF4']],
            [u'TTT03', ['D017C2870C24', '3C77E6F32EF4'], ['D017C2870C24', '3C77E6F32EF4', '5056856397', '5056856397']],
            [u'TTT04', ['D017C2870C24'], ['D017C2870C24', '3C77E6F32EF4', '5056856397', '5056856397', '5056856397']]],
         ],
        ['1900012986', [
            [u'TTT02', ['D017C2870C24', '3C77E6F32EF4', '5056856397'], ['D017C2870C24', '3C77E6F32EF4']],
            [u'TTT03', ['D017C2870C24', '3C77E6F32EF4'], ['D017C2870C24', '3C77E6F32EF4', '5056856397', '5056856397']],
            [u'TTT04', ['D017C2870C24'], ['D017C2870C24', '3C77E6F32EF4', '5056856397', '5056856397', '5056856397']]],
         ],
        ['1900012987', [
            [u'TTT03', ['D017C2870C24', '3C77E6F32EF4'], ['D017C2870C24', '3C77E6F32EF4', '5056856397', '5056856397']],
            [u'TTT04', ['D017C2870C24'], ['D017C2870C24', '3C77E6F32EF4', '5056856397', '5056856397', '5056856397']]],
         ],
        ['1900012987', [
            [u'TTT04', ['D017C2870C24'], ['D017C2870C24', '3C77E6F32EF4', '5056856397', '5056856397', '5056856397']]],
         ]
    ]

    data_dict = {}
    data_dict['rows'] = rows

    payload_t2 = {'tpl_name': '2', 'sheet_name': u'sheet2', 'ctx': data_dict}
    payload_t3 = {'tpl_name': '4', 'sheet_name': u'sheet4', 'ctx': data_dict}
    payload_t4 = {'tpl_name': '1', 'sheet_name': u'sheet1', 'ctx': data_dict}
    payload_t5 = {'tpl_name': '3', 'sheet_name': u'sheet3', 'ctx': data_dict}
    payload_t6 = {'tpl_name': '1p', 'sheet_name': u'sheet1p', 'ctx': data_dict}
    payload_t7 = {'tpl_name': '2p', 'sheet_name': u'sheet2p', 'ctx': data_dict}
    payload_t8 = {'tpl_name': '3p', 'sheet_name': u'sheet3p', 'ctx': data_dict}
    payload_t9 = {'tpl_name': '4p', 'sheet_name': u'sheet4p', 'ctx': data_dict}
    payloads = [payload_t4, payload_t6, payload_t2, payload_t7, payload_t5, payload_t8, payload_t3, payload_t9]

    writer.render_book2(payloads=payloads)
    fname = os.path.join(pth, result_fname)
    writer.save(fname)


if __name__ == "__main__":
    tpl_name = 'ranges_dv.xlsx'
    result_fname = 'ranges_dv_result.xlsx'
    writer_cls = BookWriterx
    write_test(writer_cls, tpl_name, result_fname)
