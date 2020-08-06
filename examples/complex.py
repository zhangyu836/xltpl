# -*- coding: utf-8 -*-

import os
from xltpl.writer import BookWriter
from xltpl.writerx import BookWriter as BookWriterx


def write_test(writer_cls, tpl_fname, result_fname, debug=False, nocache=False):
    pth = os.path.dirname(__file__)
    fname = os.path.join(pth, tpl_fname)
    writer = writer_cls(fname, debug=debug, nocache=nocache)

    rows = [
        ['1800909831', [
            [u'TTT00', ['aaa', 'bbb', 'ccc', 'ccc', 'ccc'], ['aaa']],
            [u'TTT01', ['aaa', 'bbb', 'ccc', 'ccc'], ['aaa', 'ccc']],
            [u'TTT02', ],
            [u'TTT03', ['aaa', 'bbb'], ['aaa', 'bbb', 'ccc', 'ccc']],
            [u'TTT04', ['aaa']]],
         ],
        ['1900012985', [
            [u'TTT01', ['aaa', 'bbb', 'ccc', 'ccc'], ['aaa', 'ccc']],
            [u'TTT02', ['aaa', 'bbb', 'ccc'], ['aaa', 'bbb']],],
         ],

        ['1900012987', [
            [u'TTT04', ['aaa'], ['kkk']] ],
         ],
    ]

    data_dict = {}
    data_dict['rows'] = rows

    payload_t2 = {'tpl_name': 'replace1', 'sheet_name': u'replace1', 'ctx': data_dict}
    payload_t3 = {'tpl_name': 'nomain', 'sheet_name': u'nomain', 'ctx': data_dict}
    payload_t4 = {'tpl_name': 'complex', 'sheet_name': u'complex', 'ctx': data_dict}
    payload_t5 = {'tpl_name': 'replace2', 'sheet_name': u'replace2', 'ctx': data_dict}
    payload_t6 = {'tpl_name': 'noifelse', 'sheet_name': u'noifelse', 'ctx': data_dict}
    payloads = [payload_t4, payload_t2, payload_t5, payload_t6, payload_t3]

    writer.render_book2(payloads=payloads)
    fname = os.path.join(pth, result_fname)
    writer.save(fname)


if __name__ == "__main__":
    tpl_name = 'complex.xls'
    result_fname = 'complex_result.xls'
    result_fname_nocache = 'complex_result_nocache.xls'
    writer_cls = BookWriter
    write_test(writer_cls, tpl_name, result_fname, False, False)
    write_test(writer_cls, tpl_name, result_fname_nocache, False, True)

    tpl_name = 'complex.xlsx'
    result_fname = 'complex_result.xlsx'
    result_fname_nocache = 'complex_result_nocache.xlsx'
    writer_cls = BookWriterx
    write_test(writer_cls, tpl_name, result_fname, True, False)
    write_test(writer_cls, tpl_name, result_fname_nocache, False, True)
