# -*- coding: utf-8 -*-

from xltpl.writerx import XlsxWriter

def write_test():
    fname = './example.xlsx'
    writer = XlsxWriter(fname)
    writer.jinja_env.globals.update(dir=dir, getattr=getattr)

    person_info = {'address': u'福建行中书省福宁州傲龙山庄', 'name': u'龙傲天', 'fm': 1, }
    person_info2 = {'address': u'Somewhere over the rainbow', 'name': u'Hello Wizard', 'fm': 1, }
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
    writer.save('./result0.xlsx')
    payloads = [payload2, payload1, payload0]
    writer.render_book2(payloads=payloads)
    writer.render_sheet(person_info2, 'form2', 1)
    writer.save('./result1.xlsx')


if __name__ == "__main__":
    write_test()
