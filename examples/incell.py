# -*- coding: utf-8 -*-

import os
from datetime import datetime
from xltpl.writer import BookWriter
from xltpl.writerx import BookWriter as BookWriterx

try:
    from PIL import Image as PILImage
except ImportError:
    raise ImportError('Pillow is required for fetching image objects')


def write_test(writer_cls, tpl_fname, result_fname):
    pth = os.path.dirname(__file__)
    fname = os.path.join(pth, tpl_fname)
    writer = writer_cls(fname)
    now = datetime.now()
    img0 = os.path.join(pth, 'images/0.jpg')
    img1 = os.path.join(pth, 'images/1.jpg')
    img1 = PILImage.open(img1)
    person_info0 = {'address': u'Somewhere over the rainbow', 'name': u'Hello Wizard', 'date': now, 'img': img0}
    person_info1 = {'address': u'Nowhere', 'name': u'No name'}
    person_info2 = {'address': u'福建行中书省福宁州傲龙山庄', 'name': u'龙傲天', 'img': img1}
    persons = {'ps': [person_info0, person_info1, person_info2], 'tpl_name': 'images'}
    rows = [['1'], ['1'], [1], [1], [1],]
    person_info0['rows'] = rows
    person_info0['tpl_idx'] = 0
    person_info0['sheet_name'] = 'form'
    person_info2['rows'] = rows
    person_info2['tpl_name'] = 'loops'
    person_info2['sheet_name'] = 'loops'
    payloads = [person_info0, persons, person_info2, persons]#
    writer.render_book(payloads=payloads)
    fname = os.path.join(pth, result_fname)
    writer.save(fname)


if __name__ == "__main__":
    tpl_name = 'incell.xlsx'
    result_fname = 'incell_result.xlsx'
    writer_cls = BookWriterx
    write_test(writer_cls, tpl_name, result_fname)

    tpl_name = 'incell.xls'
    result_fname = 'incell_result.xls'
    writer_cls = BookWriter
    write_test(writer_cls, tpl_name, result_fname)

