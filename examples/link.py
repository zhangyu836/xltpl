# -*- coding: utf-8 -*-
import os
from xltpl.writer import BookWriter
from xltpl.writerx import BookWriter as BookWriterx

pth = os.path.dirname(__file__)


def set_hyperlink(url, _, cell_state):
    target_cell = cell_state.target_cell
    target_cell.hyperlink = url


def set_hyperlink1(url, cell_state):
    target_cell = cell_state.target_cell
    target_cell.value = 'hello ' + cell_state.value
    target_cell.hyperlink = url


def set_hyperlink2(cell_state):
    target_cell = cell_state.target_cell
    target_cell.value = 'hello ' + cell_state.value
    target_cell.hyperlink = 'url'


def write_test(writer_cls, tpl_fname, result_fname):
    fname = os.path.join(pth, tpl_fname)
    writer = writer_cls(fname)
    person = {'address': u'NoWhere', 'name': u'NoName',
              'fm': 333, 'url': u'http://www.bing.com'}
    person['link'] = set_hyperlink
    person['link1'] = set_hyperlink1
    person['link2'] = set_hyperlink2
    person['sheet_name'] = 'cn'
    writer.render_sheet(person)
    writer.save(result_fname)


if __name__ == "__main__":
    tpl_name = 'link.xls'
    result_fname = 'link_00.xls'
    writer_cls = BookWriter
    write_test(writer_cls, tpl_name, result_fname)

    tpl_name = 'link.xlsx'
    result_fname = 'link_00.xlsx'
    writer_cls = BookWriterx
    write_test(writer_cls, tpl_name, result_fname)
