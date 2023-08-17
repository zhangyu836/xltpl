# -*- coding: utf-8 -*-
import os
import copy
from datetime import datetime
from jinja2 import pass_environment
from xltpl.filters import add_filter
from xltpl.writerx import BookWriter as BookWriterx

@pass_environment
@add_filter
def link(context, url, tip=None):
    cell = context.target_cell
    cell.hyperlink = url
    if tip:
        cell.hyperlink.tooltip = tip

@pass_environment
@add_filter
def fontColor(context, color):
    cell = context.target_cell
    font = copy.copy(cell.font)
    font.color = color
    cell.font = font

@pass_environment
@add_filter
def cellColor(context, color):
    cell = context.target_cell
    fill = copy.copy(cell.fill)
    fill.fill_type = 'solid'
    fill.fgColor = color
    cell.fill = fill


def write_test(tpl_fname, result_fname):
    pth = os.path.dirname(__file__)
    fname = os.path.join(pth, tpl_fname)
    writer = BookWriterx(fname)
    writer.add_filter('link', link)
    writer.add_filter('fontColor', fontColor)
    writer.add_filter('cellColor', cellColor)
    now = datetime.now()
    person = {'address': u'NoWhere', 'name': u'NoName', 'date': now,
              'fm': 333, 'url': u'http://www.bing.com',
              'red': 'ff0000', 'blue': '70a9e1'}
    person['tpl_name'] = 'en'
    person['rows'] = [1,2,3,4,5,6,7,8,9]
    person['sheet_name'] = 'aa'
    writer.render_sheet(person)
    writer.save(result_fname)



if __name__ == "__main__":
    tpl_name = 'filter.xlsx'
    result_fname = 'filter_00.xlsx'
    write_test( tpl_name, result_fname)
