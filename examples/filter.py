# -*- coding: utf-8 -*-
import os
import copy
from datetime import datetime
import xlwt
from jinja2 import pass_environment
from xltpl.filters import add_filter
from xltpl.writer import BookWriter

@pass_environment
@add_filter
def bold(cell_context, bold=None):
    style = cell_context.get_style()
    if bold:
        font = copy.copy(style.font)
        font.bold = True
        style.font = font
        pattern = copy.copy(style.pattern)
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = xlwt.Style.colour_map['pale_blue']
        style.pattern = pattern


def write_test(tpl_fname, result_fname):
    pth = os.path.dirname(__file__)
    fname = os.path.join(pth, tpl_fname)
    writer = BookWriter(fname)
    writer.add_filter('bold', bold)
    writer.add_global('bold', bold)
    now = datetime.now()
    person = {'address': u'NoWhere', 'name': u'NoName', 'date': now,
              'fm': 333, 'url': u'http://www.bing.com'}
    person['tpl_name'] = 'en'
    person['rows'] = [1,2,3,4,5,6,7,8,9]
    writer.render_sheet(person)
    writer.save(result_fname)


if __name__ == "__main__":
    tpl_name = 'filter.xls'
    result_fname = 'filter_00.xls'
    write_test(tpl_name, result_fname)

