# -*- coding: utf-8 -*-
import os
from datetime import datetime
from xltpl.writer import BookWriter
from xltpl.writerx import BookWriter as BookWriterx

pth = os.path.dirname(__file__)


def write_test(writer_cls, tpl_fname, result_fname):
    fname = os.path.join(pth, tpl_fname)
    writer = writer_cls(fname)
    now = datetime.now()
    items = get_items()

    person = {'address': u'No Where', 'name': u'No Name',
               'fm': 333, 'date': now}
    person['rows'] = items
    person['items'] = items
    person['sheet_name'] = 'box'
    top_left = None
    for i in range(3):
        person['tpl_name'] = 'top'
        top_box = writer.render_sheet(person, top_left)

        person['tpl_name'] = 'left'
        left_box = writer.render_sheet(person, (top_box.bottom, top_box.left))

        person['tpl_name'] = 'list1'
        middle_box = writer.render_sheet(person, (left_box.top, left_box.right))

        person['tpl_name'] = 'right'
        right_box = writer.render_sheet(person, (middle_box.top, middle_box.right))

        _left = top_box.left
        _top = max(left_box.bottom, middle_box.bottom, right_box.bottom)
        top_left = (_top, _left)


    fname = os.path.join(pth, result_fname)
    writer.save(fname)


class Item():

    def __init__(self, name, category, price, count):
        self.name = name
        self.category =category
        self.price = price
        self.count = count
        self.date = datetime.now()

def get_items():
    items = []
    item = Item("萝卜", "蔬菜", 1.11, 5)
    items.append(item)
    item = Item("苹果", "水果", 2.22, 4)
    items.append(item)
    item = Item("香蕉", "水果", 3.33, 3)
    items.append(item)
    item = Item("白菜", "蔬菜", 1.11, 2)
    items.append(item)
    item = Item("白菜", "蔬菜", 1.11, 2)
    items.append(item)
    return items


if __name__ == "__main__":
    tpl_name = 'box.xls'
    result_fname = 'box_result.xls'
    writer_cls = BookWriter
    write_test(writer_cls, tpl_name, result_fname)

    tpl_name = 'box.xlsx'
    result_fname = 'box_result.xlsx'
    writer_cls = BookWriterx
    write_test(writer_cls, tpl_name, result_fname)
