# -*- coding: utf-8 -*-
import os
from datetime import datetime
from xltpl.writer import BookWriter
from xltpl.writerx import BookWriter as BookWriterx

pth = os.path.dirname(__file__)
try:
    from PIL import Image as PILImage
    img0 = os.path.join(pth, 'images/0.jpg')
    img1 = os.path.join(pth, 'images/1.jpg')
    img1 = PILImage.open(img1)
except ImportError:
    print('Pillow is required for fetching image objects')
    img0 = None
    img1 = None

def write_test(writer_cls, tpl_fname, result_fname0, result_fname1):
    fname = os.path.join(pth, tpl_fname)
    writer = writer_cls(fname)
    writer.set_jinja_globals(dir=dir, getattr=getattr)
    now = datetime.now()
    items = get_items()
    person0 = {'address': u'福建行中书省福宁州傲龙山庄', 'name': u'龙傲天',
               'fm': 178, 'date': now, 'img':img1}
    person1 = {'address': u'Somewhere over the rainbow', 'name': u'Hello Wizard',
               'fm': 156, 'date': now, 'img': img0}
    person2 = {'address': u'No Where', 'name': u'No Name',
               'fm': 333, 'date': now}
    person0['rows'] = items
    person1['rows'] = items
    person2['rows'] = items
    person0['items'] = items
    person1['items'] = items
    person2['items'] = items
    person0['sheet_name'] = 'cn'
    person1['sheet_name'] = 'en'
    #person0['tpl_index'] = 0 default 0
    person1['tpl_index'] = 1
    payloads = [person0, person1]
    writer.render_sheets(payloads=payloads)
    person0['sheet_name'] = 'en'
    person1['sheet_name'] = 'cn'
    writer.render_sheets(payloads=payloads)#append
    person2['tpl_name'] = 'en'
    person2['sheet_name'] = 'hello sheet'
    writer.render_sheet(person2)
    person2['tpl_name'] = 'list0'
    person2['sheet_name'] = 'list0'
    time0 = datetime.now()
    for i in range(1111):
        writer.render_sheet(person2)
    time1 = datetime.now()
    print('list0:', time1 - time0)
    person2['tpl_name'] = 'list1'
    person2['sheet_name'] = 'list1'
    for i in range(1111):
        writer.render_sheet(person2)
    time2 = datetime.now()
    print('list1:', time2 - time1)
    persons = {'ps': [person1, person2, person0, person0], 'tpl_name': 'image', 'sheet_name': 'image'}
    writer.render_sheet(persons)
    person0['sheet_name'] = ''# no sheet name
    person1['sheet_name'] = ''
    writer.render_sheets(payloads=payloads)
    fname = os.path.join(pth, result_fname0)
    writer.save(fname)
    payloads = [person0, person1, person2, persons]
    writer.render_sheets(payloads)
    for p in payloads:
        p['sheet_name'] = 'all_in_one'
    writer.render_sheets(payloads)
    fname = os.path.join(pth, result_fname1)
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
    tpl_name = 'example.xls'
    result_fname0 = 'result00.xls'
    result_fname1 = 'result01.xls'
    writer_cls = BookWriter
    write_test(writer_cls, tpl_name, result_fname0, result_fname1)

    tpl_name = 'example.xlsx'
    result_fname0 = 'result00.xlsx'
    result_fname1 = 'result01.xlsx'
    writer_cls = BookWriterx
    write_test(writer_cls, tpl_name, result_fname0, result_fname1)
