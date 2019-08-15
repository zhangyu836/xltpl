# -*- coding: utf-8 -*-

import re
import six
from xltpl.utils import tag_test

class Node(object):

    def __init__(self):
        self._children = []
        self._root = None
        self.parent = None
        self.no = None

    def __len__(self):
        return len(self._children)

    def __getitem__(self, index):
        return self._children[index]

    def __setitem__(self, index, node):
        self._children[index] = node

    def __delitem__(self, index):
        del self._children[index]

    @property
    def root(self):
        if self._root is None:
            self._root = self.parent.root
        return self._root

    def append(self, child):
        self._children.append(child)

    def add_child(self, child):
        child.no = len(self._children)
        child.parent = self
        self.append(child)

    def children_to_tag(self):
        x = []
        for child in self._children:
            x.append(child.to_tag())
        return ''.join(x)

    def to_tag(self):
        return self.tag_tpl % (self.no, self.children_to_tag())

class Row(Node):
    tag_tpl = '{%% row %d %%}\n%s'#{%% endrow %%}
    html_tpl = '<tr>\n'

    def __init__(self, rowx, first_row=0):
        Node.__init__(self)
        self.rowx = rowx
        self.first_row = first_row

    def process_rv(self, rv):
        self.write()
        return self.to_html()

    def write(self):
        if self.rowx > self.first_row:
            self.root.next_row()
        wtrowx = self.root.rowx
        self.root.writer.row(self.rowx, wtrowx)

    def to_html(self):
        if self.rowx > self.first_row:
            self.html_tpl = '</tr>' + self.html_tpl
        return self.html_tpl

class Cell(Node):
    tag_tpl = '{%% cell %d %%}%s{%% endcell %%}\n'
    html_tpl = '<td>%s</td>\n'

    def __init__(self, rowx, colx, value, cty):
        Node.__init__(self)
        self.rowx = rowx
        self.colx = colx
        self.value = value
        self.cty = cty

    def process_rv(self, rv):
        self.write(self.value)
        return self.to_html()

    def write(self, value):
        wtrowx, wtcolx = self.root.coords()
        self.root.writer.set_cell(self.rowx, self.colx, wtrowx, wtcolx, value, self.cty)
        self.root.next_cell()

    def to_html(self):
        if self.value is None:
            value = ''
        else:
            value = self.value
        return self.html_tpl % six.text_type(value)

class EmptyCell(Cell):

    def __init__(self, rowx, colx):
        Cell.__init__(self, rowx, colx, None, None)

    def process_rv(self, rv):
        self.root.next_cell()
        return self.to_html()

class RichCell(Cell):

    def __init__(self, rowx, colx, value, cty, rich_text):
        Cell.__init__(self, rowx, colx, value, cty)
        self.rich_text = rich_text

    def process_rv(self, rv):
        self.write(self.rich_text)
        return self.to_html()

    def to_html(self):
        return self.html_tpl % self.value


class VMap():

    def __init__(self, font):
        self.font = font
        self.kvs = {}
        self.section_keys = []

    def addv(self, value):
        if isinstance(value, six.text_type):
            return value
        key = "__%d__" % len(self.kvs)
        self.kvs[key] = value
        return key

    def add_section(self, section):
        key = self.addv(section)
        self.section_keys.append(key)
        s = Section(section, key)
        self.add_child(s)

    def add_tag_section(self, text, font):
        s = TagSection(text, font, 1)
        self.add_child(s)

    def pack(self, text):
        if len(self.kvs) == 0:
            return text
        pattern = '(__\d*__)'
        p = re.compile(pattern)
        parts = p.split(text)
        rv = []
        for index, part in enumerate(parts):
            if index % 2 == 0:
                if part:
                    v = self.root.handler.rich_segment(part, self.font)
                    rv.append(v)
            else:
                value = self.kvs[part]
                if isinstance(value, list):
                    rv.extend(value)
                else:
                    rv.append(value)
        self.reset()
        return rv

    def reset(self):
        if len(self.section_keys) == 0:
            self.kvs.clear()
            return
        for k in list(self.kvs.keys()):
            if k not in self.section_keys:
                del self.kvs[k]

    def unpack(self, rich_text):
        section = []
        for text,font,segment in self.root.handler.iter(rich_text, self.font):
            if tag_test(text):
                if section:
                    self.add_section(section)
                    section = []
                self.add_tag_section(text, font)
            else:
                section.append(segment)
        if section:
            self.add_section(section)

class Section():

    def __init__(self, section, key):
        self.section = section
        self.key = key

    def to_tag(self):
        return self.key


class TagSection(VMap, Node):
    tag_tpl = '{%% sec %d %%}%s{%% endsec %%}'

    def __init__(self, text, font, level=0):
        Node.__init__(self)
        VMap.__init__(self, font)
        self.text = text
        self.level = level

    def to_tag(self):
        return self.tag_tpl % (self.no, self.text)

    def process_rv(self, rv):
        rv = self.pack(rv)
        if isinstance(rv, six.text_type) and self.level > 0:
            rv = self.root.handler.rich_segment(rv, self.font)
        if not isinstance(rv, six.text_type):
            rv = self.parent.addv(rv)
        return rv

class TagCell(Cell):

    def __init__(self, rowx, colx, value, cty, font):
        Cell.__init__(self, rowx, colx, value, cty)
        self.font = font
        self.rv = None

    def addv(self, rv):
        self.rv = rv
        return ''

    def process_rv(self, rv):
        if self.rv:
            rv = self.root.handler.rich_content(self.rv)
            self.rv = None
        self.write(rv)
        return self.to_html(rv)

    def to_tag(self):
        tag_section = TagSection(self.value, self.font)
        self.add_child(tag_section)
        return Cell.to_tag(self)

    def to_html(self, value=None):
        if not value:
            return self.html_tpl % six.text_type(self.value)
        else:
            return self.html_tpl % self.root.handler.text_content(value)

class RichText(VMap):

    def __init__(self, rich_text, font):
        VMap.__init__(self, font)
        self.rich_text = rich_text


class RichTagCell(Cell, RichText):

    def __init__(self, rowx, colx, value, cty, rich_text, font):
        Cell.__init__(self, rowx, colx, value, cty)
        RichText.__init__(self, rich_text, font)

    def process_rv(self, rv):
        rv = self.pack(rv)
        rv = self.root.handler.rich_content(rv)
        self.write(rv)
        return self.to_html(rv)

    def to_tag(self):
        self.unpack(self.rich_text)
        return Cell.to_tag(self)

    def to_html(self, value=None):
        if not value:
            value = self.rich_text
        return self.html_tpl % self.root.handler.text_content(value)


class Sheet(Node):

    def __init__(self, writer, first_row=0, first_col=0):
        Node.__init__(self)
        self.writer = writer
        self.first_row = first_row
        self.first_col = first_col
        self.reset()
        self._root = self

    def get_row(self, no):
        row = self[no]
        if not isinstance(row, Row):
            raise Exception('not a row: %d' % no)
        self.current_row = row
        return row

    def get_cell(self, no):
        #cell = self.current_row[no]
        cell = self[no]
        if not isinstance(cell, Cell):
            raise Exception('not a cell: %d' % no)
        self.current_cell = cell
        return cell

    def get_section(self, no):
        section = self.current_cell[no]
        if not isinstance(section, TagSection):
            raise Exception('not a tag section: %d' % no)
        self.current_section = section
        return section

    def next_cell(self):
        self.colx += 1

    def next_row(self):
        self.rowx += 1
        self.colx = self.first_col

    def coords(self):
        return self.rowx, self.colx

    def reset(self):
        self.rowx = self.first_row
        self.colx = self.first_col

    def to_tag(self):
        return self.children_to_tag()

