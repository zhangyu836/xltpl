# -*- coding: utf-8 -*-

import re
import six
from .utils import tag_test
from .utils import TreeProperty
from .pos import Pos


class Node(object):

    rich_handler = TreeProperty('rich_handler')
    node_map = TreeProperty('node_map')

    def __init__(self):
        self._children = []
        self.parent = None
        self.no = None

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
        self.node_map[self.tpl_key] = self
        return self.tag_tpl % (self.tpl_key, self.children_to_tag())

    def process_xv(self, xv):
        self.xv = xv
        return six.text_type(xv)


class Row(Node):
    tag_tpl = "{%% row '%s' %%}\n%s"#{%% endrow %%}
    html_tpl = '</tr><tr>\n'

    def __init__(self, rowx):
        Node.__init__(self)
        self.rowx = rowx

    @property
    def tpl_key(self):
        return str(self.rowx)

    def process_rv(self, rv, sheet_pos):
        wtrowx = sheet_pos.next_row()
        return self.to_html()

    def to_html(self):
        return self.html_tpl

class Cell(Node):
    tag_tpl = "{%% cell '%s' %%}%s{%% endcell %%}"
    html_tpl = '<td>%s</td>\n'

    def __init__(self, rowx, colx, value, cty):
        Node.__init__(self)
        self.rowx = rowx
        self.colx = colx
        self.value = value
        self.cty = cty

    @property
    def tpl_key(self):
        return '%d,%d' % (self.rowx, self.colx)

    def process_rv(self, rv, sheet_pos):
        self.write(self.value, self.cty, sheet_pos)
        return self.to_html()

    def write(self, value, cty, sheet_pos):
        sheet_pos.write_cell(self.rowx, self.colx, value, cty)

    def to_html(self):
        if self.value is None:
            value = ''
        else:
            value = self.value
        return self.html_tpl % six.text_type(value)

class EmptyCell(Cell):

    def __init__(self, rowx, colx):
        Cell.__init__(self, rowx, colx, None, None)

    def process_rv(self, rv, sheet_pos):
        self.write(None, None, sheet_pos)
        #sheet_pos.next_cell()
        return self.to_html()

class RichCell(Cell):

    def __init__(self, rowx, colx, value, cty, rich_text):
        Cell.__init__(self, rowx, colx, value, cty)
        self.rich_text = rich_text

    def process_rv(self, rv, sheet_pos):
        self.write(self.rich_text, self.cty, sheet_pos)
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
                    v = self.rich_handler.rich_segment(part, self.font)
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
        for text,font,segment in self.rich_handler.iter(rich_text, self.font):
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
    tag_tpl = "{%% sec '%s' %%}%s{%% endsec %%}"

    def __init__(self, text, font, level=0):
        Node.__init__(self)
        VMap.__init__(self, font)
        self.text = text
        self.level = level

    @property
    def tpl_key(self):
        return '%s,%d' % (self.parent.tpl_key, self.no)

    def to_tag(self):
        self.node_map[self.tpl_key] = self
        return self.tag_tpl % (self.tpl_key, self.text)

    def process_rv(self, rv, sheet_pos):
        rv = self.pack(rv)
        if isinstance(rv, six.text_type) and self.level > 0:
            rv = self.rich_handler.rich_segment(rv, self.font)
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

    def process_rv(self, rv, sheet_pos):
        if self.rv:
            rv = self.rich_handler.rich_content(self.rv)
            self.rv = None
        self.write(rv, self.cty, sheet_pos)
        return self.to_html(rv)

    def to_tag(self):
        tag_section = TagSection(self.value, self.font)
        self.add_child(tag_section)
        return Cell.to_tag(self)

    def to_html(self, value=None):
        if not value:
            return self.html_tpl % six.text_type(self.value)
        else:
            return self.html_tpl % self.rich_handler.text_content(value)

class XvCell(Cell):

    def __init__(self, rowx, colx, value, cty):
        Cell.__init__(self, rowx, colx, value, cty)
        self.xv = None

    def process_rv(self, rv, sheet_pos):
        if self.xv:
            self.write(self.xv, None, sheet_pos)
        else:
            self.write(rv, None, sheet_pos)
        return self.to_html(rv)

    def to_tag(self):
        self.node_map[self.tpl_key] = self
        return self.tag_tpl % (self.tpl_key, self.value)

    def to_html(self, value=None):
        return self.html_tpl % value

class RichText(VMap):

    def __init__(self, rich_text, font):
        VMap.__init__(self, font)
        self.rich_text = rich_text


class RichTagCell(Cell, RichText):

    def __init__(self, rowx, colx, value, cty, rich_text, font):
        Cell.__init__(self, rowx, colx, value, cty)
        RichText.__init__(self, rich_text, font)

    def process_rv(self, rv, sheet_pos):
        rv = self.pack(rv)
        rv = self.rich_handler.rich_content(rv)
        self.write(rv, self.cty, sheet_pos)
        return self.to_html(rv)

    def to_tag(self):
        self.unpack(self.rich_text)
        return Cell.to_tag(self)

    def to_html(self, value=None):
        if not value:
            value = self.rich_text
        return self.html_tpl % self.rich_handler.text_content(value)


