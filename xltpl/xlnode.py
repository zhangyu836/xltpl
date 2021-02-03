# -*- coding: utf-8 -*-

from __future__ import print_function
import six
from copy import copy
from .misc import TreeProperty
from .utils import tag_test, xv_test, find_cell_tag, block_split, rich_split, img_test

class Node(object):
    node_map = TreeProperty('node_map')
    ext_tag = 'node'

    def __init__(self):
        self._children = []
        self._parent = None
        self.no = 0

    @property
    def depth(self):
        if not hasattr(self, '_depth'):
            if not hasattr(self, '_parent') or self._parent is None or self._parent is self:
                self._depth = 0
            else:
                self._depth = self._parent.depth + 1
        return self._depth

    @property
    def node_key(self):
        return '%s,%s' % (self._parent.node_key, self.no)

    @property
    def node_tag(self):
        return "{%%%s '%s' %%}" % (self.ext_tag, self.node_key)

    @property
    def print_tag(self):
        return self.node_tag

    def children_to_tag(self):
        x = []
        for child in self._children:
            x.append(child.to_tag())
        return '\n'.join(x)

    def to_tag(self):
        if self._children:
            return self.children_to_tag()
        self.node_map[self.node_key] = self
        return self.node_tag

    def tag_tree(self):
        print('\t' * self.depth, self.print_tag)
        for child in self._children:
            child.tag_tree()

    def add_child(self, child):
        child.no = len(self._children)
        child._parent = self
        self._children.append(child)

    def enter(self):
        pass

    def reenter(self):
        self.enter()

    def child_reenter(self):
        pass

    def exit(self):
        pass

    def __str__(self):
        return self.__class__.__name__ + ' , ' + self.node_tag

    def copy(self):
        node = copy(self)
        node._children = []
        for child in self._children:
            node.add_child(child.copy())
        return node

    def set_image_ref(self, image_ref, image_key):
        self._parent.set_image_ref(image_ref, image_key)

class Segment(Node):

    def __init__(self, text):
        Node.__init__(self)
        self.text = text

    @property
    def node_tag(self):
        fmt = "{%%seg '%s'%%}%s{%%endseg%%}"
        return fmt % (self.node_key, self.text)

    def process_rv(self, rv):
        self._parent.process_child_rv(rv)
        return rv

class RichSegment(Segment):

    def __init__(self, text, font):
        Segment.__init__(self, text)
        self.font = font

    def process_rv(self, rv):
        self._parent.process_child_rv(self.rv)
        return self.text

    def process_rich_rv(self, rv):
        self.rv = rv
        return self.text

class BlockSegment(Segment):

    @property
    def node_tag(self):
        return self.text

class ImageSegment(Segment):

    @property
    def node_tag(self):
        fmt = "{%%seg '%s'%%}{%%endseg%%}%s"
        return fmt % (self.node_key, self.text)

class Section(Node):

    def __init__(self, text, font, rich_handler):
        Node.__init__(self)
        self.font = font
        self.rich_handler = rich_handler
        self.unpack(text)

    def unpack(self, text):
        parts = block_split(text)
        for index, part in enumerate(parts):
            if part == '':
                continue
            if index % 2 == 0:
                sub_parts = rich_split(part)
                for sub_index, sub_part in enumerate(sub_parts):
                    if sub_part == '':
                        continue
                    if sub_index % 2 == 0:
                        child = Segment(sub_part)
                    else:
                        if img_test(sub_part):
                            child = ImageSegment(sub_part)
                        else:
                            child = RichSegment(sub_part, self.font)
                    self.add_child(child)
            else:
                child = BlockSegment(part)
                self.add_child(child)

    def pack(self):
        if not self.richs:
            text = ''.join(self.child_rvs)
            if isinstance(self, TagCell):
                return text
            else:
                return self.rich_handler.rich_segment(text, self.font)
        richs = self.richs
        rvs = self.child_rvs
        font = self.font
        richs.append(len(rvs))
        st = 0
        rv = []
        for i in richs:
            if st == i:
                pass
            else:
                if st + 1 == i:
                    text = rvs[st]
                else:
                    slice = rvs[st:i]
                    text = ''.join(slice)
                rich_text = self.rich_handler.rich_segment(text, self.font)
                rv.append(rich_text)
            if i < len(rvs):
                rv.append(rvs[i])
            st = i + 1
        return rv

    def enter(self):
        self.child_rvs = []
        self.richs = []

    def exit(self):
        rv = self.pack()
        self._parent.process_child_rv(rv)

    def process_child_rv(self, rv):
        if not isinstance(rv, six.text_type):
            self.richs.append(len(self.child_rvs))
        self.child_rvs.append(rv)

class Cell(Node):
    ext_tag = 'cell'

    def __init__(self, rowx, colx, value, cty):
        Node.__init__(self)
        self.rowx = rowx
        self.colx = colx
        self.value = value
        self.cty = cty
        self.cell_tag = None

    def to_tag(self):
        tag = Node.to_tag(self)
        if self.cell_tag:
            beforecell = self.cell_tag.beforecell
            aftercell = self.cell_tag.aftercell
            return beforecell + tag + aftercell
        return tag

    def extend_cell_tag(self, other):
        if self.cell_tag:
            self.cell_tag.extend(other)
        else:
            self.cell_tag = other

    def exit(self):
        self.write(self.value, self.cty)

    def process_child_rv(self, rv):
        pass

    def write(self, value, cty):
        self._parent.write_cell(self.rowx, self.colx, value, cty)

    def set_image_ref(self, image_ref, image_key):
        self._parent.set_image_ref(image_ref, (self.rowx,self.colx,image_key))

class TagCell(Section, Cell):

    def __init__(self, rowx, colx, value, cty, font, rich_handler):
        Cell.__init__(self, rowx, colx, value, cty)
        self.font = font
        self.rich_handler = rich_handler
        self.unpack(value)

    def exit(self):
        rv = self.pack()
        if not isinstance(rv, six.text_type):
            rv = self.rich_handler.rich_content(rv)
        self.write(rv, self.cty)

class RichTagCell(Cell):

    def __init__(self, rowx, colx, value, cty, font, rich_handler):
        Cell.__init__(self, rowx, colx, value, cty)
        self.font = font
        self.rich_handler = rich_handler
        self.unpack(value)

    def unpack(self, rich_text):
        for text, font, segment in self.rich_handler.iter(rich_text, self.font):
            section = Section(text, font, self.rich_handler)
            self.add_child(section)

    def enter(self):
        self.child_rvs = []

    def exit(self):
        rv = self.rich_handler.rich_content(self.child_rvs)
        self.write(rv, self.cty)

    def process_child_rv(self, rv):
        if isinstance(rv, list):
            self.child_rvs.extend(rv)
        else:
            self.child_rvs.append(rv)

class Empty(object):
    pass

class EmptyCell(Cell):

    def __init__(self, rowx, colx):
        Node.__init__(self)
        self.rowx = rowx
        self.colx = colx
        self.value = ''
        self.cty = None
        self.cell_tag = None

    def exit(self):
        self.write(Empty, None)

class XvCell(Cell):

    def __init__(self, rowx, colx, value, cty):
        Cell.__init__(self, rowx, colx, value, cty)
        seg = Segment(self.value)
        self.add_child(seg)

    def enter(self):
        self.rv = None

    def exit(self):
        self.write(self.rv, None)

class ExtraCell(Cell):
    ext_tag = 'extra'

    def to_tag(self):
        tag = Node.to_tag(self)
        if self.cell_tag:
            extra = self.cell_tag.extracell
            return tag + extra
        return tag

    def exit(self):
        pass

class Row(Node):
    ext_tag = 'row'

    def __init__(self, rowx):
        Node.__init__(self)
        self.rowx = rowx
        self.cell_tag = None

    def to_tag(self):
        tag = Node.to_tag(self)
        if self.cell_tag:
            beforerow = self.cell_tag.beforerow
            return beforerow + tag
        return tag

    def enter(self):
        self._parent.next_row()

class CellRange(Node):
    ext_tag = 'crange'

    def __init__(self, xlrange):
        Node.__init__(self)
        self.sheet_pos = None
        self.xlrange = xlrange
        self.rkey = xlrange.rkey
        self._depth = xlrange.depth + 1
        self._node_map = xlrange.node_map
        self.get_children()

    @property
    def node_key(self):
        return self.rkey

    def get_children(self):
        if self.rkey == 'tail':
            cell = ExtraCell(0, 0, 0, 0)
            self.add_child(cell)
            return
        xlrange = self.xlrange
        cell_map = self.xlrange.cell_map
        cell_tag_map = self.xlrange.cell_tag_map
        cp = '*' in self.rkey
        for rowx in range(xlrange.min_rowx, xlrange.max_rowx + 1):
            row = Row(rowx)
            self.add_child(row)
            for colx in range(xlrange.min_colx, xlrange.max_colx + 1):
                cell = cell_map.get((rowx, colx))
                if cp:
                    cell = cell.copy()
                self.add_child(cell)
                cell_tag = cell_tag_map.get((rowx, colx))
                if cell_tag:
                    cell.extend_cell_tag(cell_tag)
                if cell.cell_tag:
                    if colx == xlrange.min_colx:
                        row.cell_tag = cell.cell_tag
                    if cell.cell_tag.extracell:
                        extra = ExtraCell(0, 0, 0, 0)
                        extra.cell_tag = cell.cell_tag
                        self.add_child(extra)
        if self._depth == 1:
            cell = ExtraCell(0, 0, 0, 0)
            self.add_child(cell)

    def set_parent(self, sheet_pos):
        if sheet_pos is self.sheet_pos:
            return
        else:
            self._parent = sheet_pos.get_pos(self.node_key)
            self.sheet_pos = sheet_pos

    def to_tag(self):
        tag = Node.to_tag(self)
        self.node_map[self.node_key] = self
        return self.node_tag + tag

    def next_row(self):
        self._parent.next_row()

    def write_cell(self, rowx, colx, value, cty):
        self._parent.write_cell(rowx, colx, value, cty)


def create_cell(rowx, colx, value, rich_text, data_type, font, rich_handler):
    s,cell_tag,head,tail = find_cell_tag(value)
    if s == '':
        cell = Cell(rowx, colx, s, data_type)
    elif xv_test(s):
        cell = XvCell(rowx, colx, s, data_type)
    elif not rich_text:
        cell = TagCell(rowx, colx, s, data_type, font, rich_handler)
    else:
        if head == 0 and tail == 0:
            cell = RichTagCell(rowx, colx, rich_text, data_type, font, rich_handler)
        else:
            _tail = head + len(s) - 1
            _rich,_text = rich_handler.mid(rich_text, head, _tail)
            cell = RichTagCell(rowx, colx, _rich, data_type, font, rich_handler)
    if cell_tag:
        cell.cell_tag = cell_tag
    return cell
