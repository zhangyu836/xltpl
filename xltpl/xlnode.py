# -*- coding: utf-8 -*-

import six
from openpyxl.utils import get_column_letter
from .utils import tag_test, xv_test, v_test
from .misc import TreeProperty
from .celltag import find_cell_tag, block_split_pattern, tag_parser

class DebugInfo():

    def __init__(self):
        self.value = None
        self.cell_tag = None

class Node(object):
    node_map = TreeProperty('node_map')
    ext_tag = 'node'

    def __init__(self):
        self._children = []
        self._parent = None
        self.no = 0
        self._depth = -1

    @property
    def depth(self):
        if self._depth == -1:
            if self._parent is None or self._parent is self:
                self._depth = 0
            else:
                self._depth = self._parent.depth + 1
        return self._depth

    @property
    def node_key(self):
        return '%s,%d' % (self._parent.node_key, self.no)

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
        self.node_map.put(self.node_key, self)
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

    def set_image_ref(self, image_ref):
        self._parent.set_image_ref(image_ref)

    def get_debug_info(self, offset):
        return self._parent.get_debug_info(offset)

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

    def get_debug_info(self, offset):
        debug = super().get_debug_info(offset)
        debug.value = self.text
        if debug.cell_tag:
            if debug.cell_tag.beforecell:
                if self.no==0:
                    if isinstance(self._parent, TagCell) or self._parent.no==0:
                        debug.value = debug.cell_tag.beforecell + debug.value
            if debug.cell_tag.aftercell:
                if self.no == len(self._parent._children) - 1:
                    if isinstance(self._parent, TagCell) or self._parent.no == len(self._parent._parent._children) - 1:
                        debug.value += debug.cell_tag.aftercell
        return debug

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

class OpSegment(Segment):

    @property
    def node_tag(self):
        fmt = "{%%seg '%s'%%}{%%endseg%%}%s"
        return fmt % (self.node_key, self.text)

    def add_op(self, op):
        self._parent.add_op(op)

class Section(Node):

    def __init__(self, text, font, rich_handler):
        Node.__init__(self)
        self.font = font
        self.rich_handler = rich_handler
        self.unpack(text)

    def unpack(self, text):
        parts = block_split_pattern.split(text)
        for index, part in enumerate(parts):
            if index % 2 == 0:
                if part == '':
                    continue
                child = Segment(part)
            else:
                tag = tag_parser.parse_tag(part)
                if tag == 'img':
                    child = ImageSegment(part)
                elif tag == 'yn':
                    child = RichSegment(part, self.font)
                elif tag == 'xv':
                    child = Segment(part)
                elif tag == 'op':
                    child = OpSegment(part)
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

    def add_op(self, op):
        self._parent.add_op(op)

class Cell(Node):
    ext_tag = 'cell'

    def __init__(self, sheet_cell, rowx, colx, value, cty):
        Node.__init__(self)
        self.sheet_cell = sheet_cell
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

    def write(self, rv, cty):
        self._parent.write_cell(self, rv, cty)

    def set_image_ref(self, image_ref):
        image_ref.rdrowx = self.rowx
        image_ref.rdcolx = self.colx
        self._parent.set_image_ref(image_ref)

    def get_coordinate(self, offset):
        col_letter = get_column_letter(self.colx + offset)
        return "%s%d" % (col_letter, self.rowx + offset)

    def get_debug_info(self, offset):
        debug = DebugInfo()
        coordinate = self.get_coordinate(offset)
        debug.coordinate = coordinate
        debug.address = 'Cell %s' % coordinate
        if self.cell_tag:
            debug.cell_tag = self.cell_tag
            debug.value = self.cell_tag.beforecell + str(self.value) + self.cell_tag.aftercell
        else:
            debug.value = self.value
        return debug

class TagCell(Section, Cell):

    def __init__(self, sheet_cell, rowx, colx, value, cty, font, rich_handler):
        Cell.__init__(self, sheet_cell, rowx, colx, value, cty)
        self.font = font
        self.rich_handler = rich_handler
        self.unpack(value)
        self.ops = []

    def exit(self):
        rv = self.pack()
        if not isinstance(rv, six.text_type):
            rv = self.rich_handler.rich_content(rv)
        self.write(rv, self.cty)
		
    def add_op(self, op):
        self.ops.append(op)


class RichTagCell(Cell):

    def __init__(self, sheet_cell, rowx, colx, value, cty, font, rich_handler):
        Cell.__init__(self, sheet_cell, rowx, colx, value, cty)
        self.font = font
        self.rich_handler = rich_handler
        self.unpack(value)
        self.ops = []

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
	
    def add_op(self, op):
        self.ops.append(op)

class EmptyCell(Cell):

    def __init__(self, rowx, colx):
        Node.__init__(self)
        self.sheet_cell = None
        self.rowx = rowx
        self.colx = colx
        self.value = None
        self.cty = None
        self.cell_tag = None

class XvCell(Cell):

    def __init__(self, sheet_cell, rowx, colx, value, cty, isXv):
        Cell.__init__(self, sheet_cell, rowx, colx, value, cty)
        self.isXv = isXv

    @property
    def node_tag(self):
        tag = self.value.strip()
        if self.isXv:
            head = tag[:-2].strip()
            #tag = "%s,%d%%}" % (head, self.node_key)
            tag = "%s,'%s'%%}" % (head, self.node_key)
        else:
            body = tag[2:-2].strip()
            #tag = "{%%xv %s,%d%%}" % (body, self.node_key)
            tag = "{%%xv %s,'%s'%%}" % (body, self.node_key)
        return tag

    def enter(self):
        self.rv = None

    def exit(self):
        self.write(self.rv, None)


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
        self._parent.write_row(self)

    def get_debug_info(self, offset):
        debug = DebugInfo()
        debug.address = 'Row %d' % (self.rowx + offset)
        if self.cell_tag:
            debug.value = self.cell_tag.beforerow
        return debug

class Tree(Node):
    ext_tag = 'tree'

    def __init__(self, index, node_map):
        Node.__init__(self)
        self._depth = 0
        self.no = index
        self.node_map = node_map
        self.sheet_writer = None

    @property
    def node_key(self):
        return str(self.no)

    def set_sheet_writer(self, sheet_writer):
        self.sheet_writer = sheet_writer
        self.node_map.set_current_node(self)

    def write_row(self, row_node):
        self.sheet_writer.write_row(row_node)

    def write_cell(self, cell_node, rv, cty):
        self.sheet_writer.write_cell(cell_node, rv, cty)

    def set_image_ref(self, image_ref):
        self.sheet_writer.set_image_ref(image_ref)

    def get_debug_info(self, offset):
        debug = DebugInfo()
        debug.address = 'Sheet %s' % self.no
        return debug

def create_cell(sheet_cell, rowx, colx, value, rich_text, data_type, font, rich_handler):
    s,cell_tag,head,tail = find_cell_tag(value)
    if s == '':
        cell = Cell(sheet_cell, rowx, colx, s, data_type)
    elif xv_test(s):
        cell = XvCell(sheet_cell, rowx, colx, s, data_type, True)
    elif v_test(s):
        cell = XvCell(sheet_cell, rowx, colx, s, data_type, False)
    elif not rich_text:
        cell = TagCell(sheet_cell, rowx, colx, s, data_type, font, rich_handler)
    else:
        if head == 0 and tail == 0:
            cell = RichTagCell(sheet_cell, rowx, colx, rich_text, data_type, font, rich_handler)
        else:
            _tail = head + len(s) - 1
            _rich,_text = rich_handler.mid(rich_text, head, _tail)
            cell = RichTagCell(sheet_cell, rowx, colx, _rich, data_type, font, rich_handler)
    if cell_tag:
        cell.cell_tag = cell_tag
    return cell
