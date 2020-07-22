# -*- coding: utf-8 -*-

from __future__ import print_function

from openpyxl.worksheet.cell_range import CellRange
from openpyxl.utils.cell import coordinate_to_tuple

from .utils import TreeProperty
from .utils import parse_range_tag, parse_cell_tag
from .pos import SheetPos, HRangePos, VRangePos


class XlRange(CellRange):

    node_map = TreeProperty('node_map')
    cell_map = TreeProperty('cell_map')
    range_tag_map = TreeProperty('range_tag_map')
    cell_tag_map = TreeProperty('cell_tag_map')

    def __init__(self, range_string=None, min_col=None, min_row=None,
                 max_col=None, max_row=None, title=None, index_base=1):
        CellRange.__init__(self, range_string, min_col, min_row, max_col, max_row, title)
        self._children = []
        self._children_split = []
        self.index_base = index_base
        self.offset = offset = 1 - index_base
        self.min_rowx = self.min_row - offset
        self.max_rowx = self.max_row - offset
        self.min_colx = self.min_col - offset
        self.max_colx = self.max_col - offset

    def __str__(self):
        fmt = '%s -> %s'
        return fmt % (self.__class__.__name__, self.rkey)

    @property
    def rkey(self):
        return '%s, %d' % (self.coord, self.depth)

    @property
    def root(self):
        if not hasattr(self, '_root'):
            if hasattr(self, 'parent'):
                if self.parent is self:
                    self._root = self
                else:
                    self._root = self.parent.root
            else:
                self._root = self
                self.parent = self
        return self._root

    @property
    def depth(self):
        if not hasattr(self, 'parent') or self.parent is self:
            return 0
        else:
            return self.parent.depth + 1

    def print(self):
        print('\t' * self.depth, self)
        for child in self._children:
            child.print()

    def print_split(self):
        print('\t' * self.depth, self)
        for child in self._children_split:
            child.print_split()

    def equal_col_span(self, other):
        return self.min_col == other.min_col and self.max_col == other.max_col

    def equal_row_span(self, other):
        return self.min_row == other.min_row and self.max_row == other.max_row

    def row_span_intersect(self, other):
        return (self.min_row < other.min_row and other.min_row < self.max_row < other.max_row) \
               or (other.min_row < self.min_row and self.min_row < other.max_row < self.max_row)

    def add_child(self, new_range):
        for old_range in list(self._children):
            if old_range == new_range:
                print(old_range, "equal to", new_range)
                return
            elif old_range > new_range:
                old_range.add_child(new_range)
                return
            elif old_range < new_range:
                new_range.add_child(old_range)
                self._children.remove(old_range)
        new_range.parent = self
        for index,old_range in enumerate(self._children):
            if not new_range.isdisjoint(old_range):
                raise Exception('intersect', new_range, old_range)
            if new_range.row_span_intersect(old_range):
                raise Exception('row span intersect', new_range, old_range)
            if new_range.min_row < old_range.min_row or \
                    (new_range.equal_col_span(old_range) and new_range.min_col < old_range.min_col):
                self._children.insert(index, new_range)
                return
        self._children.append(new_range)

    def split(self):
        if not self._children:
            return
        min_row = self.min_row
        while self._children:
            center_range = self._children.pop(0)
            if self == center_range:
                self._children = center_range._children
                self.split()
                return
            if min_row < center_range.min_row:
                up_range = HorizRange(min_row=min_row, max_row=center_range.min_row - 1,
                                      min_col=self.min_col, max_col=self.max_col, index_base=self.index_base)
                up_range.parent = self
                self._children_split.append(up_range)

            mid_range = HorizRange(min_row=center_range.min_row, max_row=center_range.max_row,
                                   min_col=self.min_col, max_col=self.max_col, index_base=self.index_base)
            mid_range.parent = self
            self._children_split.append(mid_range)
            min_col = self.min_col
            while True:
                if min_col < center_range.min_col:
                    left_range = VertRange(min_row=center_range.min_row, max_row=center_range.max_row,
                                           min_col=min_col, max_col=center_range.min_col - 1, index_base=self.index_base)
                    left_range.parent = mid_range
                    mid_range._children_split.append(left_range)
                    for child in list(self._children):
                        if child < left_range:
                            left_range.add_child(child)
                            self._children.remove(child)
                    left_range.split()

                center_range.parent = mid_range
                mid_range._children_split.append(center_range)
                center_range.split()

                min_col = center_range.max_col + 1
                if self._children:
                    next = self._children[0]
                    if next.equal_row_span(center_range):
                        center_range = self._children.pop(0)
                    else:
                        break
                else:
                    break

            if self.max_col >= min_col:
                right_range = VertRange(min_row=center_range.min_row, max_row=center_range.max_row,
                                        min_col=min_col, max_col=self.max_col, index_base=self.index_base)
                right_range.parent = mid_range
                mid_range._children_split.append(right_range)
                for child in list(self._children):
                    if child < right_range:
                        right_range.add_child(child)
                        self._children.remove(child)
                right_range.split()
            min_row = center_range.max_row + 1

        if min_row <= self.max_row:
            down_range = HorizRange(min_row=min_row, max_row=self.max_row,
                                    min_col=self.min_col, max_col=self.max_col, index_base=self.index_base)
            down_range.parent = self
            self._children_split.append(down_range)

    def to_tag(self):
        if self._children_split:
            tags = []
            for child in self._children_split:
                range_tag = self.range_tag_map.get(child.coord)
                if range_tag and range_tag.beforerange:
                    tags.append(range_tag.beforerange)
                    del self.range_tag_map[child.coord]
                tags.append(child.to_tag())
                if range_tag and range_tag.afterrange:
                    tags.append(range_tag.afterrange)
            if self is self.root:
                row_tag = "{%% row '%s' %%}" % self.rkey
                tags.append(row_tag)
            range_tag = '\n'.join(tags)
        else:
            range_tag = range_to_tag(self)
        return range_tag

    def get_pos(self, wtsheet, pos_parent=None):
        pos_parent = self.pos_cls(wtsheet, self, pos_parent)
        for child in self._children_split:
            p = child.get_pos(wtsheet, pos_parent)
        return pos_parent

def range_to_tag(xlrange):
    if xlrange.depth == 1:
        rkey = xlrange.root.rkey
    else:
        rkey = xlrange.rkey
    tags = []
    for rowx in range(xlrange.min_rowx, xlrange.max_rowx + 1):
        row_tag = "{%% row '%s' %%}" % rkey
        tag = xlrange.cell_tag_map.get((rowx, xlrange.min_colx))
        if tag and tag.beforerow:
            tags.append(tag.beforerow)
        tags.append(row_tag)
        for colx in range(xlrange.min_colx, xlrange.max_colx + 1):
            cell = xlrange.cell_map[(rowx, colx)]
            tag = xlrange.cell_tag_map.get((rowx, colx))
            if tag and tag.beforecell:
                tags.append(tag.beforecell)
            tags.append(cell.to_tag())
            if tag and tag.aftercell:
                tags.append(tag.aftercell)
    return '\n'.join(tags)


class SheetRange(XlRange):
    pos_cls = SheetPos

    def __init__(self, *args, **kwargs):
        XlRange.__init__(self, *args, **kwargs)
        self._node_map = {}
        self._cell_map = {}
        self._cell_tag_map = {}
        self._range_tag_map = {}

    def add_cell(self, cell):
        rowx = cell.rowx
        colx = cell.colx
        self._cell_map[(rowx, colx)] = cell
        cell.parent = self

    def add_cell_tag(self, cell_tag_map, rowx, colx):
        self._cell_tag_map[(rowx, colx)] = CellTag(cell_tag_map)

    def add_range(self, range_str, range_tag):
        new_range = VertRange(range_str, index_base=self.index_base)
        self.add_child(new_range)
        if range_tag:
            self._range_tag_map[new_range.coord] = RangeTag(range_tag)

    def parse_tag(self, lines, rowx, colx):
        lines = lines.split(';;')
        for line in lines:
            range_str, range_tag = parse_range_tag(line)
            if range_str:
                self.add_range(range_str, range_tag)
            else:
                cell_coord, cell_tag = parse_cell_tag(line)
                if cell_coord:
                    cell_rowx, cell_colx = coordinate_to_tuple(cell_coord)
                else:
                    cell_rowx = rowx
                    cell_colx = colx
                if cell_tag:
                    self.add_cell_tag(cell_tag, cell_rowx, cell_colx)
                else:
                    pass
                    #print(rowx, colx, 'no tag')
                    #print(line)

class HorizRange(XlRange):
    pos_cls = HRangePos

class VertRange(XlRange):
    pos_cls = VRangePos

class CellTag():

    def __init__(self, cell_tag=dict()):
        self.beforerow = ''
        self.beforecell = ''
        self.aftercell = ''
        if cell_tag:
            self.__dict__.update(cell_tag)

class RangeTag():

    def __init__(self, range_tag=dict()):
        self.beforerange = ''
        self.afterrange = ''
        if range_tag:
            self.__dict__.update(range_tag)