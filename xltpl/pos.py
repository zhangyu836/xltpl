# -*- coding: utf-8 -*-

from __future__ import print_function
from .utils import TreeProperty

class Pos():

    def __init__(self, min_rowx, min_colx):
        self.set_mins(min_rowx, min_colx)

    def next_cell(self):
        self.colx += 1
        if self.colx > self.max_colx:
            self.max_colx = self.colx
        return self.rowx, self.colx

    def next_row(self):
        self.rowx += 1
        self.colx = self.colx_start
        if self.rowx > self.max_rowx:
            self.max_rowx = self.rowx
        return self.rowx

    def coord(self):
        return self.rowx, self.colx

    def set_mins(self, min_rowx, min_colx):
        self.min_rowx = min_rowx
        self.min_colx = min_colx
        self.rowx_start = min_rowx - 1
        self.colx_start = min_colx - 1
        self.rowx = self.rowx_start
        self.colx = self.colx_start
        self.max_rowx = self.rowx_start
        self.max_colx = self.colx_start

class RangePos(Pos):

    pos_map = TreeProperty('pos_map')

    def __init__(self, wtsheet, xlrange, parent=None):
        Pos.__init__(self, xlrange.min_rowx, xlrange.min_colx)
        self.wtsheet = wtsheet
        self.xlrange = xlrange
        self._children = []
        if parent:
            parent.add_child(self)
        else:
            self._pos_map = {}
        self.pos_map[xlrange.rkey] = self

    def __str__(self):
        fmt = '%s -> %s'
        return fmt % (self.__class__.__name__, self.xlrange.rkey)

    @property
    def depth(self):
        if not hasattr(self, '_depth'):
            if not hasattr(self, 'parent') or self.parent is self:
                self._depth =  0
            else:
                self._depth = self.parent.depth + 1
        return self._depth

    def add_child(self, child):
        child.parent = self
        self._children.append(child)

    def get_node(self, key):
        self.current_node = self.xlrange.node_map.get(key)
        return self.current_node

    def write_cell(self, rdrowx, rdcolx, value, cty):
        wtrowx, wtcolx = self.next_cell()
        #print('write cell', rdrowx, rdcolx, wtrowx, wtcolx, value, cty)
        self.wtsheet.cell(rdrowx, rdcolx, wtrowx, wtcolx, value, cty)

    def print(self):
        print('\t' * self.depth, self)
        for child in self._children:
            child.print()

class SheetPos(RangePos):

    def __init__(self, wtsheet, xlrange, parent):
        RangePos.__init__(self, wtsheet, xlrange, parent)
        self.current_pos = self
        self.current_rkey = self.xlrange.rkey
        self.last_pos = None
        self.last_rkey = None
        self.parent = self

    def enter(self):
        pass

    def exit(self):
        pass

    def find_lca(self, pre, next):
        # find lowest common ancestor
        next_branch = []
        if pre.depth > next.depth:
            for i in range(next.depth, pre.depth):
                pre.exit()
                #print(pre, 'pre up', pre.parent)
                pre = pre.parent

        elif pre.depth < next.depth:
            for i in range(pre.depth, next.depth):
                next_branch.insert(0, next)
                #print(next, 'next up', next.parent)
                next = next.parent
        if pre is next:
            pass
        else:
            pre_parent = pre.parent
            next_parent = next.parent
            while pre_parent != next_parent:
                #print(pre, next, 'up together')
                pre.exit()
                pre = pre_parent
                pre_parent = pre.parent
                next_branch.insert(0, next)
                next = next_parent
                next_parent = next.parent
            pre.exit()
            if pre_parent._children.index(pre) > pre_parent._children.index(next):
                pre_parent.child_reenter()
                next.reenter()
            else:
                next.enter()

        for next in next_branch:
            next.enter()

    def get_pos(self, rkey):
        if rkey == self.current_rkey:
            return self.current_pos
        else:
            self.last_rkey = self.current_rkey
            self.last_pos = self.current_pos
            self.current_pos = self.pos_map.get(rkey)
            self.current_rkey = rkey
            self.find_lca(self.last_pos, self.current_pos)
        return self.current_pos

    def process_row(self, key, rv):
        current = self.get_pos(key)
        current.next_row()
        return 'row ' + key


class HRangePos(RangePos):

    def enter(self):
        #print(self, self.parent, 'hrange enter')
        #print(self.parent.max_rowx + 1, self.parent.min_colx, )
        self.set_mins(self.parent.max_rowx + 1, self.parent.min_colx)
        self.cells = {}

    def reenter(self):
        #print('reenter', self.parent.max_rowx + 1, self.parent.min_colx)
        self.parent.min_rowx = self.parent.max_rowx + 1
        self.set_mins(self.parent.max_rowx + 1, self.parent.min_colx)
        self.cells = {}

    def exit(self):
        self.parent.colx = self.colx
        #self.parent.max_colx = self.max_colx
        self.parent.rowx = self.max_rowx
        self.parent.max_rowx = max(self.max_rowx, self.parent.max_rowx)

        self.align_children()
        if self.depth == 1:
            self.write_cells()

    def child_reenter(self):
        self.align_children()

    def align_children(self):
        if len(self._children) == 1:
            child = self._children[0]
            self.cells.update(child.cells)
            #print(child.min_rowx, child.rowx, self.min_rowx, self.max_rowx)
            return
        for child in self._children:
            aligned = align(child.min_rowx, child.rowx, self.min_rowx, self.max_rowx)
            child.align(aligned)
            self.cells.update(child.cells)

    def write_cells(self):
        min_rowx, min_colx = min(self.cells)
        max_rowx, max_colx = max(self.cells)

        for wtrowx in range(min_rowx, max_rowx + 1):
            for wtcolx in range(min_colx, max_colx + 1):
                cell = self.cells.get((wtrowx, wtcolx))
                if cell:
                    cell.write_cell(self.wtsheet)
                    #print(cell, cell.rdrowx, cell.rdcolx, cell.wtrowx, cell.wtcolx, 'write cell')
                else:
                    pass
                    #print(wtrowx, wtcolx, 'no cell')


class VRangePos(RangePos):

    def enter(self):
        #print(self, self.parent, 'vrange enter')
        self.set_mins(self.parent.min_rowx, self.parent.colx + 1)
        self.cells = {}

    def reenter(self):
        #print(self, self.parent, 'vrange reenter')
        #print(self.parent.min_rowx, self.parent.max_rowx, 'vrange reenter')
        self.parent.min_rowx = self.parent.max_rowx + 1
        self.set_mins(self.parent.max_rowx + 1, self.parent.min_colx)
        self.cells = {}


    def exit(self):
        #print(self, self.parent, 'vrange exit')
        #print(self.max_rowx, self.max_colx)
        #print(self.parent.max_rowx, self.parent.max_colx)
        self.parent.colx = self.colx
        # self.parent.max_colx = self.max_colx
        self.parent.rowx = self.max_rowx
        self.parent.max_rowx = max(self.max_rowx, self.parent.max_rowx)
        self.merge_children()

    def child_reenter(self):
        self.merge_children()

    def merge_children(self):
        for child in self._children:
            self.cells.update(child.cells)

    def write_cell(self, rdrowx, rdcolx, value, cty):
        wtrowx, wtcolx = self.next_cell()
        #print('write cell', rdrowx, rdcolx, wtrowx, wtcolx, value, cty)
        #self.wtsheet.cell(rdrowx, rdcolx, wtrowx, wtcolx, value, cty)
        self.cells[(wtrowx, wtcolx)] = CachedCell(rdrowx, rdcolx, wtrowx, wtcolx, value, cty)

    def align(self, aligned):
        if not aligned:
            return
        for rdrowx, wtsetting in aligned:
            for colx in range(self.min_colx, self.colx + 1):
                cell = self.cells.get((rdrowx, colx))
                if not cell:
                    raise Exception('no cell to align')
                    continue
                for wtrowx,merged in wtsetting:
                    if merged:
                        self.cells[(wtrowx, colx)] = cell.create_mcell(wtrowx)
                    else:
                        cell.move_row(wtrowx)
                        self.cells[(wtrowx, colx)] = cell

class CachedCell():

    def __init__(self, rdrowx, rdcolx, wtrowx, wtcolx, value, cty):
        self.rdrowx = rdrowx
        self.rdcolx = rdcolx
        self.wtrowx = wtrowx
        self.wtcolx = wtcolx
        self.value = value
        self.cty = cty

    def move_row(self, target_rowx):
        self.wtrowx = target_rowx

    def write_cell(self, wtsheet):
        wtsheet.cell(self.rdrowx, self.rdcolx, self.wtrowx, self.wtcolx, self.value, self.cty)

    def create_mcell(self, target_rowx):
        return MergedCell(self.rdrowx, self.rdcolx, target_rowx, self.wtcolx, self)

class MergedCell():

    def __init__(self, rdrowx, rdcolx, wtrowx, wtcolx, cached_cell):
        self.rdrowx = rdrowx
        self.rdcolx = rdcolx
        self.wtrowx = wtrowx
        self.wtcolx = wtcolx
        self.cached_cell = cached_cell

    def move_row(self, target_rowx):
        self.wtrowx = target_rowx

    def create_mcell(self, target_rowx):
        return MergedCell(self.rdrowx, self.rdcolx, target_rowx, self.wtcolx, self.cached_cell)

    def write_cell(self, wtsheet):
        #return
        wtsheet.mcell(self.rdrowx, self.rdcolx, self.wtrowx, self.wtcolx, self.cached_cell.wtrowx)

def align(mina, maxa, minb, maxb):
    if mina != minb:
        print(mina, minb, maxa, maxb)
        raise Exception( 'mins not equal')
    if maxb < maxa:
        print(maxa, maxb)
        raise Exception('maxb smaller')
    if maxa == maxb:
        return
    a = maxa - mina + 1
    b = maxb - minb + 1 - a
    d, r = divmod(b, a)
    aligned = []
    rowb = minb
    for index,rowa in enumerate(range(mina, maxa + 1)):
        rowbs = []
        rowbs.append((rowb, 0))
        rowb += 1
        for _ in range(d):
            rowbs.append((rowb, 1))
            rowb += 1
        if index < r:
            rowbs.append((rowb, 1))
            rowb += 1
        rowbs.reverse()
        aligned.append((rowa, rowbs))
    aligned.reverse()
    return aligned

