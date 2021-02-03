# -*- coding: utf-8 -*-

from __future__ import print_function
from .misc import TreeProperty

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
    nocache = TreeProperty('nocache')

    def __init__(self, wtsheet, xlrange, parent=None):
        Pos.__init__(self, xlrange.min_rowx, xlrange.min_colx)
        self.wtsheet = wtsheet
        self.xlrange = xlrange
        self.cells = {}
        self._children = []
        if parent:
            parent.add_child(self)
        else:
            self._pos_map = {}
        self.pos_map[xlrange.rkey] = self

    def __str__(self):
        fmt = '%s -> %s -> %s'
        return fmt % (self.__class__.__name__, self.xlrange.rkey, self.depth)

    @property
    def depth(self):
        if not hasattr(self, '_depth'):
            if not hasattr(self, '_parent') or self._parent is self:
                self._depth =  0
            else:
                self._depth = self._parent.depth + 1
        return self._depth

    def add_child(self, child):
        child._parent = self
        self._children.append(child)

    def write_cell(self, rdrowx, rdcolx, value, cty):
        wtrowx, wtcolx = self.next_cell()
        #print('write cell', rdrowx, rdcolx, wtrowx, wtcolx, value, cty)
        self.wtsheet.cell(rdrowx, rdcolx, wtrowx, wtcolx, value, cty)

    def print(self):
        print('\t' * self.depth, self)
        for child in self._children:
            child.print()

    def set_image_ref(self, image_ref, image_key):
        if hasattr(self.wtsheet, 'merger'):
            self.wtsheet.merger.set_image_ref(image_ref, (self.rowx,self.colx+1,image_key))

class SheetPos(RangePos):

    def __init__(self, wtsheet, xlrange, nocache=False):
        RangePos.__init__(self, wtsheet, xlrange, None)
        self.set_mins(xlrange.index_base, xlrange.index_base)
        self.current_node = self
        self.current_key = ''
        self.last_node = None
        self.last_key = None
        self._parent = self
        self._nocache = nocache
        self.node_map = xlrange.node_map

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
                #print(pre, 'pre up', pre._parent)
                pre = pre._parent

        elif pre.depth < next.depth:
            for i in range(pre.depth, next.depth):
                next_branch.insert(0, next)
                #print(next, 'next up', next._parent)
                next = next._parent
        if pre is next:
            pass
        else:
            pre_parent = pre._parent
            next_parent = next._parent
            while pre_parent != next_parent:
                #print(pre, next, 'up together')
                pre.exit()
                pre = pre_parent
                pre_parent = pre._parent
                next_branch.insert(0, next)
                next = next_parent
                next_parent = next._parent
            pre.exit()
            if pre_parent._children.index(pre) > pre_parent._children.index(next):
                pre_parent.child_reenter()
                next.reenter()
            else:
                next.enter()

        for next in next_branch:
            next.enter()
			
    def get_pos(self, key):
        self.current_pos = self.pos_map.get(key)
        return self.current_pos

    def get_crange(self, key):
        self.current_range = self.node_map.get(key)
        return self.current_range

    def get_node(self, key):
        if key == self.current_key:
            return self.current_node
        else:
            self.last_key = self.current_key
            self.last_node = self.current_node
            self.current_node = self.node_map.get(key)
            self.current_key = key
            self.find_lca(self.last_node, self.current_node)
        return self.current_node

class HRangePos(RangePos):

    def enter(self):
        #print(self, self._parent, 'hrange enter')
        #print(self._parent.max_rowx + 1, self._parent.min_colx, )
        self.set_mins(self._parent.max_rowx + 1, self._parent.min_colx)
        self.cells.clear()

    def reenter(self):
        #print('reenter', self._parent.max_rowx + 1, self._parent.min_colx)
        self._parent.min_rowx = self._parent.max_rowx + 1
        self.set_mins(self._parent.max_rowx + 1, self._parent.min_colx)
        self.cells.clear()

    def exit(self):
        self._parent.colx = self.colx
        #self._parent.max_colx = self.max_colx
        self._parent.rowx = self.max_rowx
        self._parent.max_rowx = max(self.max_rowx, self._parent.max_rowx)

        self.align_children()
        if self.depth == 1:
            self.write_cells()

    def child_reenter(self):
        self.align_children()

    def align_children(self):
        if self.nocache:
            return
        for child in self._children:
            if not child.cells:
                continue
            try:
                aligned = align(child.min_rowx, child.max_rowx, self.min_rowx, self.max_rowx)
                child.align(aligned)
            except Exception as e:
                print(e)
                print(child.cells)
            self.cells.update(child.cells)
            child.cells.clear()

    def write_cells(self):
        if not self.cells:
            return
        min_rowx, min_colx = min(self.cells)
        max_rowx, max_colx = max(self.cells)

        for wtrowx in range(min_rowx, max_rowx + 1):
            for wtcolx in range(min_colx, max_colx + 1):
                cell = self.cells.get((wtrowx, wtcolx))
                if cell:
                    cell.write_cell(self.wtsheet)
                else:
                    pass
                    #print(wtrowx, wtcolx, 'no cell')


class VRangePos(RangePos):

    def enter(self):
        #print(self, self._parent, 'vrange enter')
        self.set_mins(self._parent.min_rowx, self._parent.colx + 1)
        self.cells.clear()

    def reenter(self):
        #print(self, self._parent, 'vrange reenter')
        #print(self._parent.min_rowx, self._parent.max_rowx, 'vrange reenter')
        self._parent.min_rowx = self._parent.max_rowx + 1
        self.set_mins(self._parent.max_rowx + 1, self._parent.min_colx)
        self.cells.clear()


    def exit(self):
        #print(self, self._parent, 'vrange exit')
        #print(self.max_rowx, self.max_colx)
        #print(self._parent.max_rowx, self._parent.max_colx)
        self._parent.colx = self.colx
        # self._parent.max_colx = self.max_colx
        self._parent.rowx = self.max_rowx
        self._parent.max_rowx = max(self.max_rowx, self._parent.max_rowx)
        self.merge_children()
        self.min_rowx = self._parent.min_rowx

    def child_reenter(self):
        self.merge_children()

    def merge_children(self):
        for child in self._children:
            self.cells.update(child.cells)
            child.cells.clear()

    def write_cell(self, rdrowx, rdcolx, value, cty):
        wtrowx, wtcolx = self.next_cell()
        #print('write cell', rdrowx, rdcolx, wtrowx, wtcolx, value, cty)
        if self.nocache:
            self.wtsheet.cell(rdrowx, rdcolx, wtrowx, wtcolx, value, cty)
        else:
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

