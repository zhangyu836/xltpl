# -*- coding: utf-8 -*-

class MergeMixin():

    def set_range(self, rdrowx=-1, rdcolx=-1, wtrowx=-1, wtcolx=-1):
        self.start_rdrowx = rdrowx
        self.start_rdcolx = rdcolx
        self.start_wtrowx = wtrowx
        self.start_wtcolx = wtcolx
        self.end_wtrowx = wtrowx
        self.end_wtcolx = wtcolx

    def is_in_range(self, rdrowx, rdcolx):
        return self._first_row <= rdrowx <= self._last_row and \
               self._first_col <= rdcolx <= self._last_col

    def to_be_merged(self, rdrowx, rdcolx):
        if rdrowx > self.start_rdrowx:
            return True
        else:
            return rdrowx == self.start_rdrowx and rdcolx > self.start_rdcolx

    def merge_cell(self, rdrowx, rdcolx, wtrowx, wtcolx):
        if not self.is_in_range(rdrowx, rdcolx):
            return False
        if self.start_rdrowx == -1:
            self.set_range(rdrowx, rdcolx, wtrowx, wtcolx)
        elif self.to_be_merged(rdrowx, rdcolx):
            self.end_wtrowx = max(self.end_wtrowx, wtrowx)
            self.end_wtcolx = max(self.end_wtcolx, wtcolx)
        else:
            self.new_range()
            self.set_range(rdrowx, rdcolx, wtrowx, wtcolx)
        return True

    def new_range(self):
        pass

    def collect_range(self):
        self.new_range()
        self.set_range()

class CellMerge(MergeMixin):

    def __init__(self, cell_range, merger):
        self.merger = merger
        self.set_range()
        rlo, rhi, clo, chi = cell_range
        self._first_row = rlo
        self._last_row = rhi - 1
        self._first_col = clo
        self._last_col = chi - 1

    def new_range(self):
        if self.start_wtrowx==self.end_wtrowx and self.start_wtcolx==self.end_wtcolx:
            return
        range = (self.start_wtrowx, self.end_wtrowx, self.start_wtcolx, self.end_wtcolx)
        self.merger.add_new_range(range)

class Merger:

    def __init__(self, sheet):
        self.range_list = []
        self._merge_list = []
        self.get_merge_list(sheet)

    def get_merge_list(self, sheet):
        for range in sheet.merged_cells:
            _merge = CellMerge(range, self)
            self._merge_list.append(_merge)

    def add_new_range(self, range):
        self.range_list.append(range)

    def merge_cell(self, rdrowx, rdcolx, wtrowx, wtcolx):
        for _merge in self._merge_list:
            is_in_range = _merge.merge_cell(rdrowx, rdcolx, wtrowx, wtcolx)
            if is_in_range:
                break

    def collect_range(self, wtsheet):
        for _merge in self._merge_list:
            _merge.collect_range()
        for range in self.range_list:
            wtsheet.merged_ranges.append(range)
        self.range_list.clear()
