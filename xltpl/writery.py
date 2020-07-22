# -*- coding: utf-8 -*-

from .base import BookBase, SheetBase
from .basex import BookBase as BookBasex, SheetBase as SheetBasex
from .pos import Pos

class SheetWriterBase():

    def set_range(self, body_row):
        self.head_range = range(self.min_rowx, body_row)
        self.body_row = body_row
        self.foot_range = range(body_row + 1, self.max_rowx + 1)

    def copy_row(self, rdrowx):
        self.pos.next_row()
        for rdcolx in range(self.min_colx, self.max_colx + 1):
            wtrowx, wtcolx = self.pos.next_cell()
            self.cell(rdrowx, rdcolx, wtrowx, wtcolx)

    def copy_rows(self, row_range):
        for rdrowx in row_range:
            self.copy_row(rdrowx)

    def copy_head(self):
        self.copy_rows(self.head_range)

    def copy_foot(self):
        self.copy_rows(self.foot_range)

    def write_row(self, rdrowx, row):
        self.pos.next_row()
        for index,cell_value in enumerate(row):
            wtrowx, wtcolx = self.pos.next_cell()
            self.cell(rdrowx, index + self.index_base, wtrowx, wtcolx, cell_value)
        next_col = index + self.index_base + 1
        for rdcolx in range(next_col, self.max_colx + 1):
            wtrowx, wtcolx = self.pos.next_cell()
            self.cell(rdrowx, rdcolx, wtrowx, wtcolx, '')

    def write_rows(self, data):
        for row in data:
            self.write_row(self.body_row, row)

    def write_sheet(self, data):
        self.copy_head()
        self.write_rows(data)
        self.copy_foot()



class SheetWriter(SheetWriterBase, SheetBase):

    def __init__(self, bookwriter, rdsheet, sheet_name):
        SheetBase.__init__(self, bookwriter, rdsheet, sheet_name)
        self.min_rowx = 0
        self.min_colx = 0
        self.index_base = 0
        self.max_rowx = rdsheet.nrows - 1
        self.max_colx = rdsheet.ncols - 1
        self.set_range(rdsheet.body_row - 1)
        self.pos = Pos(self.min_rowx, self.min_colx)

class SheetWriterx(SheetWriterBase, SheetBasex):

    def __init__(self, bookwriter, rdsheet, sheet_name):
        SheetBasex.__init__(self, bookwriter, rdsheet, sheet_name)
        self.min_rowx = 1
        self.min_colx = 1
        self.index_base = 1
        self.max_rowx = rdsheet.max_row
        self.max_colx = rdsheet.max_column
        self.set_range(rdsheet.body_row)
        self.pos = Pos(self.min_rowx, self.min_colx)

class BookWriterBase():

    def __init__(self, fname, body_rows=list):
        self.load(fname)
        self.wtsheet_map = {}
        for rdsheet_idx,rdsheet in enumerate(self.rdsheet_list):
            try:
                body_row = body_rows[rdsheet_idx]
            except:
                body_row = 2
            rdsheet.body_row = body_row

    def write_list(self, ls, sheet_name=None, rdsheet_idx=0):
        sheet_writer = self.get_sheet_writer(sheet_name, rdsheet_idx)
        sheet_writer.write_sheet(ls)

    def write_dict(self, dict, rdsheet_idx=0):
        for key,ls in dict.items():
            sheet_name = key
            sheet_writer = self.get_sheet_writer(sheet_name, rdsheet_idx)
            sheet_writer.write_sheet(ls)

    def write_payload(self, payload, body_row_name='data'):
        idx = self.get_tpl_idx(payload)
        sheet_name = self.get_sheet_name(payload)
        ls = payload[body_row_name]
        self.write_list(ls, sheet_name, idx)

    def write_payloads(self, payloads, body_row_name='data'):
        for payload in payloads:
            self.write_payload(payload, body_row_name)


class BookWriter(BookWriterBase, BookBase):

    def create_wtbook(self):
        if not hasattr(self, 'wtbook'):
            self.create_workbook()

    def get_sheet_writer(self, sheet_name, rdsheet_idx=0):
        self.create_wtbook()
        if sheet_name is None:
            sheet_name = 'XLSheet%d' % len(self.wtbook.wtsheet_names)
        writer = self.wtsheet_map.get(sheet_name)
        if not writer:
            rdsheet = self.rdsheet_list[rdsheet_idx]
            writer = SheetWriter(self, rdsheet, sheet_name)
            self.wtsheet_map[sheet_name] = writer
        return writer


    def save(self, fname):
        if self.wtbook is not None:
            for wtsheet in self.wtsheet_map.values():
                wtsheet.set_mc_ranges()
            stream = open(fname, 'wb')
            self.wtbook.save(stream)
            stream.close()
            del self.wtbook
            self.wtsheet_map = {}


class BookWriterx(BookWriterBase, BookBasex):

    def get_sheet_writer(self, sheet_name, rdsheet_idx=0):
        if sheet_name is None:
            sheet_name = 'XLSheet%d' % len(self.workbook._sheets)
        writer = self.wtsheet_map.get(sheet_name)
        if not writer:
            rdsheet = self.rdsheet_list[rdsheet_idx]
            writer = SheetWriterx(self, rdsheet, sheet_name)
            self.wtsheet_map[sheet_name] = writer
        return writer

    def save(self, fname):
        for wtsheet in self.wtsheet_map.values():
            wtsheet.set_mc_ranges()
        self.workbook.save(fname)
        self.wtsheet_map = {}
        for sheet in self.workbook.worksheets:
            self.workbook.remove(sheet)


