# -*- coding: utf-8 -*-

import copy
from .patchx import *
from openpyxl import load_workbook
from openpyxl.cell.text import InlineFont
from openpyxl.cell.cell import Cell
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange
from openpyxl.utils import get_column_letter
from .merger import Merger
from .xlnode import Empty

class SheetBase():

    def __init__(self, bookwriter, rdsheet, sheet_name):
        self.workbook = bookwriter.workbook
        self.rdsheet = rdsheet
        self.wtsheet = self.workbook.create_sheet(title=sheet_name)
        self.copy_sheet_settings()
        self.merger = Merger(self.rdsheet, self.wtsheet)
        self.wtrows = set()
        self.wtcols = set()

    def copy_sheet_settings(self):
        self.wtsheet.sheet_format = copy.copy(self.rdsheet.sheet_format)
        self.wtsheet.sheet_properties = copy.copy(self.rdsheet.sheet_properties)
        # copy print settings
        self.wtsheet.page_setup = copy.copy(self.rdsheet.page_setup)
        self.wtsheet.print_options = copy.copy(self.rdsheet.print_options)
        self.wtsheet._print_rows = copy.copy(self.rdsheet._print_rows)
        self.wtsheet._print_cols = copy.copy(self.rdsheet._print_cols)
        self.wtsheet._print_area = copy.copy(self.rdsheet._print_area)
        self.wtsheet.page_margins = copy.copy(self.rdsheet.page_margins)
        self.wtsheet.protection = copy.copy(self.rdsheet.protection)
        self.wtsheet.HeaderFooter = copy.copy(self.rdsheet.HeaderFooter)
        self.wtsheet.views = copy.copy(self.rdsheet.views)
        self.wtsheet._images = copy.copy(self.rdsheet._images)

    def copy_row_dimension(self, rdrowx, wtrowx):
        if wtrowx in self.wtrows:
            return
        dim = self.rdsheet.row_dimensions.get(rdrowx)
        if dim:
            self.wtsheet.row_dimensions[wtrowx] = copy.copy(dim)
            self.wtsheet.row_dimensions[wtrowx].worksheet = self.wtsheet
            self.wtrows.add(wtrowx)

    def copy_col_dimension(self, rdcolx, wtcolx):
        if wtcolx in self.wtcols:
            return
        rdkey = get_column_letter(rdcolx)
        rddim = self.rdsheet.column_dimensions.get(rdkey)
        if not rddim:
            return
        wtdim = copy.copy(rddim)
        if rdcolx != wtcolx:
            wtkey = get_column_letter(wtcolx)
            wtdim.index = wtkey
            d = wtcolx - rdcolx
            wtdim.min += d
            wtdim.max += d
        else:
            wtkey = rdkey
        self.wtsheet.column_dimensions[wtkey] = wtdim
        self.wtsheet.column_dimensions[wtkey].worksheet = self.wtsheet
        self.wtcols.add(wtcolx)

    def _get_cell(self, rdrowx, rdcolx):
        try:
            return self.rdsheet._cells[(rdrowx, rdcolx)]
        except:
            return Cell(self, row=rdrowx, column=rdcolx)

    def _cell(self, rdrowx, rdcolx, wtrowx, wtcolx, value=None, data_type=None):
        if value is Empty:
            return
        source_cell = self._get_cell(rdrowx, rdcolx)
        target_cell = self.wtsheet.cell(column=wtcolx, row=wtrowx)
        if value is None:
            target_cell.value = source_cell._value
            target_cell.data_type = source_cell.data_type
        elif data_type:
            target_cell._value = value
            target_cell.data_type = data_type
        else:
            target_cell.value = value
        if source_cell.has_style:
            target_cell._style = copy.copy(source_cell._style)
        if source_cell.hyperlink:
            target_cell._hyperlink = copy.copy(source_cell.hyperlink)
        #if source_cell.comment:
        #    target_cell.comment = copy.copy(source_cell.comment)

    def cell(self, rdrowx, rdcolx, wtrowx, wtcolx, value=None, data_type=None):
        self.copy_row_dimension(rdrowx, wtrowx)
        self.copy_col_dimension(rdcolx, wtcolx)
        self.merge_cell(rdrowx, rdcolx, wtrowx, wtcolx)
        self._cell(rdrowx, rdcolx, wtrowx, wtcolx, value, data_type)

    def merge_cell(self, rdrowx, rdcolx, wtrowx, wtcolx):
        self.merger.merge_cell(rdrowx, rdcolx, wtrowx, wtcolx)

    def _mcell(self, rdrowx, rdcolx, wtrowx, wtcolx):
        source_cell = self._get_cell(rdrowx, rdcolx)
        target_cell = self.wtsheet.cell(column=wtcolx, row=wtrowx)
        if source_cell.has_style:
            target_cell._style = copy.copy(source_cell._style)

    def merge_mcell(self, rdrowx, rdcolx, wtrowx, wtcolx, wt_top):
        self.merger.merge_mcell( rdrowx, rdcolx, wtrowx, wtcolx, wt_top)

    def mcell(self, rdrowx, rdcolx, wtrowx, wtcolx, wt_top):
        self.copy_row_dimension(rdrowx, wtrowx)
        self.copy_col_dimension(rdcolx, wtcolx)
        self.merge_mcell(rdrowx, rdcolx, wtrowx, wtcolx, wt_top)
        self._mcell(rdrowx, rdcolx, wtrowx, wtcolx)

    def merge_finish(self):
        self.merger.finish()

class BookBase():

    def load(self, fname):
        self.workbook = load_workbook(fname)
        self.font_map = {}
        self.sheet_name_map = {}
        self.rdsheet_list = []
        for rdsheet in self.workbook.worksheets:
            self.get_sheet_maps(rdsheet)
            self.sheet_name_map[rdsheet.title] = len(self.sheet_name_map)
            self.rdsheet_list.append(rdsheet)
            self.workbook.remove(rdsheet)

    def get_sheet_maps(self, sheet):
        Merger.get_sheet_maps(sheet)

    def get_font(self, fontId):
        ifont = self.font_map.get(fontId)
        if ifont:
            return ifont
        else:
            font = self.workbook._fonts[fontId]
            ifont = InlineFont()
            ifont.rFont = font.name
            ifont.charset = font.charset
            ifont.family = font.family
            ifont.b = font.b
            ifont.i = font.i
            ifont.strike = font.strike
            ifont.outline = font.outline
            ifont.shadow = font.shadow
            ifont.condense = font.condense
            ifont.extend = font.extend
            ifont.color = font.color
            ifont.sz = font.sz
            ifont.u = font.u
            ifont.vertAlign = font.vertAlign
            ifont.scheme = font.scheme
            self.font_map[fontId] = ifont
            return ifont


    def get_tpl_idx(self, payload):
        idx = payload.get('tpl_idx')
        if not idx:
            name = payload.get('tpl_name')
            if name:
                idx = self.sheet_name_map[name]
            else:
                idx = 0
        return idx

    def get_sheet_name(self, payload, key=None):
        name = payload.get('sheet_name')
        if not name:
            if key:
                name = key
            else:
                name = 'XLSheet%d' % len(self.workbook._sheets)
        return name
