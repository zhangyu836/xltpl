# -*- coding: utf-8 -*-

from copy import copy
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange

class MergeCell():

    def __init__(self, rdsheet, wtsheet):
        self.rdsheet = rdsheet
        self.wtsheet = wtsheet
        self.mc_ranges = {}

    @classmethod
    def get_sheet_mc_map(cls, sheet):
        mc_map = {}
        mc_nfa = {}
        for crange in sheet.merged_cells.ranges:
            rlo, rhi, clo, chi = crange.min_row, crange.max_row, crange.min_col, crange.max_col
            mc_map[(rlo, clo)] = (rlo, rhi, clo, chi)
            for rowx in range(rlo, rhi + 1):
                for colx in range(clo, chi + 1):
                    mc_nfa[(rowx, colx)] = (rlo, clo)
        sheet.mc_top_left_map = mc_map
        sheet.mc_already_set = mc_nfa

    def merge_cell(self, rdrowx, rdcolx, wtrowx, wtcolx):
        rdcoords2d = (rdrowx, rdcolx)
        if rdcoords2d in self.rdsheet.mc_top_left_map:
            if self.mc_ranges.get(rdcoords2d):
                rlo, rhi, clo, chi = self.mc_ranges.get(rdcoords2d)
                cr = CellRange(min_row=rlo, max_row=rhi, min_col=clo, max_col=chi)
                self.wtsheet.merged_cells.add(cr)
            self.mc_ranges[rdcoords2d] = (wtrowx, wtrowx, wtcolx, wtcolx)
        else:
            mc_top_left = self.rdsheet.mc_already_set.get(rdcoords2d)
            if mc_top_left:
                rlo, rhi, clo, chi = self.mc_ranges.get(mc_top_left)
                self.mc_ranges[mc_top_left] = (rlo, max(rhi, wtrowx), clo, max(chi, wtcolx))

    def merge_mcell(self, rdrowx, rdcolx, wtrowx, wtcolx, wt_top):
        rdcoords2d = (rdrowx, rdcolx)
        if rdcoords2d in self.rdsheet.mc_top_left_map:
            rlo, rhi, clo, chi = self.mc_ranges.get(rdcoords2d)
            self.mc_ranges[rdcoords2d] = (rlo, max(rhi, wtrowx), clo, max(chi, wtcolx))
        else:
            mc_top_left = self.rdsheet.mc_already_set.get(rdcoords2d)
            if mc_top_left:
                rlo, rhi, clo, chi = self.mc_ranges.get(mc_top_left)
                self.mc_ranges[mc_top_left] = (rlo, max(rhi, wtrowx), clo, max(chi, wtcolx))
            else:
                key = (wt_top, rdcoords2d)
                single_cell_cr = self.mc_ranges.get(key)
                if single_cell_cr:
                    rlo, rhi, clo, chi = self.mc_ranges.get(key)
                else:
                    rlo, rhi, clo, chi = wt_top, wtrowx, wtcolx, wtcolx
                self.mc_ranges[key] = (rlo, max(rhi, wtrowx), clo, max(chi, wtcolx))

    def finish(self):
        for key, crange in self.mc_ranges.items():
            rlo, rhi, clo, chi = crange
            cr = CellRange(min_row=rlo, max_row=rhi, min_col=clo, max_col=chi)
            self.wtsheet.merged_cells.add(cr)
        self.mc_ranges.clear()


class DataValidation():

    def __init__(self, rdsheet, wtsheet):
        self.rdsheet = rdsheet
        self.wtsheet = wtsheet
        self.dv_ranges = {}
        self.dv_copies = {}

    @classmethod
    def get_sheet_dv_map(cls, sheet):
        dv_map = {}
        dv_nfa = {}
        dv_orig_map = {}
        for dv in sheet.data_validations.dataValidation:
            for crange in dv.ranges:
                rlo, rhi, clo, chi = crange.min_row, crange.max_row, crange.min_col, crange.max_col
                dv_map[(rlo, clo)] = (rlo, rhi, clo, chi)
                dv_orig_map[(rlo, clo)] = dv
                for rowx in range(rlo, rhi + 1):
                    for colx in range(clo, chi + 1):
                        dv_nfa[(rowx, colx)] = (rlo, clo)
        sheet.dv_top_left_map = dv_map
        sheet.dv_already_set = dv_nfa
        sheet.dv_orig_map = dv_orig_map

    def get_dv(self, rdcoords2d):
        dv = self.dv_copies.get(rdcoords2d)
        if not dv:
            dv = copy(self.rdsheet.dv_orig_map.get(rdcoords2d))
            dv.ranges = MultiCellRange()
            self.dv_copies[rdcoords2d] = dv
        return dv

    def merge_cell(self, rdrowx, rdcolx, wtrowx, wtcolx):
        rdcoords2d = (rdrowx, rdcolx)
        if rdcoords2d in self.rdsheet.dv_top_left_map:
            if self.dv_ranges.get(rdcoords2d):
                rlo, rhi, clo, chi = self.dv_ranges.get(rdcoords2d)
                cr = CellRange(min_row=rlo, max_row=rhi, min_col=clo, max_col=chi)
                dv = self.get_dv(rdcoords2d)
                dv.ranges.add(cr)
                #print(dv, 'new range')
            self.dv_ranges[rdcoords2d] = (wtrowx, wtrowx, wtcolx, wtcolx)
        else:
            dv_top_left = self.rdsheet.dv_already_set.get(rdcoords2d)
            if dv_top_left:
                rlo, rhi, clo, chi = self.dv_ranges.get(dv_top_left)
                self.dv_ranges[dv_top_left] = (rlo, max(rhi, wtrowx), clo, max(chi, wtcolx))


    def merge_mcell(self, rdrowx, rdcolx, wtrowx, wtcolx, wt_top):
        rdcoords2d = (rdrowx, rdcolx)
        if rdcoords2d in self.rdsheet.dv_top_left_map:
            rlo, rhi, clo, chi = self.dv_ranges.get(rdcoords2d)
            self.dv_ranges[rdcoords2d] = (rlo, max(rhi, wtrowx), clo, max(chi, wtcolx))
        else:
            dv_top_left = self.rdsheet.dv_already_set.get(rdcoords2d)
            if dv_top_left:
                rlo, rhi, clo, chi = self.dv_ranges.get(dv_top_left)
                self.dv_ranges[dv_top_left] = (rlo, max(rhi, wtrowx), clo, max(chi, wtcolx))

    def finish(self):
        for key,crange in self.dv_ranges.items():
            rlo, rhi, clo, chi = crange
            dv = self.get_dv(key)
            cr = CellRange(min_row=rlo, max_row=rhi, min_col=clo, max_col=chi)
            dv.ranges.add(cr)
        for key,dv in self.dv_copies.items():
            self.wtsheet.data_validations.append(dv)
        self.dv_copies.clear()
        self.dv_ranges.clear()

class Merger(object):

    def __init__(self, rdsheet, wtsheet):
        self.mc = MergeCell(rdsheet, wtsheet)
        self.dv = DataValidation(rdsheet, wtsheet)

    @classmethod
    def get_sheet_maps(cls, sheet):
        MergeCell.get_sheet_mc_map(sheet)
        DataValidation.get_sheet_dv_map(sheet)

    def merge_cell(self, rdrowx, rdcolx, wtrowx, wtcolx):
        self.mc.merge_cell(rdrowx, rdcolx, wtrowx, wtcolx)
        self.dv.merge_cell(rdrowx, rdcolx, wtrowx, wtcolx)

    def merge_mcell(self, rdrowx, rdcolx, wtrowx, wtcolx, wt_top):
        self.mc.merge_mcell(rdrowx, rdcolx, wtrowx, wtcolx, wt_top)
        self.dv.merge_mcell(rdrowx, rdcolx, wtrowx, wtcolx, wt_top)

    def finish(self):
        self.mc.finish()
        self.dv.finish()