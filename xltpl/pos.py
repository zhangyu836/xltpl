# -*- coding: utf-8 -*-

class Pos():

    def __init__(self, min_rowx, min_colx):
        self.set_mins(min_rowx, min_colx)

    def next_cell(self):
        self.colx += 1
        return self.rowx, self.colx

    def next_row(self):
        self.rowx += 1
        self.colx = self.colx_start
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

