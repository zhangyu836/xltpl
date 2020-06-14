# -*- coding: utf-8 -*-


class Pos():

    def __init__(self, min_row, min_col):
        self.min_row = min_row - 1
        self.min_col = min_col - 1
        self.rowx = self.min_row
        self.colx = self.min_col

    def next_cell(self):
        self.colx += 1
        return self.rowx, self.colx

    def next_row(self):
        self.rowx += 1
        self.colx = self.min_col
        return self.rowx

    def coords(self):
        return self.rowx, self.colx
