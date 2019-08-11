# -*- coding: utf-8 -*-

from jinja2 import nodes
from jinja2.ext import Extension

class CellExtension(Extension):
    tags = set(['cell'])

    def __init__(self, environment):
        super(self.__class__, self).__init__(environment)
        environment.extend(sheet_tpl = None)

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        body = parser.parse_statements(['name:endcell'], drop_needle=True)
        return nodes.CallBlock(self.call_method('_cell', args),
                               [], [], body).set_lineno(lineno)

    def _cell(self, number, caller):
        cell = self.environment.sheet_tpl.get_cell(number)
        rv = caller()
        rv = cell.process_rv(rv)
        return rv

class SectionExtension(Extension):
    tags = set(['sec'])

    def __init__(self, environment):
        super(self.__class__, self).__init__(environment)

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        body = parser.parse_statements(['name:endsec'], drop_needle=True)
        return nodes.CallBlock(self.call_method('_sec', args),
                               [], [], body).set_lineno(lineno)

    def _sec(self, number, caller):
        section = self.environment.sheet_tpl.get_section(number)
        rv = caller()
        rv = section.process_rv(rv)
        return rv

class RowExtension(Extension):
    tags = set(['row'])

    def __init__(self, environment):
        super(self.__class__, self).__init__(environment)

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        #body = parser.parse_statements(['name:endrow'], drop_needle=True)
        body = []
        return nodes.CallBlock(self.call_method('_row', args),
                               [], [], body).set_lineno(lineno)

    def _row(self, number, caller):
        row = self.environment.sheet_tpl.get_row(number)
        rv = caller()
        rv = row.process_rv(rv)
        return rv
