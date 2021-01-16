# -*- coding: utf-8 -*-

from jinja2 import nodes
from jinja2.ext import Extension

class CellExtension(Extension):
    tags = set(['cell'])

    def __init__(self, environment):
        super(self.__class__, self).__init__(environment)
        environment.extend(sheet_pos = None)

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        body = parser.parse_statements(['name:endcell'], drop_needle=True)
        return nodes.CallBlock(self.call_method('_cell', args),
                               [], [], body).set_lineno(lineno)

    def _cell(self, key, caller):
        current_pos = self.environment.sheet_pos.current_pos
        cell = current_pos.get_node(key)
        rv = caller()
        rv = cell.process_rv(rv, current_pos)
        return rv

class SectionExtension(Extension):
    tags = set(['sec'])

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        body = parser.parse_statements(['name:endsec'], drop_needle=True)
        return nodes.CallBlock(self.call_method('_sec', args),
                               [], [], body).set_lineno(lineno)

    def _sec(self, key, caller):
        current_pos = self.environment.sheet_pos.current_pos
        section = current_pos.get_node(key)
        rv = caller()
        rv = section.process_rv(rv, current_pos)
        return rv

class RowExtension(Extension):
    tags = set(['row'])

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        #body = parser.parse_statements(['name:endrow'], drop_needle=True)
        body = []
        return nodes.CallBlock(self.call_method('_row', args),
                               [], [], body).set_lineno(lineno)

    def _row(self, key, caller):
        rv = caller()
        rv = self.environment.sheet_pos.process_row(key, rv)
        return rv

class XvExtension(Extension):
    tags = set(['xv'])

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        body = []
        return nodes.CallBlock(self.call_method('_row', args),
                               [], [], body).set_lineno(lineno)

    def _row(self, xv, caller):
        current_pos = self.environment.sheet_pos.current_pos
        node = current_pos.current_node
        #rv = caller()
        rv = node.process_xv(xv)
        return rv

class ImageExtension(Extension):
    tags = set(['img'])

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        if parser.stream.skip_if('comma'):
            args.append(parser.parse_expression())
        else:
            args.append(nodes.Const(0))
        body = []
        return nodes.CallBlock(self.call_method('_seg', args),
                               [], [], body).set_lineno(lineno)

    def _seg(self, image_ref, image_key, caller):
        current_pos = self.environment.sheet_pos.current_pos
        node = self.environment.sheet_pos.current_node
        image_key = node.get_image_key(image_key)
        current_pos.set_image_ref(image_ref, image_key)
        return ''