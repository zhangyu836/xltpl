# -*- coding: utf-8 -*-

import copy
from jinja2 import nodes
from jinja2.ext import Extension
from openpyxl.cell.text import RichText

def yes(font):
    wfont = copy.copy(font)
    wfont.name = 'Wingdings 2'
    return 'R', wfont

def yesx(font):
    wfont = copy.copy(font)
    wfont.rFont = 'Wingdings 2'
    return RichText(t='R', rPr=wfont)

def no():
    return u'□'

def yn(value, font, xlsx):
    if value:
        if xlsx:
            return yesx(font)
        else:
            return yes(font)
    else:
        return no()

class YnExtension(Extension):
    tags = set(['yn'])

    def __init__(self, environment):
        super(self.__class__, self).__init__(environment)

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        if parser.stream.skip_if('comma'):
            args.append(parser.parse_expression())
        else:
            args.append(nodes.Const(None))
        body = []
        return nodes.CallBlock(self.call_method('_yn', args),
                               [], [], body).set_lineno(lineno)

    def _yn(self, arg0, arg1, caller):
        section = self.environment.sheet_pos.current_node
        #rv = caller()
        if arg1 is not None:
            arg0 = not arg0
        rv = yn(arg0, section.font, self.environment.xlsx)
        rv = section.addv(rv)
        return rv