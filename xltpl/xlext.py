# -*- coding: utf-8 -*-

import os
import six
from jinja2 import nodes
from jinja2.ext import Extension
from jinja2.runtime import Undefined
from .xlnode import XvCell

class NodeExtension(Extension):
    tags = set(['row', 'cell', 'node', 'extra'])

    def __init__(self, environment):
        super(self.__class__, self).__init__(environment)
        environment.extend(sheet_pos = None)

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        body = []
        return nodes.CallBlock(self.call_method('_node', args),
                               [], [], body).set_lineno(lineno)

    def _node(self, key, caller):
        node = self.environment.sheet_pos.get_node(key)
        return key

class SegmentExtension(Extension):
    tags = set(['seg'])

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        body = parser.parse_statements(['name:endseg'], drop_needle=True)
        return nodes.CallBlock(self.call_method('_seg', args),
                               [], [], body).set_lineno(lineno)

    def _seg(self, key, caller):
        segment = self.environment.sheet_pos.get_node(key)
        rv = caller()
        rv = segment.process_rv(rv)
        return rv

class XvExtension(Extension):
    tags = set(['xv'])

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        body = []
        return nodes.CallBlock(self.call_method('_xv', args),
                               [], [], body).set_lineno(lineno)

    def _xv(self, xv, caller):
        segment = self.environment.sheet_pos.current_node
        if xv is None or type(xv) is Undefined:
            xv = ''
        if isinstance(segment._parent, XvCell):
            segment._parent.rv = xv
        rv = six.text_type(xv)
        return rv

try:
    pil = True
    from PIL.ImageFile import ImageFile
except:
    pil = False

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
        return nodes.CallBlock(self.call_method('_image', args),
                               [], [], body).set_lineno(lineno)

    def _image(self, image_ref, image_key, caller):
        if not pil:
            return ''
        node = self.environment.sheet_pos.current_node
        if not isinstance(image_ref, ImageFile):
            fname = six.text_type(image_ref)
            if not os.path.exists(fname):
                image_ref = ''
        node.set_image_ref(image_ref, image_key)
        return 'image'

class RangeExtension(Extension):
    tags = set(['crange'])

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        body = []
        return nodes.CallBlock(self.call_method('_range', args),
                               [], [], body).set_lineno(lineno)

    def _range(self, key, caller):
        sheet_pos = self.environment.sheet_pos
        cr = sheet_pos.get_crange(key)
        cr.set_parent(sheet_pos)
        return key
