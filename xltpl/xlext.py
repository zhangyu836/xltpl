# -*- coding: utf-8 -*-

import os
import six
from jinja2 import nodes
from jinja2.ext import Extension
from jinja2.runtime import Undefined
from inspect import isfunction

class NodeExtension(Extension):
    tags = set(['row', 'cell', 'node', 'extra'])

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        body = []
        return nodes.CallBlock(self.call_method('_node', args),
                               [], [], body).set_lineno(lineno)

    def _node(self, key, caller):
        node = self.environment.node_map.get_node(key)
        return str(key)

class SegmentExtension(Extension):
    tags = set(['seg'])

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        body = parser.parse_statements(['name:endseg'], drop_needle=True)
        return nodes.CallBlock(self.call_method('_seg', args),
                               [], [], body).set_lineno(lineno)

    def _seg(self, key, caller):
        segment = self.environment.node_map.get_node(key)
        rv = caller()
        rv = segment.process_rv(rv)
        return rv

class XvExtension(Extension):
    tags = set(['xv'])

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        if parser.stream.skip_if('comma'):
            args.append(parser.parse_expression())
        else:
            args.append(nodes.Const(0))
        body = []
        return nodes.CallBlock(self.call_method('_xv', args),
                               [], [], body).set_lineno(lineno)

    def _xv(self, xv, key, caller):
        if key==0:
            return six.text_type(xv)
        xvcell = self.environment.node_map.get_node(key)
        if xv is None or type(xv) is Undefined:
            xv = ''
        xvcell.rv = xv
        return six.text_type(xv)

class OpExtension(Extension):
    tags = set(['op'])

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        func_args = []
        while parser.stream.skip_if('comma'):
            func_args.append(parser.parse_expression())
        args.append(nodes.List(func_args))
        body = []
        return nodes.CallBlock(self.call_method('_op', args),
                               [], [], body).set_lineno(lineno)

    def _op(self, func, func_args, caller):
        if(isfunction(func)):
            node = self.environment.node_map.current_node
            node.add_op((func, func_args))
        return six.text_type(func)

class NoopExtension(Extension):
    tags = set(['op'])

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        func_args = []
        while parser.stream.skip_if('comma'):
            func_args.append(parser.parse_expression())
        args.append(nodes.List(func_args))
        body = []
        return nodes.CallBlock(self.call_method('_op', args),
                               [], [], body).set_lineno(lineno)

    def _op(self, func, func_args, caller):
        return six.text_type(func)

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
         return ''

try:
    pil = True
    from PIL.ImageFile import ImageFile
except:
    pil = False

class ImageRef():

    def __init__(self, image, image_index):
        self.image = image
        self.image_index = image_index
        self.rdrowx = -1
        self.rdcolx = -1
        self.wtrowx = -1
        self.wtcolx = -1
        if not isinstance(image, ImageFile):
            fname = six.text_type(image)
            if not os.path.exists(fname):
                self.image = None

    @property
    def image_key(self):
        return (self.rdrowx,self.rdcolx,self.image_index)

    @property
    def wt_top_left(self):
        return (self.wtrowx,self.wtcolx)

class ImagexExtension(Extension):
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

    def _image(self, image, image_index, caller):
        if not pil:
            return ''
        image_ref = ImageRef(image, image_index)
        if image_ref.image:
            node = self.environment.node_map.current_node
            node.set_image_ref(image_ref)
        return 'image'
