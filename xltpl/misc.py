# -*- coding: utf-8 -*-
import sys
from jinja2 import Environment
from jinja2.exceptions import TemplateSyntaxError

class Env(Environment):

    def handle_exception(self, *args, **kwargs):
        exc_type, exc_value, tb = sys.exc_info()
        red_fmt = '\033[31m%s\033[0m'
        blue_fmt = '\033[34m%s\033[0m'
        error_type = red_fmt % ('error type:  %s' % exc_type)
        error_message = red_fmt % ('error message:  %s' % exc_value)
        print(error_type)
        print(error_message)
        if exc_type is TemplateSyntaxError:
            lineno = exc_value.lineno
            source = kwargs['source']
            src_lines = source.splitlines()
            for i, line in enumerate(src_lines):
                if i + 1 == lineno:
                    line_str = red_fmt % ('line %d : %s' % (i + 1, line))
                elif i + 1 in [lineno - 1, lineno + 1]:
                    line_str = blue_fmt % ('line %d : %s' % (i + 1, line))
                else:
                    line_str = 'line %d : %s' % (i + 1, line)
                print(line_str)
        Environment.handle_exception(self, *args, **kwargs)

class TreeProperty(object):

    def __init__(self, name):
        self.name = name
        self._name = '_' + name

    def __set__(self, instance, value):
        instance.__dict__[self._name] = value

    def __get__(self, instance, cls):
        if not hasattr(instance, self._name):
            instance.__dict__[self._name] = getattr(instance._parent, self.name)
        return instance.__dict__[self._name]

class CellTag():

    def __init__(self, cell_tag=dict()):
        self.beforerow = ''
        self.beforecell = ''
        self.aftercell = ''
        self.extracell = ''
        if cell_tag:
            self.__dict__.update(cell_tag)

    def extend(self, other):
        if isinstance(other, CellTag):
            self.beforerow = other.beforerow + self.beforerow
            self.beforecell = other.beforecell + self.beforecell
            self.aftercell += other.aftercell
            self.extracell += other.extracell
