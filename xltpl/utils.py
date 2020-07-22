# -*- coding: utf-8 -*-

from __future__ import print_function
import re

BLOCK_START_STRING = '{%'
BLOCK_END_STRING = '%}'
VARIABLE_START_STRING = '{{'
VARIABLE_END_STRING = '}}'
TAGTEST = '%s.+%s|%s.+%s' % (BLOCK_START_STRING, BLOCK_END_STRING, VARIABLE_START_STRING, VARIABLE_END_STRING)
XVTEST = '^ *%s *xv.+%s *$' % (BLOCK_START_STRING, BLOCK_END_STRING)
BLOCKTEST = '%s.+%s' % (BLOCK_START_STRING, BLOCK_END_STRING)

CELL = 'cell'
BEFOREROW = 'beforerow'
BEFORECELL = 'beforecell'
AFTERCELL = 'aftercell'
RANGE = 'range'
BEFORERANGE = 'beforerange'
AFTERRANGE = 'afterrange'
COORD = '[a-z]{1,3}\d+'
RANGETAG = '%s *%s *?(%s:%s) *%s' % (RANGE, VARIABLE_START_STRING, COORD, COORD, VARIABLE_END_STRING)
CELLTAG = '%s *%s *?(%s) *%s' % (CELL, VARIABLE_START_STRING, COORD, VARIABLE_END_STRING)
CELLSEPERATORS = [BEFOREROW, BEFORECELL, AFTERCELL]
CELLSPLITTER = '(%s)' % '|'.join(CELLSEPERATORS)
RANGESEPERATORS = [BEFORERANGE, AFTERRANGE]
RANGESPLITTER = '(%s)' % '|'.join(RANGESEPERATORS)

def tag_test(txt):
    p = re.compile(TAGTEST)
    rv = p.findall(txt)
    return bool(rv)

def xv_test(txt):
    p = re.compile(XVTEST)
    xv = p.findall(txt)
    return bool(xv)

def block_tag_test(txt):
    p = re.compile(BLOCKTEST)
    rv = p.findall(txt)
    return bool(rv)

def split(splitter, txt):
    parts = splitter.split(txt)
    d = {}
    for i in range(1, len(parts), 2):
            d[parts[i]] = parts[i + 1].strip()
    return d

def parse_cell_tag(txt):
    cell_pattern = re.compile(CELLTAG, re.I)
    split_pattern = re.compile(CELLSPLITTER, re.I)
    parts = cell_pattern.split(txt)
    if len(parts) > 1:
        tag = parts[2]
        tag_map = split(split_pattern, tag)
        return parts[1], tag_map
    else:
        tag = txt
        tag_map = split(split_pattern, tag)
        return None, tag_map
    return None, None

def parse_range_tag(txt):
    range_pattern = re.compile(RANGETAG, re.I)
    split_pattern = re.compile(RANGESPLITTER, re.I)
    parts = range_pattern.split(txt)
    if len(parts) > 1:
        tag = parts[2]
        tag_map = split(split_pattern, tag)
        return parts[1], tag_map
    return None, None


class TreeProperty(object):

    def __init__(self, name):
        self.name = name
        self._name = '_' + name

    def __set__(self, instance, value):
        instance.__dict__[self._name] = value

    def __get__(self, instance, cls):
        if not hasattr(instance, self._name):
            instance.__dict__[self._name] = getattr(instance.parent, self.name)
        return instance.__dict__[self._name]



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

