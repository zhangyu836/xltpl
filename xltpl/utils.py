# -*- coding: utf-8 -*-

import re

BLOCK_START_STRING = '{%'
BLOCK_END_STRING = '%}'
VARIABLE_START_STRING = '{{'
VARIABLE_END_STRING = '}}'
BEFORECELL = 'beforecell'
BEFOREROW = 'beforerow'
AFTERCELL = 'aftercell'

Delimiters = [BEFORECELL, BEFOREROW, AFTERCELL]

def tag_test(txt):
    pattern = '%s.+%s|%s.+%s' % (BLOCK_START_STRING, BLOCK_END_STRING, VARIABLE_START_STRING, VARIABLE_END_STRING)
    p = re.compile(pattern)
    rv = p.findall(txt)
    return  bool(rv)

def xv_test(txt):
    pattern = '^ *%s *xv.+%s *$' % (BLOCK_START_STRING, BLOCK_END_STRING)
    p = re.compile(pattern)
    xv = p.findall(txt)
    return bool(xv)

def block_tag_test(txt):
    pattern = '%s.+%s' % (BLOCK_START_STRING, BLOCK_END_STRING)
    p = re.compile(pattern)
    rv = p.findall(txt)
    return  bool(rv)

def parse_tag(txt):
    pattern = '(%s)' % '|'.join(Delimiters)
    p = re.compile(pattern)
    parts = p.split(txt)
    d = {}
    for i in range(1, len(parts), 2):
        if tag_test(parts[i + 1]):
            d[parts[i]] = Tag(parts[i + 1])
    return d


class Tag():

    def __init__(self, tag):
        self.tag = tag

    def to_tag(self):
        return self.tag
