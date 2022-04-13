# -*- coding: utf-8 -*-

import re

BLOCK_START_STRING = '{%'
BLOCK_END_STRING = '%}'
VARIABLE_START_STRING = '{{'
VARIABLE_END_STRING = '}}'
TAGTEST = '%s.+%s|%s.+%s' % (BLOCK_START_STRING, BLOCK_END_STRING, VARIABLE_START_STRING, VARIABLE_END_STRING)
XVTEST = '^ *%s *xv.+%s *$' % (BLOCK_START_STRING, BLOCK_END_STRING)
SingleXV = '%s *xv.+?%s' % (BLOCK_START_STRING, BLOCK_END_STRING)
VTEST = '^%s.+%s$' % (VARIABLE_START_STRING, VARIABLE_END_STRING)
SingleV = '%s.+?%s' % (VARIABLE_START_STRING, VARIABLE_END_STRING)
BLOCKTEST = '%s.+%s' % (BLOCK_START_STRING, BLOCK_END_STRING)

CELL = 'cell'
BEFOREROW = 'beforerow'
BEFORECELL = 'beforecell'
AFTERCELL = 'aftercell'
RANGE = 'range'
BEFORERANGE = 'beforerange'
AFTERRANGE = 'afterrange'
COORD = '[a-z]{1,3}\d+'
RANGETAG = '%s *%s *(%s:%s\S*?) *%s' % (RANGE, VARIABLE_START_STRING, COORD, COORD, VARIABLE_END_STRING)
CELLTAG = '%s *%s *(%s) *%s' % (CELL, VARIABLE_START_STRING, COORD, VARIABLE_END_STRING)
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
    p2 = re.compile(SingleXV)
    if p.findall(txt):
        parts = p2.split(txt)
        if len(parts) < 3:
            return True

def v_test(txt):
    p = re.compile(VTEST)
    p2 = re.compile(SingleV)
    if p.findall(txt):
        parts = p2.split(txt)
        if len(parts) < 3:
            return True

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


CTEx = ' *yn | *xv | *img '
BLOCKSPLIT = '((?:%s(?:(?!%s).)+?%s)+)' % (BLOCK_START_STRING, CTEx, BLOCK_END_STRING)
CTEx2 = ' *yn | *img '
YNSPLIT = '(%s(?:(?=%s)).+?%s)' % (BLOCK_START_STRING, CTEx2, BLOCK_END_STRING)
IMGTEST = '%s *img .+%s' % (BLOCK_START_STRING, BLOCK_END_STRING)

def block_split(txt):
    split_pattern = re.compile(BLOCKSPLIT)
    parts = split_pattern.split(txt)
    return parts

def rich_split(txt):
    split_pattern = re.compile(YNSPLIT)
    parts = split_pattern.split(txt)
    return parts

def img_test(txt):
    p = re.compile(IMGTEST)
    rv = p.findall(txt)
    return bool(rv)


FIXTEST = '({(?:___\d+___)?(?:{|%).+?(?:}|%)(?:___\d+___)?})'
RUNSPLIT = '(___\d+___)'
RUNSPLIT2 = '___(\d+)___'

#need a better way to handle this
def fix_test(txt):
    split_pattern = re.compile(FIXTEST)
    parts = split_pattern.split(txt)
    for i, part in enumerate(parts):
        if i % 2 == 1:
            pattern = re.compile(RUNSPLIT)
            rv = pattern.findall(part)
            if rv :
                return True

def tag_fix(txt):
    split_pattern = re.compile(FIXTEST)
    parts = split_pattern.split(txt)
    p = ''
    for i, part in enumerate(parts):
        if i % 2 == 1:
            p += fix_step2(part)
        else:
            p += part
    split_pattern2 = re.compile(RUNSPLIT2)
    parts = split_pattern2.split(p)
    d = {}
    for i in range(1, len(parts), 2):
            d[int(parts[i])] = parts[i + 1]
    return d

def fix_step2(txt):
    split_pattern = re.compile(RUNSPLIT)
    parts = split_pattern.split(txt)
    p0 = ''
    p1 = ''
    for index,part in enumerate(parts):
        if index % 2 == 0:
            p0 += part
        else:
            p1 += part
    return p0+p1
