import re
from jinja2.lexer import Lexer
from jinja2.environment import Environment

BLOCK_START_STRING = '{%'
BLOCK_END_STRING = '%}'
BLOCKSPLIT = '(%s.*?%s)' % (BLOCK_START_STRING, BLOCK_END_STRING)
block_split_pattern = re.compile(BLOCKSPLIT)

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


class TagParser():

    def __init__(self):
        env = Environment()
        self.lexer = Lexer(env)

    def parse(self, text):
        tokens = self.lexer.tokenize(text)
        begin = next(tokens)
        tag = next(tokens)
        ignored = False
        beforerow = False
        if tag.value in ['yn', 'img', 'xv', 'op']:
            ignored = True
        if begin.value[-1] == '-':
            beforerow = True
        return ignored,beforerow

    def parse_tag(self, text):
        tokens = self.lexer.tokenize(text)
        begin = next(tokens)
        return next(tokens).value

tag_parser = TagParser()
def find_cell_tag(text):
    parts = block_split_pattern.split(text)
    #print(parts)
    _beforerow = ''
    _beforecell = ''
    _aftercell = ''
    stop = True
    index = 0
    for part in parts:
        if index % 2 == 0:
            if part == '':
                index += 1
                continue
            break
        else:
            ignored, is_beforerow = tag_parser.parse(part)
            if ignored:
                break
            if is_beforerow:
                _beforerow += part
                index += 1
            else:
                stop = False
                break
    if not stop:
        for part in parts[index:]:
            if index % 2 == 0:
                if part == '':
                    index += 1
                    continue
                break
            else:
                ignored, _ = tag_parser.parse(part)
                if ignored:
                    break
                _beforecell += part
                index += 1
    _reversed = reversed(parts[index:])
    index = 0
    for part in _reversed:
        if index % 2 == 0:
            if part == '':
                index += 1
                continue
            break
        else:
            ignored, _ = tag_parser.parse(part)
            if ignored:
                break
            _aftercell = part + _aftercell
            index += 1
    if _beforerow or _beforecell or _aftercell:
        head = len(_beforerow ) + len(_beforecell)
        tail = len(_aftercell)
        cell_tag = CellTag()
        cell_tag.beforerow = _beforerow
        cell_tag.beforecell = _beforecell
        cell_tag.aftercell = _aftercell
        s = text[head:-tail]
        return s, cell_tag, head, tail
    else:
        return text, None, 0, 0