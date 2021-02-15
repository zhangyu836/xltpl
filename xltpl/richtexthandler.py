# -*- coding: utf-8 -*-

import six
from copy import copy
from .utils import block_tag_test, fix_test, tag_fix
from openpyxl.cell.text import RichText, Text

class RichTextHandler():

    def iter(self, rich_text, font):
        text_4_fix = self.text_4_fix(rich_text)
        if fix_test(text_4_fix):
            fixed = tag_fix(text_4_fix)
            #print(text_4_fix)
            #print(fixed)
            for i, segment in enumerate(rich_text):
                if i in fixed:
                    text = fixed[i]
                    if text == '':
                        continue
                    else:
                        yield text,segment[1],segment
        else:
            for segment in rich_text:
                yield segment[0],segment[1],segment

    @classmethod
    def rich_segment(self, text, font):
        return (text, font)

    @classmethod
    def text_content(cls, value):
        if isinstance(value, six.text_type):
            return value
        elif isinstance(value, list):
            x = []
            for text, font in value:
                x.append(text)
            return ''.join(x)

    @classmethod
    def text_4_fix(self, rich_text):
        text = []
        fmt = '___%d___'
        for i, segment in enumerate(rich_text):
            text.append(fmt % i)
            text.append(segment[0])
        return ''.join(text)

    @classmethod
    def rich_content(cls, value):
        segments = []
        for segment in value:
            if segment[0]:
                segments.append(segment)
        if segments:
            return segments
        else:
            return ''

    @classmethod
    def mid(cls, rich_text, head, tail):
        st = 0
        end = -1
        segments = []
        texts = []
        for index, segment in enumerate(rich_text):
            segment_text = segment[0]
            font = segment[1]
            l_text = len(segment_text)
            st = end + 1
            end += l_text
            if end < head:
                continue
            elif st <= head <= end:
                if end < tail:
                    text_st = head - st
                    text = segment_text[text_st:]
                    segments.append((text, font))
                    texts.append(text)
                else:
                    text_st = head - st
                    text_end = tail - st
                    text = segment_text[text_st:text_end + 1]
                    segments.append((text, font))
                    texts.append(text)
                    break
            elif end < tail:
                text = segment_text
                segments.append((text, font))
                texts.append(text)
            else:
                text_end = tail - st
                text = segment_text[:text_end + 1]
                segments.append((text, font))
                texts.append(text)
                break
        return segments, ''.join(texts)

class RichTextHandlerX():

    @classmethod
    def iter(self, rich_text, font):
        text_4_fix = self.text_4_fix(rich_text)
        if fix_test(text_4_fix):
            fixed = tag_fix(text_4_fix)
            #print(text_4_fix)
            #print(fixed)
            for i, segment in enumerate(rich_text.r):
                if i in fixed:
                    text = fixed[i]
                    if text == '':
                        continue
                    else:
                        segment_font = segment.font or font
                        yield text, segment_font, segment
        else:
            for segment in rich_text.r:
                segment_font = segment.font or font
                yield segment.text, segment_font, segment

    @classmethod
    def rich_segment(cls, text, font):
        return RichText(t=text, rPr=font)

    @classmethod
    def text_content(cls, value):
        return value

    @classmethod
    def text_4_fix(self, rich_text):
        text = []
        fmt = '___%d___'
        for i,segment in enumerate(rich_text.r):
            text.append(fmt % i)
            text.append(segment.text)
        return ''.join(text)

    @classmethod
    def rich_content(cls, value):
        if type(value) is Text:
            return value.content
        segments = []
        for segment in value:
            if segment.text:
                segments.append(segment)
        if segments:
            return Text(r=segments).content
        else:
            return ''

    @classmethod
    def mid(cls, rich_text, head, tail):
        st = 0
        end = -1
        segments = []
        texts = []
        for index, segment in enumerate(rich_text.r):
            l_text = len(segment.text)
            st = end + 1
            end += l_text
            if end < head:
                continue
            elif st <= head <= end:
                if end < tail:
                    text_st = head - st
                    segment_copy = copy(segment)
                    text = segment.text[text_st:]
                    segment_copy.text = text
                    segments.append(segment_copy)
                    texts.append(text)
                else:
                    text_st = head - st
                    text_end = tail - st
                    segment_copy = copy(segment)
                    text = segment.text[text_st:text_end+1]
                    segment_copy.text = text
                    segments.append(segment_copy)
                    texts.append(text)
                    break
            elif end < tail:
                segment_copy = copy(segment)
                text = segment.text
                #segment_copy.text = text
                segments.append(segment_copy)
                texts.append(text)
            else:
                text_end = tail - st
                segment_copy = copy(segment)
                text = segment.text[:text_end + 1]
                segment_copy.text = text
                segments.append(segment_copy)
                texts.append(text)
                break
        return Text(r=segments), ''.join(texts)


rich_handlerx = RichTextHandlerX()
rich_handler = RichTextHandler()
