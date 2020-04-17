# -*- coding: utf-8 -*-

import six
from .utils import block_tag_test
from openpyxl.cell.text import RichText, Text

class RichTextHandler():

    def iter(self, rich_text, font):
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
    def rich_content(cls, value):
        return value

class RichTextHandlerX():

    @classmethod
    def iter(self, rich_text, font):
        for segment in rich_text.r:
            if segment.font is None and block_tag_test(segment.text):
                segment.font = font
            yield segment.text,segment.font,segment

    @classmethod
    def rich_segment(cls, text, font):
        return RichText(t=text, rPr=font)

    @classmethod
    def text_content(cls, value):
        return value

    @classmethod
    def rich_content(cls, value):
        return Text(r=value).content

rich_handlerx = RichTextHandlerX()
rich_handler = RichTextHandler()
