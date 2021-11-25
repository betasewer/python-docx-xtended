# encoding: utf-8

"""
comments
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from docxx.document import _Body
from docxx.shared import ElementProxy

class Comments(ElementProxy):    
    @property
    def comments(self):
        return [Comment(n, self) for n in self._element.comment_lst]
    
    def add(self):
        # IDを発行
        if len(self._element.comment_lst) > 0:
            nid = self._element.comment_lst[-1].id + 1
        else:
            nid = 0
        # 要素を追加し初期化する
        cmt = self._element._add_comment()
        cmt.id = nid
        cmt.author = "machine"
        comment = Comment(cmt, self)
        return comment
        
class Comment(ElementProxy):    
    def __init__(self, r, parent):
        super().__init__(r)
        self.__comment = None
        
    @property
    def id(self):
        return self._element.id
        
    @property
    def paragraphs(self):
        return self._comment.paragraphs

    @property
    def lines(self):
        return [p.text for p in self.paragraphs]

    @property
    def text(self):
        return "\n".join(self.lines)
    
    @text.setter
    def text(self, text):
        for line in text.splitlines():
            self._comment.add_paragraph(line)
        
    @property
    def _comment(self):
        if self.__comment is None:
            self.__comment = _Body(self._element, self)
        return self.__comment

    
class CommentRangeStart(ElementProxy):  
    @property
    def notes(self):
        return [Note(n, self) for n in self._element.footnote_lst]
