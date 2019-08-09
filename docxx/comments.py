# encoding: utf-8

"""
Styles object, container for all objects in the styles part.
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
    def texts(self):
        return [ p.text for p in self.paragraphs ]
        
    @property
    def _comment(self):
        if self.__comment is None:
            self.__comment = _Body(self._element, self)
        return self.__comment

    
class CommentRangeStart(ElementProxy):  
    @property
    def notes(self):
        return [Note(n, self) for n in self._element.footnote_lst]
