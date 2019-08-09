# encoding: utf-8

"""
Styles object, container for all objects in the styles part.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from warnings import warn
from docxx.document import _Body
from docxx.shared import ElementProxy

class Endnotes(ElementProxy):    
    @property
    def notes(self):
        return [Note(n, self) for n in self._element.endnote_lst]
    
    
class Footnotes(ElementProxy):  
    @property
    def notes(self):
        return [Note(n, self) for n in self._element.footnote_lst]
        
        
class Note(ElementProxy):    
    def __init__(self, r, parent):
        super().__init__(r)
        self.__note = None
        
    @property
    def id(self):
        return self._element.id
    
    @property
    def type(self):
        el = self._element.type
        return el
        
    @type.setter
    def type(self, val):
        self._element.type = val
        
    @property
    def paragraphs(self):
        return self._note.paragraphs
    
    @property
    def _note(self):
        if self.__note is None:
            self.__note = _Body(self._element, self)
        return self.__note

