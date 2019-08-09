# encoding: utf-8

"""
Run-related proxy objects for python-docx, Run in particular.
"""

from __future__ import absolute_import, print_function, unicode_literals

from docxx.text.run import Run
from docxx.shared import Parented, ElementProxy, Pt, AlmostSingle

import copy

class Hyperlink(Parented):
    """
    Proxy object wrapping ``<w:hyperlink>`` element. Several of the properties on Run
    take a tri-state value, |True|, |False|, or |None|. |True| and |False|
    correspond to on and off respectively. |None| indicates the property is
    not specified directly on the run and its effective value is taken from
    the style hierarchy.
    """
    def __init__(self, h, parent):
        super().__init__(parent)
        self._h = self._element = self.element = h

    @property
    def runs(self):
        return [Run(r, self) for r in self._h.r_lst]
        
    # attributes
    @property
    def id(self):
        return self._h.id.val
        
    @id.setter
    def id(self, val):
        self._h.id.val = val
        
    @property
    def tooltip(self):
        return self._h.tooltip.val
        
    @tooltip.setter
    def tooltip(self, val):
        self._h.tooltip.val = val
        
        