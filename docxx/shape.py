# encoding: utf-8

"""
Objects related to shapes, visual objects that appear on the drawing layer of
a document.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from docxx.enum.shape import WD_INLINE_SHAPE
from docxx.oxml.ns import nsmap
from docxx.shared import Parented


class InlineShapes(Parented):
    """
    Sequence of |InlineShape| instances, supporting len(), iteration, and
    indexed access.
    """
    def __init__(self, body_elm, parent):
        super(InlineShapes, self).__init__(parent)
        self._body = body_elm

    def __getitem__(self, idx):
        """
        Provide indexed access, e.g. 'inline_shapes[idx]'
        """
        try:
            inline = self._inline_lst[idx]
        except IndexError:
            msg = "inline shape index [%d] out of range" % idx
            raise IndexError(msg)
        return InlineShape(inline)

    def __iter__(self):
        return (InlineShape(inline) for inline in self._inline_lst)

    def __len__(self):
        return len(self._inline_lst)

    @property
    def _inline_lst(self):
        body = self._body
        xpath = '//w:p/w:r/w:drawing/wp:inline'
        return body.xpath(xpath)


class BasicShape():
    """
    Basic implement of InlineShape and AnchorShape.
    """
    def __init__(self, element):
        self._element = element
        
    @property
    def height(self):
        """
        Read/write. The display height of this inline shape as an |Emu|
        instance.
        """
        return self._element.extent.cy

    @height.setter
    def height(self, cy):
        self._element.extent.cy = cy
        self._element.graphic.graphicData.pic.spPr.cy = cy

    @property
    def type(self):
        """
        The type of this inline shape as a member of
        ``docx.enum.shape.WD_INLINE_SHAPE``, e.g. ``LINKED_PICTURE``.
        Read-only.
        """
        graphicData = self._element.graphic.graphicData
        uri = graphicData.uri
        if uri == nsmap['pic']:
            blip = graphicData.pic.blipFill.blip
            if blip.link is not None:
                return WD_INLINE_SHAPE.LINKED_PICTURE
            return WD_INLINE_SHAPE.PICTURE
        if uri == nsmap['c']:
            return WD_INLINE_SHAPE.CHART
        if uri == nsmap['dgm']:
            return WD_INLINE_SHAPE.SMART_ART
        return WD_INLINE_SHAPE.NOT_IMPLEMENTED

    @property
    def width(self):
        """
        Read/write. The display width of this inline shape as an |Emu|
        instance.
        """
        return self._element.extent.cx

    @width.setter
    def width(self, cx):
        self._element.extent.cx = cx
        self._element.graphic.graphicData.pic.spPr.cx = cx
        
    def get_picture_rId(self):
        """
        埋め込み・リンク画像のrelationIdを返す。
        """
        t = self.type
        if t == WD_INLINE_SHAPE.LINKED_PICTURE:
            return self._element.graphic.graphicData.pic.blipFill.blip.link
        elif t == WD_INLINE_SHAPE.PICTURE:
            return self._element.graphic.graphicData.pic.blipFill.blip.embed
        else:
            return None
    

class InlineShape(BasicShape):
    """
    Proxy for an ``<wp:inline>`` element, representing the container for an
    inline graphical object.
    """
    def __init__(self, inline):
        super().__init__(inline)


class AnchorShape(BasicShape):
    """
    Proxy for an ``<wp:anchor>`` element.
    """
    def __init__(self, anchor):
        super().__init__(anchor)
        

    