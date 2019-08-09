# encoding: utf-8

"""
Custom element classes related to the numbering part
"""

from docxx.oxml import OxmlElement
from docxx.oxml.shared import CT_DecimalNumber
from docxx.oxml.simpletypes import ST_DecimalNumber, ST_String, ST_OnOff
from docxx.oxml.xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, RequiredAttribute, OptionalAttribute, ZeroOrMore, ZeroOrOne
)


class CT_Comments(BaseOxmlElement):
    """
    ``<w:num>`` element, which represents a concrete list definition
    instance, having a required child <w:abstractNumId> that references an
    abstract numbering definition that defines most of the formatting details.
    """
    comment = ZeroOrMore('w:comment',successors=())
    
class CT_Comment(BaseOxmlElement):
    p = ZeroOrMore('w:p', successors=('w:altChunk',))
    tbl = ZeroOrMore('w:tbl', successors=('w:altChunk',))
    altChunk = ZeroOrMore('w:altChunk', successors=())
    
    id = RequiredAttribute('w:id', ST_DecimalNumber)
    author = RequiredAttribute('w:author', ST_String)
