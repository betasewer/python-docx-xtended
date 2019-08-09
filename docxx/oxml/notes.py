# encoding: utf-8

"""
Custom element classes related to the numbering part
"""

from docxx.oxml import OxmlElement
from docxx.oxml.shared import CT_DecimalNumber
from docxx.oxml.simpletypes import ST_DecimalNumber, ST_FtnEdn, ST_OnOff
from docxx.oxml.xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, RequiredAttribute, OptionalAttribute, ZeroOrMore, ZeroOrOne
)


class CT_Endnotes(BaseOxmlElement):
    """
    ``<w:num>`` element, which represents a concrete list definition
    instance, having a required child <w:abstractNumId> that references an
    abstract numbering definition that defines most of the formatting details.
    """
    endnote = ZeroOrMore('w:endnote',successors=())
    
class CT_Footnotes(BaseOxmlElement):
    """
    ``<w:num>`` element, which represents a concrete list definition
    instance, having a required child <w:abstractNumId> that references an
    abstract numbering definition that defines most of the formatting details.
    """
    footnote = ZeroOrMore('w:footnote',successors=())
    
    
class CT_FtnEdn(BaseOxmlElement):
    p = ZeroOrMore('w:p', successors=('w:altChunk',))
    tbl = ZeroOrMore('w:tbl', successors=('w:altChunk',))
    altChunk = ZeroOrMore('w:altChunk', successors=())
    
    type = OptionalAttribute('w:type', ST_FtnEdn)
    id = RequiredAttribute('w:id', ST_DecimalNumber)
    
    
class CT_FtnEdnRef(BaseOxmlElement):
    customMarkFollows = OptionalAttribute('w:customMarkFollows', ST_OnOff)
    id = RequiredAttribute('w:id', ST_DecimalNumber)
    

