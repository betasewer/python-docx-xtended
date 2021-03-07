# encoding: utf-8

"""|EndnotesPart|FootnotesPart| and closely related objects"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import os

from docxx.opc.constants import CONTENT_TYPE as CT
from docxx.opc.packuri import PackURI
from docxx.opc.part import XmlPart
from docxx.oxml import parse_xml
from docxx.notes import Endnotes, Footnotes

class EndnotesPart(XmlPart):
    """
    Proxy for the endnotes.xml part
    """
    @classmethod
    def default(cls, package):
        """
        Return a newly created styles part, containing a default set of
        elements.
        """
        partname = PackURI('/word/endnotes.xml')
        content_type = CT.WML_ENDNOTES
        element = parse_xml(cls._default_xml())
        return cls(partname, content_type, element, package)

    @property
    def endnotes(self):
        """
        The |_Styles| instance containing the styles (<w:style> element
        proxies) for this styles part.
        """
        return Endnotes(self._element, self)
        
    @classmethod
    def _default_xml(cls):
        """
        Return a bytestream containing XML for a default styles part.
        """
        return ""
    
    @property
    def part(self):
        """
        DocumentPartへと転送する。
        """
        return self._package.main_document_part

        
class FootnotesPart(XmlPart):
    """
    Proxy for the footnotes.xml part
    """
    @classmethod
    def default(cls, package):
        """
        Return a newly created styles part, containing a default set of
        elements.
        """
        partname = PackURI('/word/footnotes.xml')
        content_type = CT.WML_FOOTNOTES
        element = parse_xml(cls._default_xml())
        return cls(partname, content_type, element, package)

    @property
    def footnotes(self):
        """
        The |_Styles| instance containing the styles (<w:style> element
        proxies) for this styles part.
        """
        return Footnotes(self._element, self)
        
    @classmethod
    def _default_xml(cls):
        """
        Return a bytestream containing XML for a default styles part.
        """
        return ""
    
    @property
    def part(self):
        """
        DocumentPartへと転送する。
        """
        return self._package.main_document_part
