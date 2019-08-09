# encoding: utf-8

"""
|EndnotesPart|FootnotesPart| and closely related objects
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import os

from docxx.opc.constants import CONTENT_TYPE as CT
from docxx.opc.packuri import PackURI
from docxx.opc.part import XmlPart
from docxx.oxml import parse_xml
from docxx.comments import Comments


class CommentsPart(XmlPart):
    """
    Proxy for the endnotes.xml part
    """
    @classmethod
    def default(cls, package):
        """
        Return a newly created styles part, containing a default set of
        elements.
        """
        partname = PackURI('/word/comments.xml')
        content_type = CT.WML_COMMENTS
        element = parse_xml(cls._default_xml())
        return cls(partname, content_type, element, package)

    @property
    def comments(self):
        return Comments(self._element)
        
    @classmethod
    def _default_xml(cls):
        """
        Return a bytestream containing XML for a default styles part.
        """
        return ""
