# encoding: utf-8

"""
Custom element classes related to paragraphs (CT_P).
"""

from docxx.oxml.ns import qn
from docxx.oxml.xmlchemy import BaseOxmlElement, OxmlElement, ZeroOrMore, ZeroOrOne

class CT_P(BaseOxmlElement):
    """
    ``<w:p>`` element, containing the properties and text for a paragraph.
    """
    content_tags = (
        'w:r', 'w:commentRangeStart', 'w:commentRangeEnd', 
        'w:bookmarkStart', 'w:bookmarkEnd','w:hyperlink'
    )
    pPr = ZeroOrOne('w:pPr')
    r = ZeroOrMore('w:r')
    commentRangeStart = ZeroOrMore('w:commentRangeStart')
    commentRangeEnd = ZeroOrMore('w:commentRangeEnd')
    bookmarkStart = ZeroOrMore('w:bookmarkStart')
    bookmarkEnd = ZeroOrMore('w:bookmarkEnd')
    hyperlink = ZeroOrMore('w:hyperlink')

    def _insert_pPr(self, pPr):
        self.insert(0, pPr)
        return pPr

    def add_p_before(self):
        """
        Return a new ``<w:p>`` element inserted directly prior to this one.
        """
        new_p = OxmlElement('w:p')
        self.addprevious(new_p)
        return new_p

    @property
    def alignment(self):
        """
        The value of the ``<w:jc>`` grandchild element or |None| if not
        present.
        """
        pPr = self.pPr
        if pPr is None:
            return None
        return pPr.jc_val

    @alignment.setter
    def alignment(self, value):
        pPr = self.get_or_add_pPr()
        pPr.jc_val = value

    def clear_content(self):
        """
        Remove all child elements, except the ``<w:pPr>`` element if present.
        """
        for child in self[:]:
            if child.tag == qn('w:pPr'):
                continue
            self.remove(child)

    def set_sectPr(self, sectPr):
        """
        Unconditionally replace or add *sectPr* as a grandchild in the
        correct sequence.
        """
        pPr = self.get_or_add_pPr()
        pPr._remove_sectPr()
        pPr._insert_sectPr(sectPr)

    @property
    def style(self):
        """
        String contained in w:val attribute of ./w:pPr/w:pStyle grandchild,
        or |None| if not present.
        """
        pPr = self.pPr
        if pPr is None:
            return None
        return pPr.style

    @style.setter
    def style(self, style):
        pPr = self.get_or_add_pPr()
        pPr.style = style
    
    @classmethod
    def get_contenttype(cls, child):
        for idx, tg in enumerate(cls.content_tags):
            if child.tag in qn(tg):
                return idx
        return -1
    
    def new_run(self):
        return self._new_r()

