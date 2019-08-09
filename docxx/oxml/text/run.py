# encoding: utf-8

"""
Custom element classes related to text runs (CT_R).
"""

from docxx.oxml import parse_xml
from docxx.oxml.ns import nsdecls, qn
from docxx.oxml.simpletypes import ST_BrClear, ST_BrType, ST_String, ST_OnOff
from docxx.oxml.xmlchemy import (
    BaseOxmlElement, OptionalAttribute, ZeroOrMore, ZeroOrOne, 
    OneAndOnlyOne, RequiredAttribute
)
from copy import deepcopy

class CT_Br(BaseOxmlElement):
    """
    ``<w:br>`` element, indicating a line, page, or column break in a run.
    """
    type = OptionalAttribute('w:type', ST_BrType)
    clear = OptionalAttribute('w:clear', ST_BrClear)


class CT_R(BaseOxmlElement):
    """
    ``<w:r>`` element, containing the properties and text for a run.
    """
    rPr = ZeroOrOne('w:rPr')
    t = ZeroOrMore('w:t')
    br = ZeroOrMore('w:br')
    cr = ZeroOrMore('w:cr')
    tab = ZeroOrMore('w:tab')
    drawing = ZeroOrMore('w:drawing')
    ruby = ZeroOrMore('w:ruby')
    
    delText = ZeroOrMore('w:delText')
    instrText = ZeroOrMore('w:instrText')
    delInstrText = ZeroOrMore('w:delInstrText')
    fldChar = ZeroOrMore('w:fldChar')
    
    footnoteRef = ZeroOrMore('w:footnoteRef')
    endnoteRef = ZeroOrMore('w:endnoteRef') 
    footnoteReference = ZeroOrMore('w:footnoteReference')
    endnoteReference = ZeroOrMore('w:endnoteReference')
    
    noBreakHyphen = ZeroOrMore('w:noBreakHyphen')
    softHyphen = ZeroOrOne('w:softHyphen')
    separator = ZeroOrOne('w:separator')
    continuationSeparator = ZeroOrOne('w:continuationSeparator')
    #
    #w:dayShort [0..1]    Date Block - Short Day Format
    #w:monthShort [0..1]    Date Block - Short Month Format
    #w:yearShort [0..1]    Date Block - Short Year Format
    #w:dayLong [0..1]    Date Block - Long Day Format
    #w:monthLong [0..1]    Date Block - Long Month Format
    #w:yearLong [0..1]    Date Block - Long Year Format
    #w:annotationRef [0..1]    Comment Information Block
    #w:sym [0..1]    Symbol Character
    #w:pgNum [0..1]    Page Number Block
    #w:object    Inline Embedded Object
    #w:pict    VML Object
    #w:commentReference    Comment Content Reference Mark
    #w:ptab [0..1]    Absolute Position Tab Character
    #w:lastRenderedPageBreak [0..1]    Position of Last Calculated Page Break
    #

    def _insert_rPr(self, rPr):
        self.insert(0, rPr)
        return rPr

    def add_t(self, text):
        """
        Return a newly added ``<w:t>`` element containing *text*.
        """
        t = self._add_t(text=text)
        if len(text.strip()) < len(text):
            t.set(qn('xml:space'), 'preserve')
        return t

    def add_drawing(self, inline_or_anchor):
        """
        Return a newly appended ``CT_Drawing`` (``<w:drawing>``) child
        element having *inline_or_anchor* as its child.
        """
        drawing = self._add_drawing()
        drawing.append(inline_or_anchor)
        return drawing

    def clear_content(self):
        """
        Remove all child elements except the ``<w:rPr>`` element if present.
        """
        content_child_elms = self[1:] if self.rPr is not None else self[:]
        for child in content_child_elms:
            self.remove(child)

    @property
    def style(self):
        """
        String contained in w:val attribute of <w:rStyle> grandchild, or
        |None| if that element is not present.
        """
        rPr = self.rPr
        if rPr is None:
            return None
        return rPr.style

    @style.setter
    def style(self, style):
        """
        Set the character style of this <w:r> element to *style*. If *style*
        is None, remove the style element.
        """
        rPr = self.get_or_add_rPr()
        rPr.style = style

    @property
    def text(self):
        """
        A string representing the textual content of this run, with content
        child elements like ``<w:tab/>`` translated to their Python
        equivalent.
        """
        text = ''
        for child in self:
            if child.tag == qn('w:t'):
                t_text = child.text
                text += t_text if t_text is not None else ''
            elif child.tag == qn('w:tab'):
                text += '\t'
            elif child.tag in (qn('w:br'), qn('w:cr')):
                text += '\n'
        return text

    @text.setter
    def text(self, text):
        self.clear_content()
        _RunContentAppender.append_to_run_from_text(self, text)
        
    # clear_contentせず、テキスト関連のみを変更する
    def set_text(self, text):
        for child in self:
            tag = child.tag
            if any([ child.tag == qn(x) for x in ['w:t','w:tab','w:br','w:cr'] ]):
                self.remove(child)              
        _RunContentAppender.append_to_run_from_text(self, text)

    # 要素をディープコピーする
    def copy(self, text):
        el = deepcopy(self)
        el.text = text
        return el
    



class CT_Text(BaseOxmlElement):
    """
    ``<w:t>`` element, containing a sequence of characters within a run.
    """

class CT_Ruby(BaseOxmlElement):
    """
    ``<w:ruby>`` element
    """
    rubyPr    = OneAndOnlyOne('w:rubyPr')
    rt        = OneAndOnlyOne('w:rt')
    rubyBase  = OneAndOnlyOne('w:rubyBase')
    
    @classmethod
    def new(cls, ptBaseText):
        rb = parse_xml(cls._ruby_xml(ptBaseText))
        return rb

    @classmethod
    def _ruby_xml(cls, ptBaseText):
        hpss = cls.calc_hps(ptBaseText)
        return (
            '<w:ruby {}>\n'
            '  <w:rubyPr>\n'
            '    <w:rubyAlign w:val="distributeSpace"/>\n'
            '    <w:hps w:val="{}"/>\n'
            '    <w:hpsRaise w:val="{}"/>\n'
            '    <w:hpsBaseText w:val="{}"/>\n'
            '    <w:lid w:val="ja-JP"/>\n'
            '  </w:rubyPr>\n'
            '  <w:rt>\n'
            '  </w:rt>\n'
            '  <w:rubyBase>\n'
            '  </w:rubyBase>\n'
            '</w:ruby>'
            .format( nsdecls('w'), hpss[0], hpss[1], hpss[2] )
        )
        
    @classmethod
    def calc_hps(cls, ptBaseText):
        hpsBase = ptBaseText.pt*2
        if hpsBase < 16:
            hpsBase = 16
        elif hpsBase % 2 == 1:
            hpsBase -= 1
        hpsBase = int(hpsBase)
        return ( hpsBase//2, hpsBase-2, hpsBase )        
    
class CT_RubyPr(BaseOxmlElement):
    """
    ``<w:rubyPr>`` element
    """
    rubyAlign = OneAndOnlyOne('w:rubyAlign')
    hps = OneAndOnlyOne('w:hps')
    hpsRaise = OneAndOnlyOne('w:hpsRaise')
    hpsBaseText = OneAndOnlyOne('w:hpsBaseText')
    lid = OneAndOnlyOne('w:lid')
    dirty = ZeroOrOne('w:dirty', successors=())        

class CT_RubyContent(BaseOxmlElement):
    """
    ``<w:rt>`` ``<w:rubyBase>`` element
    """
    r = ZeroOrOne('w:r')
    
class CT_RubyAlign(BaseOxmlElement):
    val = RequiredAttribute('w:val', ST_String) # ST_RubyAlign

class CT_Lang(BaseOxmlElement):
    val = RequiredAttribute('w:val', ST_String) # ST_Lang
    
class CT_FldChar(BaseOxmlElement):
    fldCharType = RequiredAttribute('w:fldCharType', ST_String) # ST_FldCharType
    fldLock = OptionalAttribute('w:fldLock', ST_OnOff)
    dirty = OptionalAttribute('w:dirty', ST_OnOff) 


class _RunContentAppender(object):
    """
    Service object that knows how to translate a Python string into run
    content elements appended to a specified ``<w:r>`` element. Contiguous
    sequences of regular characters are appended in a single ``<w:t>``
    element. Each tab character ('\t') causes a ``<w:tab/>`` element to be
    appended. Likewise a newline or carriage return character ('\n', '\r')
    causes a ``<w:cr>`` element to be appended.
    """
    def __init__(self, r):
        self._r = r
        self._bfr = []

    @classmethod
    def append_to_run_from_text(cls, r, text):
        """
        Create a "one-shot" ``_RunContentAppender`` instance and use it to
        append the run content elements corresponding to *text* to the
        ``<w:r>`` element *r*.
        """
        appender = cls(r)
        appender.add_text(text)

    def add_text(self, text):
        """
        Append the run content elements corresponding to *text* to the
        ``<w:r>`` element of this instance.
        """
        for char in text:
            self.add_char(char)
        self.flush()

    def add_char(self, char):
        """
        Process the next character of input through the translation finite
        state maching (FSM). There are two possible states, buffer pending
        and not pending, but those are hidden behind the ``.flush()`` method
        which must be called at the end of text to ensure any pending
        ``<w:t>`` element is written.
        """
        if char == '\t':
            self.flush()
            self._r.add_tab()
        elif char in '\r\n':
            self.flush()
            self._r.add_br()
        else:
            self._bfr.append(char)

    def flush(self):
        text = ''.join(self._bfr)
        if text:
            self._r.add_t(text)
        del self._bfr[:]
