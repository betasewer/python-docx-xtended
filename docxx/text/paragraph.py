# encoding: utf-8

"""
Paragraph-related proxy types.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from docxx.enum.style import WD_STYLE_TYPE
from docxx.text.run import Run, same_run
from docxx.shared import Parented
from docxx.element import insert_element_next, insert_element_prev, same_element
from docxx.text.parfmt import ParagraphFormat
from docxx.text.hyperlink import Hyperlink


class Paragraph(Parented):
    """
    Proxy object wrapping ``<w:p>`` element.
    """
    def __init__(self, p, parent):
        super(Paragraph, self).__init__(parent)
        self._p = self._element = p

    def add_run(self, text=None, style=None):
        """
        Append a run to this paragraph containing *text* and having character
        style identified by style ID *style*. *text* can contain tab
        (``\\t``) characters, which are converted to the appropriate XML form
        for a tab. *text* can also include newline (``\\n``) or carriage
        return (``\\r``) characters, each of which is converted to a line
        break.
        """
        r = self._p.add_r()
        run = Run(r, self)
        if text:
            run.text = text
        if style:
            run.style = style
        return run

    def add_comment(self, text, headrun, tailrun):
        """
        文書に新しいコメントを挿入する。
        Params:
            text(str):
            headrun(Optional[Run]):
            tailrun(Optional[Run]):
        """
        comment = self.part.use_comments().add()
        if text:
            comment.text = text
        
        comment_beg = self._element._add_commentRangeStart()
        comment_beg.id = comment.id
        comment_end = self._element._add_commentRangeEnd()
        comment_end.id = comment.id
        for (_ct, run) in self.paragraph_contents:
            if same_element(run, headrun):
                insert_element_prev(run, comment_beg)
            if same_element(run, tailrun):
                insert_element_next(run, comment_end)
                break
        
        comment_refrun = self._element.add_r()
        comref = comment_refrun._add_commentReference()
        comref.id = comment.id
        insert_element_next(tailrun, comment_refrun)

        return comment

    def add_bookmark(self, id, headrun, tailrun, name):
        """  """
        beg = self._element._add_bookmarkStart()
        beg.id = id
        beg.name = name
        end = self._element._add_bookmarkEnd()
        end.id = id
        for (_ct, run) in self.paragraph_contents:
            if same_element(run, headrun):
                insert_element_prev(run, beg)
            if same_element(run, tailrun):
                insert_element_next(run, end)
                break        
        return end

    @property
    def alignment(self):
        """
        A member of the :ref:`WdParagraphAlignment` enumeration specifying
        the justification setting for this paragraph. A value of |None|
        indicates the paragraph has no directly-applied alignment value and
        will inherit its alignment value from its style hierarchy. Assigning
        |None| to this property removes any directly-applied alignment value.
        """
        return self._p.alignment

    @alignment.setter
    def alignment(self, value):
        self._p.alignment = value

    def clear(self):
        """
        Return this same paragraph after removing all its content.
        Paragraph-level formatting, such as style, is preserved.
        """
        self._p.clear_content()
        return self

    def insert_paragraph_before(self, text=None, style=None):
        """
        Return a newly created paragraph, inserted directly before this
        paragraph. If *text* is supplied, the new paragraph contains that
        text in a single run. If *style* is provided, that style is assigned
        to the new paragraph.
        """
        paragraph = self._insert_paragraph_before()
        if text:
            paragraph.add_run(text)
        if style is not None:
            paragraph.style = style
        return paragraph

    @property
    def paragraph_format(self):
        """
        The |ParagraphFormat| object providing access to the formatting
        properties for this paragraph, such as line spacing and indentation.
        """
        return ParagraphFormat(self._element)

    @property
    def runs(self):
        """
        Sequence of |Run| instances corresponding to the <w:r> elements in
        this paragraph.
        """
        return [Run(r, self) for r in self._p.r_lst]

    @property
    def style(self):
        """
        Read/Write. |_ParagraphStyle| object representing the style assigned
        to this paragraph. If no explicit style is assigned to this
        paragraph, its value is the default paragraph style for the document.
        A paragraph style name can be assigned in lieu of a paragraph style
        object. Assigning |None| removes any applied style, making its
        effective value the default paragraph style for the document.
        """
        style_id = self._p.style
        return self.part.get_style(style_id, WD_STYLE_TYPE.PARAGRAPH)

    @style.setter
    def style(self, style_or_name):
        style_id = self.part.get_style_id(
            style_or_name, WD_STYLE_TYPE.PARAGRAPH
        )
        self._p.style = style_id

    @property
    def text(self):
        """
        String formed by concatenating the text of each run in the paragraph.
        Tabs and line breaks in the XML are mapped to ``\\t`` and ``\\n``
        characters respectively.

        Assigning text to this property causes all existing paragraph content
        to be replaced with a single run containing the assigned text.
        A ``\\t`` character in the text is mapped to a ``<w:tab/>`` element
        and each ``\\n`` or ``\\r`` character is mapped to a line break.
        Paragraph-level formatting, such as style, is preserved. All
        run-level formatting, such as bold or italic, is removed.
        """
        text = ''
        for run in self.runs:
            text += run.text
        return text

    @text.setter
    def text(self, text):
        self.clear()
        self.add_run(text)

    def _insert_paragraph_before(self):
        """
        Return a newly created paragraph, inserted directly before this
        paragraph.
        """
        p = self._p.add_p_before()
        return Paragraph(p, self._parent)

    @property
    def paragraph_contents(self):    
        for child in self._element:
            ct = ContentType.find(child.tag) # local_part
            co = None
            if ct is CONTENT_TYPE_RUN:
                co = Run(child, self)
            elif ct is CONTENT_TYPE_HYPERLINK:
                co = Hyperlink(child, self)
            else:
                co = child
            yield (ct, co)

    # 所定の位置にランの集合を追加する
    def insert_runs(self, runs, last_run=None):
        for run in runs:
            if last_run is None:
                self._element.append(run._element)
            else:
                last_run._element.addnext(run._element)
            last_run = run

#
#
#
class ContentType:
    _tag_map = {}

    def __init__(self, tag):
        self.tag = tag

    @classmethod
    def define(cls, *tags):
        c = cls(tags[-1])
        cls._tag_map[tags[0]] = c
        return c

    @classmethod
    def find(cls, tag):
        _, localp = tag.split("}")
        if localp in cls._tag_map:
            return cls._tag_map[localp]
        else:
            return ContentType(localp)

    def __eq__(self, right):
        if not isinstance(right, str):
            raise TypeError(right)
        return self.tag == right


CONTENT_TYPE_RUN            = ContentType.define("r", "run")
CONTENT_TYPE_COMMENT_BEGIN  = ContentType.define("commentRangeStart")
CONTENT_TYPE_COMMENT_END    = ContentType.define("commentRangeEnd")
CONTENT_TYPE_HYPERLINK      = ContentType.define("hyperlink")
CONTENT_TYPE_BOOKMARK_BEGIN = ContentType.define("bookmarkStart")
CONTENT_TYPE_BOOKMARK_END   = ContentType.define("bookmarkEnd")
