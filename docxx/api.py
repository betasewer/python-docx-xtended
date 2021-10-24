# encoding: utf-8

"""
Directly exposed API functions and classes, :func:`Document` for now.
Provides a syntactically more convenient API for interacting with the
OpcPackage graph.
"""

from __future__ import absolute_import, division, print_function

from typing import Generic, TypeVar

import os

from docxx.opc.constants import CONTENT_TYPE as CT
from docxx.package import Package
from docxx.element import copy_element, remove_element, query
from docxx.parts.document import DocumentPart
from copy import deepcopy

T = TypeVar("T")
class DocumentOpener(Generic[T]):
    def __init__(self, content_type, file_type, template_file):
        self.content_type = content_type
        self.file_type = file_type
        self.template_file = template_file
    
    def default_document_path(self):
        """
        Return the path to the built-in default .docx/xlsx/pptx package.
        """
        return self.template_file
        
    def __call__(self, path=None) -> T:
        if path is None:    
            path = self.default_document_path()
            
        document_part = Package.open(path).main_document_part
        ct = document_part.content_type
        if ct != self.content_type:
            tmpl = "file '{}' has unsupported content type '{}'.".format(path, ct) 
            raise ValueError(tmpl)
            
        return document_part

def _templatefile(filename):
    thisdir = os.path.split(__file__)[0]
    return os.path.join(thisdir, 'templates', filename)
    
open_docx: DocumentOpener[DocumentPart] = DocumentOpener(
    CT.WML_DOCUMENT_MAIN, "Word File", _templatefile("default.docx")
)

del _templatefile


def compose_docx(base=None, template=None):
    """
    Create an empty new document from base document or template document
    """
    if base is not None:
        newdocx = deepcopy(base)
        newdocx.document._body.clear_content()
    else:
        newdocx = open_docx()    
    newbody = newdocx.document._element.body

    if template:
        templbody = template.document._element.body
        # テンプレート文書から最初のセクションの要素をコピーする
        # 頁判型・頁余白・行数・行文字数・文字列方向
        remove_element(newbody, query("w:sectPr"))
        srcsection = templbody.get_or_add_sectPr()
        destsection = newbody.get_or_add_sectPr()
        copy_element(destsection, srcsection, query(
            "w:pgSz", 
            "w:pgMar", 
            "w:cols", 
            "w:textDirection", 
            "w:docGrid"
        ))
        # TODO: テンプレート文書からヘッダー・フッターをコピーする。
    else:
        """
        # 一から空のファイルをつくる
        sec = newdocx.document.sections[0]
        sec.text_direction = ST_TextDirection.TB_RL
        sec.page_width = Mm(297)
        sec.page_height = Mm(210)
        sec.grid_lines = 300
        """
        pass # Not implemented yet
        
    return newdocx

# for compatiblity with python-docx
Document = open_docx
