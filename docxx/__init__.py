# encoding: utf-8

from docxx.api import open_docx, compose_docx  # noqa

__version__ = '0.1.0.0'

# register custom Part classes with opc package reader

from docxx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from docxx.opc.part import PartFactory
from docxx.opc.parts.coreprops import CorePropertiesPart
from docxx.parts.document import DocumentPart
from docxx.parts.hdrftr import FooterPart, HeaderPart
from docxx.parts.image import ImagePart
from docxx.parts.numbering import NumberingPart
from docxx.parts.settings import SettingsPart
from docxx.parts.styles import StylesPart
from docxx.parts.notes import EndnotesPart, FootnotesPart
from docxx.parts.comments import CommentsPart


def part_class_selector(content_type, reltype):
    if reltype == RT.IMAGE:
        return ImagePart
    return None


PartFactory.part_class_selector = part_class_selector
PartFactory.part_type_for[CT.OPC_CORE_PROPERTIES] = CorePropertiesPart
PartFactory.part_type_for[CT.WML_DOCUMENT_MAIN] = DocumentPart
PartFactory.part_type_for[CT.WML_FOOTER] = FooterPart
PartFactory.part_type_for[CT.WML_HEADER] = HeaderPart
PartFactory.part_type_for[CT.WML_NUMBERING] = NumberingPart
PartFactory.part_type_for[CT.WML_SETTINGS] = SettingsPart
PartFactory.part_type_for[CT.WML_STYLES] = StylesPart
PartFactory.part_type_for[CT.WML_ENDNOTES] = EndnotesPart
PartFactory.part_type_for[CT.WML_FOOTNOTES] = FootnotesPart
PartFactory.part_type_for[CT.WML_COMMENTS] = CommentsPart

del (
    CT,
    CorePropertiesPart,
    DocumentPart,
    FooterPart,
    HeaderPart,
    NumberingPart,
    PartFactory,
    SettingsPart,
    StylesPart,
    part_class_selector,
    EndnotesPart,
    FootnotesPart,
    CommentsPart,
)
