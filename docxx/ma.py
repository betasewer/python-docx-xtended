#!/usr/bin/env python3
# coding: UTF-8
#
from machaon.types.file import BasicLoadFile

from docxx import open_docx
from docxx.element import ElementPrinter, docx_ns, find_element, query, get_element
from docxx.opc.constants import RELATIONSHIP_TYPE, CONTENT_TYPE

class OpcPart():
    """ @type
    docx.OpcPackage.Partを操作する
    """
    def __init__(self, part, srcrel=None):
        self._part = part
        self._srcrel = srcrel

    def name(self):
        """ @method
        このパートの名前
        Returns:
            Str:
        """
        return self._part.partname

    def content_type(self):
        """ @method
        このパートのContent Type
        Returns:
            Str:
        """
        return self._part.content_type

    def content_typename(self):
        """ @method
        Content Typeの短縮版
        Returns:
            Str:
        """
        return CONTENT_TYPE.short_content_type(self._part.content_type)

    def element(self):
        """ @method
        中に入っているXMLの要素にアクセスする。
        Returns:
            XmlElement: 
        """
        if not self.is_xml_part():
            raise ValueError("Xmlの要素ではありません")
        root = get_element(self._part)
        return root
    
    def is_xml_part(self):
        return hasattr(self._part, "element")

    def save(self, app, outpath):
        """ @task
        ファイルに保存する。
        Params:
            outpath(Path): 
        """
        if outpath.isdir():
            outpath = outpath / (self.name().filename)
        
        if self.is_xml_part():
            import lxml.etree as etree
            elem = self.element()
            sb = etree.tostring(elem, pretty_print=True)
            with open(outpath, "wb") as fo:
                fo.write(sb)
        else:
            with open(outpath, "wb") as fo:
                fo.write(self._part.blob)
    
    def part(self, path):
        """ @method
        パスで指定した関係するパーツを取り出す。
        Params:
            path(str): 名前／[type=XXX]でrelation-typeによる検索／[id=XXX]でrelation-idによる検索
        Returns:
            OpcPart:
        """
        return next(_filter_parts(self._part.package, self._part.rels, path), None)

    def parts(self):
        """ @method
        関係するすべてのパーツをリストアップする。
        Returns:
            Sheet[OpcPart](name, source_relation_id, source_relation_typename):
        """
        return list(_list_parts(self._part.rels))

    def source_relation(self):
        """ @method
        このパーツを取得した親との関係。
        Returns:
            Any:
        """
        if self._srcrel is None:
            raise ValueError("パーツ取得者の情報がありません")
        return self._srcrel
    
    def source_relation_id(self):
        """ @method
        このパーツを取得した親との関係を特定するID。
        Returns:
            Str:
        """
        if self._srcrel is None:
            raise ValueError("パーツ取得者の情報がありません")
        return self._srcrel.rId

    def source_relation_typename(self):
        """ @method
        このパーツを取得した親との関係のタイプ。
        Returns:
            Str:
        """
        return RELATIONSHIP_TYPE.short_relation_type(self._srcrel.reltype)

    def construct(self, context, v):
        """ @meta
        """
        return OpcPart(v)

#
def _filter_parts(package, rels, xpath):
    """
    条件にあてはまるすべてのパーツを返す
    Params:
        xpath(str): 名前（前方一致）
            | 'type=XXX'でrelation-typeによる検索
            | 'id=XXX'でrelation-idによる検索
    """
    valtype, sep, value = xpath.partition("=")
    if sep:
        t = valtype.strip().upper()
        v = value.strip()
        return _list_part_by_relation(rels, t, v)
    else:
        if not xpath.startswith("/"):
            raise ValueError("パーツ名は/から開始します")
        for part in package.iter_parts():
            if part.partname.startswith(xpath):
                yield OpcPart(part)

def _list_part_by_relation(rels, t, v):
    if t == "TYPE":
        name = v.upper()
        reltype = getattr(RELATIONSHIP_TYPE, name, None)
        if reltype is None:
            raise ValueError("'{}'は存在しないRELATION_TYPEです".format(name))
        def predicate(x):
            return x.reltype == reltype
    elif t == "ID":
        def predicate(x):
            return x.rId == v
    else:
        raise ValueError("不明な検索タイプです")

    for rel in rels.values():
        if predicate(rel):
            yield OpcPart(rel.target_part, rel)

def _list_parts(rels):
    ps = []
    def _tryint(x):
        try: return int(x)
        except: return x
    for k, rel in sorted(rels.items(), key=lambda x:(x[0:3],_tryint(x[3:]))): # relIdでソートする
        ps.append(OpcPart(rel.target_part, rel))
    return ps

def package_filter_parts(pkg, xpath):
    """ パッケージのルートを基準にパスでパーツを示す。 """
    return _filter_parts(pkg, pkg.rels, xpath)


class OpcPackageFile(BasicLoadFile):
    def package(self):
        """ @method
        パッケージを得る。
        Returns:
            OpcPackageView
        """
        return self.file.package
    
    def part(self, path):
        """ @method
        opcパッケージのパートを得る。
        Params:
            path(str):
        Returns:
            OpcPart:
        """
        return next(package_filter_parts(self.package(), path), None)

    def walk_parts(self, path=None):
        """ @method
        含まれる全てのパートを得る。
        Returns:
            Sheet[OpcPart](name, content_typename):
        """
        if path is None:
            itr = self.package().iter_parts()
        else:
            itr = package_filter_parts(self.package(), path)
        return [OpcPart(x) for x in itr]

    def clone(self, path):
        """ パッケージを複製する。 """
        pkg = self.package().clone()
        return type(self)(path, file=pkg.main_document_part)


class XmlElement():
    """ @type trait
    lxml.etree.Elementを操作する。
    """
    class SpiritPrinter(ElementPrinter):
        def __init__(self, app):
            super().__init__(docx_ns)
            self.app = app

        def printer(self, s):
            self.app.post("message", s)
            
        def interrupted(self):
            return self.app.interruption_point()

    def xpath(self, elem, context, path):
        """ @method context
        XPathで子要素にアクセスする。
        Params:
            path(str): XPath
        Returns:
            Tuple: 
        """
        elems = elem.xpath(path) # docxx.oxml.xmlchemy._OxmlElementBase.xpathでnamespaceが渡される
        return [context.new_object(x, type=self) for x in elems]

    def children(self, elem, context):
        """ @method context
        すべての子要素を得る。
        Returns:
            Tuple: 
        """
        return [context.new_object(ch, type=self) for ch in elem]
    
    def first_child(self, elem, *tagnames):
        """ @method
        最初の要素を得る。
        Params:
            *tagnames(str):
        Returns:
            XmlElement:
        """
        el = elem.first_child_found_in(*tagnames)
        if el is None:
            raise ValueError("要素が見つかりません")
        return el
    
    def xml(self, elem):
        """ @method
        XMLを得る。
        Returns:
            Str:
        """
        return elem.xml

    def pprint(self, app, element):
        """ @meta """
        p = XmlElement.SpiritPrinter(app)
        p(element)
    

#
#
#
class DocxFile(OpcPackageFile):
    """ @type
    ワードファイル。
    """
    def loadfile(self):
        return open_docx(self.pathstr)

    def savefile(self, p):
        self.get_document().save(p)
    
    #
    #
    #
    def get_docx(self):
        return self.file
    
    def get_document(self):
        return self.file.document
    
    def get_footnotes(self):
        if self.file.footnotes is None:
            return []
        return self.file.footnotes.notes

    def get_endnotes(self):
        if self.file.endnotes is None:
            return []
        return self.file.endnotes.notes
    
    def get_comments(self):
        if self.file.comments is None:
            return []
        return self.file.comments.comments
    
    def text(self):
        """ @method
        本文のテキストを得る。
        Returns:
            Str:
        """
        s = []
        for p in self.get_document().paragraphs:
            s.append(p.text)
        return "\n".join(s)
    
    def concat(self, app, right):
        """ @task
        別のワードファイルを結合する。
        Params:
            right(DocxFile):
        Returns:
            DocxFile:
        """
        raise NotImplementedError()
    
    def pprint(self, app):
        """ @meta """
        app.post("message", "パス：{}".format(self.pathstr))

        doc = self.get_document()
        app.post("message", "paragraphs：{}".format(len(doc.paragraphs)))
        fnotes = self.get_footnotes()
        app.post("message", "footnotes：{}".format(len(fnotes)))
        enotes = self.get_endnotes()
        app.post("message", "endnotes：{}".format(len(enotes)))
        comnts = self.get_comments()
        app.post("message", "comments：{}".format(len(comnts)))
        app.post("message", "")

        sect1 = doc.sections[0]
        
        def mm(dim):
            return round(dim.mm, 2)

        ori = sect1.orientation
        width = mm(sect1.page_width)
        height = mm(sect1.page_height)
        sizefmt = get_size_format_name(width, height) or "unknown format"
        app.post("message", "page format：")
        app.post("message", "  {}x{}（{}:{}）".format(width, height, sizefmt, "horizontal" if ori==1 else "vertical"))
        app.post("message", "")

        marl = mm(sect1.left_margin)
        marr = mm(sect1.right_margin)
        mart = mm(sect1.top_margin)
        marb = mm(sect1.bottom_margin)
        marh = mm(sect1.header_distance)
        marf = mm(sect1.footer_distance)
        marg = mm(sect1.gutter)
        app.post("message", "page margin：")
        app.post("message", "    {} [header:{}] ".format(mart, marh))
        app.post("message", " {}    {}".format(marr, marl))
        app.post("message", "    {} [footer:{}] ".format(marb, marf))
        app.post("message", "")

        gridtype = sect1.grid_type
        app.post("message", "grid：")
        if gridtype=="linesAndChars":
            chars, lines = sect1.get_grid_dimension()
            app.post("message", "  {} char x {} line".format(chars, lines))
        elif gridtype=="lines":
            lines = sect1.grid_linepitch
            app.post("message", "  <free char> x {} line".format(lines))

        if sect1.text_direction=="tbRl":
            app.post("message", "  vertical")
        else:
            app.post("message", "  horizontal")

    
#
def get_size_format_name(dim1, dim2):
    found = set()
    for d in (dim1, dim2):
        if 363<d and d<365:
            found.add("B4-l")
        if 256<d and d<258:
            found.add("B4-s")
        if 296<d and d<298:
            found.add("A4-l")
        if 209<d and d<211:
            found.add("A4-s")
        if 255<d and d<257:
            found.add("B5-l")
        if 181<d and d<183:
            found.add("B5-s")
    
    if "B4-l" in found and "B4-s" in found:
        return "B4"
    elif "A4-l" in found and "A4-s" in found:
        return "A4"
    elif "B5-l" in found and "B5-s" in found:
        return "B5"
    return None
