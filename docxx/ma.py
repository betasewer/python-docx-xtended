
from docxx import open_docx
from machaon.types.file import BasicLoadFile


class DocxFile(BasicLoadFile):
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

    def package(self):
        """ @method
        パッケージを得る。
        Returns:
            OpcPackageView
        """
        return self.file.package
    
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
        app.post("message", "　　　　{} [header:{}] 　".format(mart, marh))
        app.post("message", "　{}　　　　{}".format(marr, marl))
        app.post("message", "　　　　{} [footer:{}]　 ".format(marb, marf))
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
