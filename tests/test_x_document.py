import os
from docxx import open_docx
from docxx.text.runlist import runlist

def test_open():
    doc = open_docx("C:\\codes\\python\\python-docx-xtended\\tests\\test_x_open.docx")
    p = doc.document.paragraphs[0]
    assert runlist(p).text == "TEST SUCCESS"

def test_edit_and_open():
    doc = open_docx()
    p = doc.document.add_paragraph()
    r1 = p.add_run("TEST2")
    r2 = p.add_run("-")
    r3 = p.add_run("SUCCESS")
    assert runlist(doc.document.paragraphs[0]).text == "TEST2-SUCCESS"
    
def test_misc():
    from docxx.enum.text import WD_COLOR
    doc = open_docx()
    # 文字スタイル
    p = doc.document.add_paragraph()
    r1 = p.add_run("STYLE")
    r1.bold = True
    r2 = p.add_run("-")
    r3 = p.add_run("TEST")
    r3.italic = True
    p = doc.document.add_paragraph()
    r1 = p.add_run("ここは")
    r2 = p.add_run("傍点")
    r2.font.em = True
    r3 = p.add_run("で、ここは")
    r4 = p.add_run("太字")
    r4.font.bold = True
    r5 = p.add_run("です！")
    r5.font.highlight_color = WD_COLOR.GREEN
    # ルビ
    p = doc.document.add_paragraph()
    p.add_run("お前は")
    r1 = p.add_run("何故")
    r1.add_ruby("なぜ")
    p.add_run("これをしているのか？")
    # コメント
    

    doc.save("C:\\codes\\python\\python-docx-xtended\\tests\\test_x_document_misc.docx")
    
def test_runs():
    pass
    