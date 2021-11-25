from docxx.document import Document
from docxx.text.paragraph import Paragraph
from docxx.text.run import Run, same_run
from docxx.text.runlist import RunList, RunRange, runlist
from docxx.element import print_element
from docxx.oxml import parse_xml
from docxx import open_docx

import pytest

@pytest.fixture
def newp():
    xml = r'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:o="urn:schemas-microsoft-com:office:office" 
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
        xmlns:v="urn:schemas-microsoft-com:vml" 
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
        xmlns:w10="urn:schemas-microsoft-com:office:word" 
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" 
        xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" 
        xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" 
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" 
        xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" 
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" 
        mc:Ignorable="w14 wp14">
    <w:body>
        <w:p>
            <w:pPr>
                <w:rPr></w:rPr>
            </w:pPr>
        </w:p>
    </w:body>
    </w:document>
    '''.encode("utf-8")
    d = parse_xml(xml)
    doc = Document(d, None)
    return doc.paragraphs[0]

@pytest.fixture
def newruns(newp):
    return runlist(newp)
    
def ranges_to_texts(ranges):
    return sum([[x.text for x in x.runs()] for x in ranges], [])

#
#
#
def test_x_runs_insert(newruns):
    newruns.insert(newruns.end, "BBB")
    newruns.insert(newruns.end, "CCC")
    newruns.insert(newruns.end, "DDD")
    newruns.insert(newruns.end, "EEE")
    newruns.insert(newruns.begin, "111")
    newruns.insert(newruns.begin, "222")
    newruns.insert(newruns.begin, "333")
    newruns.insert(newruns.begin, "444")
    assert newruns.text == "444333222111BBBCCCDDDEEE"

def test_x_runs_getruns(newruns):
    newruns.insert(newruns.end, "First Run")
    li = list(newruns.runs(newruns.at(0), newruns.at(1)))
    assert len(li) == 1
    assert same_run(li[0], newruns[0])
    li2 = list(newruns.runs())
    assert len(li2) == 1
    assert same_run(li2[0], li[0])

def test_x_make_view(newruns):
    # 空のrange
    emp = newruns.range(newruns[0], newruns[0])
    assert same_run(emp.begin, newruns.at(0))
    assert same_run(emp.end, newruns.at(0))
    assert emp.text ==  ""
    assert len(emp) ==  0
    assert emp.empty()
    # 要素1つのrange
    newruns.append("AAA")
    all = newruns.range()
    assert all.begin.text == "AAA"
    assert all.tail.text == "AAA"
    assert [x.text for x in all.runs()] == ["AAA"]
    assert [x[0].text for x in all.textranges()] == ["AAA"]
    assert all.text == "AAA"
    assert len(all) == 1

    
@pytest.fixture
def numberruns(newruns):
    datas = ["111","222","333","444","555"]
    for data in datas:
        newruns.insert(newruns.end, data)
    return (newruns, datas)

def test_x_runs_iterate(numberruns):
    newruns, datas = numberruns
    for r in newruns.runs(newruns.begin, newruns.end):
        assert datas[0] == r.text
        datas.pop(0)
        
def test_x_runs_rev_iterate(numberruns):
    newruns, datas = numberruns
    for r in newruns.runs_reversed(newruns.begin, newruns.end):
        assert datas[-1] == r.text
        datas.pop()
        
def test_x_runs_view_iterate(numberruns):
    newruns, datas = numberruns
    for r in newruns.range().runs():
        assert datas[0] == r.text
        datas.pop(0)
        
def test_x_runs_view_shift(numberruns):
    newruns, datas = numberruns
    view = newruns.range()
    view.shift_head(1)
    view.shift_tail(2)
    assert view.text == "222333"
    
    datas = datas[1:]
    for r in view.runs():
        assert datas[0] == r.text
        datas.pop(0)
        
@pytest.fixture
def textruns(newruns):
    texts = ["私は、","明日の午後に","、","目覚め、そして","バナナの、種を、植えます", "。"]
    for t in texts:
        newruns.insert(newruns.end, t)
    splittexts = ["私は","、","明日の午後に","、","目覚め","、","そして","バナナの","、","種を","、","植えます", "。"]
    return (newruns, texts, splittexts)

def test_x_runs_split(textruns):
    newruns, _texts, _results = textruns
    rs = newruns.split("、")
    runs = ranges_to_texts(rs)
    assert runs == ["私は","明日の午後に","目覚め","そして","バナナの","種を","植えます","。"]
    
def test_x_runs_split_max(textruns):
    newruns, _texts, _results = textruns
    rs = newruns.split("、", 2)
    runs = ranges_to_texts(rs)
    assert runs == ["私は","明日の午後に", "、", "目覚め、そして","バナナの、種を、植えます", "。"]

def test_x_runs_rsplit_max(textruns):
    newruns, _texts, _results = textruns
    rs = newruns.rsplit("、", 2)
    runs = ranges_to_texts(rs)
    assert runs == ["私は、","明日の午後に", "、", "目覚め、そして","バナナの、種を、植えます", "。"]

def test_x_runs_separate(textruns):
    newruns, texts, spl = textruns
    rs = newruns.separate("、")
    for run, data in zip(rs, spl):
        assert run.text == data

#
@pytest.fixture(params=[
    (lambda x:len(x.text)>6, ("私は、明日の午後に、", "目覚め、そしてバナナの、種を、植えます", "。")),
    (lambda x:"、" in x.text, ("私は、", "明日の午後に", "、目覚め、そしてバナナの、種を、植えます", "。")),
])
def textpartitions(textruns, request):
    newruns, _, _ = textruns
    return (newruns, request.param[0], request.param[1])

def test_x_runs_partitions(textpartitions):
    newruns, fn, parts = textpartitions
    for (keyval, range), part in zip(newruns.partitions(fn), parts):
        assert range.text == part

#
def test_x_run_split_separations(newruns):
    newruns.insert(newruns.end, "AAA-BBB-CCC-DDD")
    assert newruns.text == "AAA-BBB-CCC-DDD"

    run = newruns.tail
    breaks = []
    for pos, t in enumerate(run.text):
        if t == "-":
            breaks.append(pos)
            breaks.append(pos+1)

    texts = []
    for run in newruns.run_split_separations(run, breaks):
        texts.append(run.text)
    
    assert texts == ["AAA", "-", "BBB", "-", "CCC", "-", "DDD"]


#
@pytest.fixture(params=[
    ("そしてバナナ", (4, 3)),
    ("私は", (0, 2)),
    ("午後に、目覚め", (3, 3)),
    ("。", (0, 1)),
])
def searchtextruns(textruns, request):
    newruns, _, _ = textruns
    return (newruns, request.param[0], request.param[1])

def test_x_runs_search_text(searchtextruns):
    newruns, srchtext, textposes = searchtextruns
    resultrange = newruns.search_text(srchtext)
    assert resultrange is not None
    assert resultrange.text == srchtext
    assert (resultrange.textbeg, resultrange.textend) == textposes

#
#
#
#
def test_x_runs_misc(newruns):
    runs = newruns
    print("[part.runs]")
    for r in runs.range(runs.at(0), runs.at(4)).runs():
        print(r.text)
        
    print("[part.text]")
    print(runs.range(runs.at(0), runs.at(4)).text)
    
    print("[insert]:")
    #runs.insert(retrange.tail, "！")
    #runs.insert(retrange.tail, "――うんこちんちん――")
    #retrange.tail.text += "それにつけても"

    runs.insert(runs.begin, "　あぁ！？")
    runs.insert(runs.begin, "           ")
    
    print("[lstrip]:")
    runs.lstrip()
    runs.range().display()
   

   
