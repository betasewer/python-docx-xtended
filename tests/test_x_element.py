
import pytest
import os
from lxml import etree
from docxx.shared import ElementProxy, Parented
from docxx.element import (
    get_element, remove_element, move_element, copy_element, 
    find_element, match_element, query, dup_element_previous, dup_element_next,
    namespaces,
    ElementPrinter, print_element
)

@pytest.fixture
def testtree():
    from docxx.oxml.ns import qn
    root = etree.Element(qn("w:body"))
    #
    p1 = etree.SubElement(root, qn("w:p"))
    # #
    child1 = etree.SubElement(p1, qn("w:r"))
    child1.text = "Apple"
    child2 = etree.SubElement(p1, qn("w:r"))
    child2.text = "Banana"
    child3 = etree.SubElement(p1, qn("w:r"))
    child3.text = "Orange"
    #
    sec = etree.SubElement(root, qn("w:section"))
    # #
    dir = etree.SubElement(sec, qn("w:direction"))
    dir.set("orientation", "vertical")
    grid = etree.SubElement(sec, qn("w:grid"))
    grid.set("x", "20")
    grid.set("y", "50")
    return root

class ElemProxy(ElementProxy):
    pass
    
class SomeParent(Parented):
    pass
    
def _test_x_printtree(testtree):
    for l in etree.tostring(testtree, pretty_print=True).splitlines():
        print(l)
    print(type(testtree))
    assert False

#
#
#
def test_x_getelement(testtree):
    assert get_element(testtree) is testtree
    assert get_element(ElemProxy(testtree)) is testtree

def test_x_matchelement(testtree):
    from docxx.oxml.ns import nsmap
    assert match_element(testtree, "w:body")
    assert match_element(testtree.find("w:section/w:direction", nsmap), "w:direction", orientation="vertical")
    assert not match_element(testtree.find("w:section/w:grid", nsmap), "w:grid", x="0", y="50")

def test_x_query(testtree):
    from docxx.oxml.ns import nsmap, qn
    assert query(testtree)(testtree)
    assert query("w:body")(testtree)
    assert query(qn("w:body"))(testtree)
    assert query("w:grid", x="20")(testtree.find("w:section/w:grid", nsmap))

def test_x_removeelement(testtree):
    assert remove_element(testtree, query("w:p"))
    assert not remove_element(testtree, query("w:grid"))
    
def test_x_findelement(testtree):
    from docxx.oxml.ns import nsmap
    assert find_element(testtree.find("w:section", nsmap), query("w:grid", x="20")) is testtree.find("w:section/w:grid", nsmap)
    assert find_element(testtree.find("w:p", nsmap), query("w:r")) is not None
    assert find_element(testtree, query("w:body")) is None

def test_x_copyelement(testtree):
    from docxx.oxml.ns import qn, nsmap
    root2 = etree.Element(qn("w:body"))
    copy_element(root2, testtree, query("w:p"))
    move_element(root2, testtree, query("w:section"))
    assert len(root2.find("w:p", nsmap)) == 3
    assert find_element(root2.find("w:section", nsmap), query("w:grid", x="20")) is not None
    assert len(testtree.find("w:p", nsmap)) == 3
    assert find_element(testtree, query("w:section")) is None

def test_x_dupelement(testtree):
    from docxx.oxml.ns import nsmap
    elem = find_element(testtree.find("w:p", nsmap), query("w:r"))
    assert elem is not None
    elem4 = dup_element_next(elem)
    elem4.text = "Onion"
    elem5 = dup_element_previous(elem)
    elem5.text = "Tomato"
    assert [x.text for x in testtree.find("w:p", nsmap)] == ["Tomato", "Apple", "Onion", "Banana", "Orange"]

def test_x_namespaces():
    from docxx.oxml.ns import qn
    from docxx.element import docx_ns
    assert docx_ns.qname("w:p") == qn("w:p")
    assert docx_ns.short_name(qn("w:p")) == "w:p"

def test_x_elementprint(testtree):
    print_element(ElemProxy(testtree))
