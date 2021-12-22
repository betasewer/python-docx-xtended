#!/usr/bin/env python3
# encoding: utf-8

"""
 *** How to use lxml.etree *** 

 elem = etree()
 
 子要素をめぐる
 for child in elem:

 アトリビュートをめぐる
 for attrname, attrval in elem.attrib.items():

 テキストを取得する
 elem.text
 
 要素を削除する
 parent.remove(child)
"""

from lxml import etree
from docxx.oxml.ns import nsmap
from copy import deepcopy

#
def get_element(proxy_or_element):
    if etree.iselement(proxy_or_element):
        return proxy_or_element
    elif hasattr(proxy_or_element, "element"):
        return proxy_or_element.element
    elif hasattr(proxy_or_element, "_element"):
        return proxy_or_element._element
    else:
        raise ValueError("No element related")

# 要素の子を取り除く
def remove_element(parent, if_=None):
    if parent is None:
        return False
    pa=get_element(parent)
    rm=False
    for child in pa:
        if if_ is None or if_(child):
            pa.remove(child)
            rm=True
    return rm

# 要素の子を選別しつつコピーする
def clone_element(parent, if_=None):
    pa=get_element(parent)
    cpa=deepcopy(pa)
    if if_ is not None:
        remove_element(cpa, if_=if_.not_)
    return cpa

# 要素から子を取り除き、ほかの要素に追加する
def insert_move_element(destparent, parent, if_=None, copying=False):
    if None in (parent, destparent):
        return False
    destparent = get_element(destparent)
    for child in get_element(parent):
        if if_ is None or if_(child):
            if copying: 
                child = deepcopy(child)
            destparent.append(child)

# 元要素を残す
def insert_copy_element(destparent, parent, if_=None):
    insert_move_element(destparent, parent, if_, copying=True)
            
# 子の中から要素を探す
def find_element(parent, if_=None):
    if parent is None:
        return None
    for child in get_element(parent):
        if if_ is None or if_(child):
            return child
    return None

# 子の中から要素をすべて探す
def find_all_element(parent, if_=None):
    if parent is None:
        return []
    rets = []
    for child in get_element(parent):
        if if_ is None or if_(child):
            rets.append(child)
    return rets

# 後ろに要素を追加する
def insert_element_next(dest, src):
    pt = get_element(dest)
    parent = pt.getparent()
    if parent is None:
        raise ValueError("")
    i = parent.index(pt)
    parent.insert(i+1, get_element(src))
    
# 前方に要素を追加する
def insert_element_prev(dest, src):
    pt = get_element(dest)
    parent = pt.getparent()
    if parent is None:
        raise ValueError("")
    i = parent.index(pt)
    parent.insert(i, get_element(src))

#
# 名前空間
#
# 要素が所定のタグ名、属性をそなえているか
#   nsmap = [namepsaces]
def match_element(element, name, **attrs):
    element = get_element(element)

    # 名前空間を指定できる
    nsmap = attrs.pop("nsmap", docx_ns)
    if not is_element_clark_name(name):
        name = nsmap.qname(name)
    if element.tag != name:
        return False
    
    # アトリビュートを判定
    for attrname, attrval in attrs.items():
        if attrname not in element.attrib:
            return False
        val = element.attrib[attrname]
        if attrval != val:
            return False
    return True
    

class query():
    """
    query(lxml.etree.element): 
        element と同一オブジェクトかを判定
    query("tagname1", "tagname2" ... "attrname"="arttrval", ...):
        tagname1, 2...のうちいずれか一つのタグ名を持ち、全てのattrname=attrval...が成立するかを判定
    """
    def __init__(self, toparg, *args, **kwargs):
        if isinstance(toparg, str):
            q = lambda x: query.match_tag(x, toparg, *args, **kwargs)
        else:
            q = lambda x: query.match_same(x, toparg)
        self._q = q
        self._sign = True
        
    @classmethod
    def match_tag(cls, obj, *names, **attrs):
        return any(match_element(obj, name, **attrs) for name in names)
    
    @classmethod
    def match_same(cls, obj, right):
        return get_element(obj) is get_element(right)
    
    @property
    def not_(self):
        self._sign = False
        return self
    
    def __call__(self, obj):
        return self._q(obj) == self._sign


class ElementChildren():
    """
    ElementChildren(p, lxml.etree.element): 
        element と同一オブジェクトかを判定
    ElementChildren(p, "tagnames"..., "attrname"="arttrval"...):
        tagnamesのうちいずれか一つのタグ名を持ち、全てのattrname=attrvalが成立するかを判定
    """
    def __init__(self, parent, *args, **kwargs):
        self._parent = get_element(parent)
        top = args[0] if len(args)>0 else None
        if etree.iselement(top):
            q = lambda x: ElementChildren.match_same(x, top)
            one = True
        else:
            q = lambda x: ElementChildren.match_tag(x, *args, **kwargs)
            one = False
        self._q = q
        self._sign = True
        self._one = kwargs.setdefault("one", one)
        
    @classmethod
    def match_tag(cls, obj, *names, **attrs):
        return any(match_element(obj, name, **attrs) for name in names)
    
    @classmethod
    def match_same(cls, obj, right):
        return get_element(obj) is get_element(right)
    
    @property
    def not_(self):
        self._sign = False
        return self
    
    def match(self, elem):
        return self._q(elem) == self._sign

    #
    # 子の中から要素を探す
    #
    def get_parent(self):
        return self._parent

    def children(self):
        if self._parent is None:
            return None
        for child in get_element(self._parent):
            if self._q is self._q(child):
                yield child
                if self._one:
                    break

    def find(self):
        return next(self.children(), None)

    # 要素の子を取り除く
    def remove(self):
        rm=False
        for child in list(self.children()):
            self._parent.remove(child)
            rm=True
        return rm
    
    # 要素から子を取り除き、ほかの要素に追加する
    def move_to(self, destparent):
        if None in (self._parent, destparent):
            return False
        destparent = get_element(destparent)
        for child in self.children():
            destparent.append(child)
        return True

    # 元要素を残す
    def copy_to(self, destparent):
        if None in (self._parent, destparent):
            return False
        destparent = get_element(destparent)
        for child in self.children():
            child = deepcopy(child)
            destparent.append(child)
        return True

#
def same_element(left, right):
    if left is None or right is None:
        return False
    l = get_element(left)
    r = get_element(right)
    return l is r

#
#
# 
class namespaces:
    def __init__(self, *maps):
        self.nsmap = {}
        for mp in maps:
            self.nsmap.update(mp)
        self.pfxmap = dict((value, key) for key, value in self.nsmap.items())

    # 表示用に、名前空間を接頭辞に戻す
    def short_name(self, qname):
        if "}" in qname:
            ns, localname = qname.split("}")
            ns = ns.lstrip("{")
            return "{}:{}".format(self.pfxmap.get(ns,ns), localname)
        return qname
    
    def qname(self, name):
        if ":" not in name:
            raise ValueError("Specify namespace: '{}'".format(name))
        pfx, localname = name.split(":")
        uri = self.nsmap.get(pfx, "")
        return '{%s}%s' % (uri, localname)

# ワードの名前空間
docx_ns = namespaces({
    'msw2010' : ('http://schemas.microsoft.com/office/word/2010/wordml'),
}, nsmap)

#
def is_element_clark_name(name):
    parts = name.split("}")
    if len(parts)>1:
        return True
    return False

#
# 要素をプリントする
# printer, interruptedをカスタマイズ可能
#
class ElementPrinter():
    def __init__(self, ns, indent=4):
        self.ns = ns
        self.ind = indent
    
    def dive(self, elem, level=0):
        tag = self.ns.short_name(elem.tag)
         
        contents = []
        if elem.text:
            contents.append('"{}"'.format(elem.text))
            
        for attrname, attrval in elem.attrib.items():
            attrname = self.ns.short_name(attrname)
            contents.append("{} = {}".format(attrname, attrval)) 
    
        lines = []
        if len(contents)==0 and len(elem)==0:
            lines.append("<{}>".format(tag))
        else:
            lines.append("<{}> {}".format(tag, ", ".join(contents)))
            for child in elem:
                if self.interrupted():
                    return
                sublines = self.dive(child, level+1)
                lines.append(sublines)
            #lines.append("}")
        return lines
    
    def walk(self, lines, level=0):
        for l in lines:
            ind = " "*self.ind*level
            if isinstance(l, str):
                self.printer(ind+l)
            elif isinstance(l, list):
                self.walk(l, level+1)
        
    def __call__(self, elem):
        lines = self.dive(elem)
        self.walk(lines, 0)
        
    def printer(self, s):
        print(s)
        
    def interrupted(self):
        return False
    
# デフォルト設定でdocxのxmlを表示する
def print_element(elem, nsmap=None, indent=4):
    nsmap = nsmap or docx_ns
    prn = ElementPrinter(nsmap, indent)
    prn(get_element(elem))
    