# encoding: utf-8

"""
RunsView, Runs
"""
from typing import List, Optional, Any, Generator, Sequence, Tuple, List, Callable, Dict
from docxx.text.run import Run, same_run, clone_run

#
class BadIterationError(Exception):
    pass

#
#
class RunRange():   
    #
    # parent : Proxy of parent element
    # beg : None | Proxy of start run element
    # end : None | Proxy of end run element
    # textbeg : None | Index of text begin of start run
    # textend : None | Index of text end of tail run
    # 
    def __init__(self, runlist, beg=None, end=None, textbeg=None, textend=None):
        self._runlist = runlist
        self.begin = beg if beg is not None else runlist.begin
        self.end = end if end is not None else runlist.end
        self.textbeg = textbeg 
        self.textend = textend
     
    def empty(self):
        return same_run(self.begin, self.end)
    
    def length(self):
        count = 0
        for _ in self.runs():
            count += 1
        return count
        
    def __len__(self):
        return self.length()
    
    @property
    def tail(self):
        if self.end is None:
            return self._runlist.at(-1)
        else:
            return self._runlist.get_prev_run(self.end)
            
    #
    def shift_head(self, count):
        if self.begin is None:
            raise BadIterationError()
        self.begin = self._runlist.get_related_run(self.begin, count)
        
    def shift_tail(self, count):
        if self.end is None:
            if count>=0:
                endrun = self._runlist.at(-1)
                count -= 1
            else:
                raise BadIterationError()
        else:
            endrun = self.end
        self.end = self._runlist.get_related_run(endrun, -count)
    
    @property
    def textbegin_index(self):
        if self.textbeg is None:
            return 0
        return self.textbeg
        
    @property
    def textend_index(self):
        if self.textend is None:
            tail = self.tail
            if tail is not None:
                return len(tail.text)
            else:
                return 0
        return self.textend
    
    def runs(self):
        return self._runlist.runs(self.begin, self.end)
        
    # 対象内のランをコピーしてリストを返す
    def clones(self) -> List[Run]:
        copied = []
        for srcrun in self.runs():
            copied.append(clone_run(srcrun))
        return copied
    
    # 巡回
    # テキスト位置と一緒にランを返す
    def textranges(self):
        for run in self.runs():
            tb = None
            if same_run(run, self.begin):
                tb = self.textbeg        
            te = None
            if same_run(run, self.tail):
                te = self.textend            
            yield run, tb, te
    
    def texts(self):
        for run, b, e in self.textranges():
            yield run.text[b:e]

    # テキスト取得
    @property
    def text(self):
        return "".join([t for t in self.texts()])

    # 主にデバッグ用
    def display(self):
        print(" ------------- ")
        for k, r in enumerate(self.runs()):
            print("run {:02}: {}".format(k, r.text))
    
#
#
#
class RunList():
    # parent : Proxy of parent element of runs
    # runadder : 
    def __init__(self, parent, runadder):
        self._parent = parent
        self._runadder = runadder
    
    @property
    def _parent_elem(self):
        return self._parent._element
        
    @property
    def _srcrunlist(self):
        return self._parent._element.r_lst
        
    def _makerun(self, elem):
        return Run(elem, self._parent)
        
    def _addrun(self, text):    
        return self._runadder(self._parent, text)
    
    #
    #
    #    
    def __len__(self):
        return len(self._srcrunlist)
        
    def __iter__(self):
        return self.runs(self.begin, self.end)
        
    def __getitem__(self, key):
        if isinstance(key, int):
            return self.at(key)
        elif isinstance(self, slice):
            return self.range(slice.start, slice.stop)
        else:
            raise ValueError()
    
    def index(self, run):
        return self._srcrunlist.index(run._element)
    
    def at(self, idx):
        try:
            elem = self._srcrunlist[idx]
        except IndexError:
            return None
        return self._makerun(elem)
        
    def get_next_run(self, run, count=1):
        if run is None: raise ValueError("get_next_run: run must be not None")
        elem = run._element
        elemtype = type(elem)
        for _ in range(count):
            while elem is not None:
                elem = elem.getnext()
                if isinstance(elem, elemtype):
                    break
        if elem is not None:
            return self._makerun(elem)
            
    def get_prev_run(self, run, count=1):
        if run is None: raise ValueError("get_prev_run: run must be not None")
        elem = run._element
        elemtype = type(elem)
        for _ in range(count):
            while elem is not None:
                elem = elem.getprevious()
                if isinstance(elem, elemtype):
                    break
        if elem is not None:
            return self._makerun(elem)
    
    def get_related_run(self, run, count):
        if count>0:
            return self.get_next_run(run, count)
        elif count<0:
            return self.get_prev_run(run, -count)
        else:
            return run

    def runs(self, begin=None, end=None):
        run = begin if begin is not None else self.begin
        while not same_run(run, end):
            nextrun = self.get_next_run(run) # yield後にrunをremoveされてもつながるように
            yield run
            run = nextrun
            
    def runs_reversed(self, begin=None, end=None):
        run = end if end is not None else self.end
        while run is not None:
            prevrun = self.get_prev_run(run)
            yield run
            if same_run(run, begin):
                break
            run = prevrun
        
    # 範囲を変更
    def range(self, begin=None, end=None, textbeg=None, textend=None):
        return RunRange(self, begin, end, textbeg, textend)
        
    @property
    def begin(self):
        return self.at(0)
    
    @property
    def tail(self):
        return self.at(-1)
    
    @property
    def end(self):
        return None
        
    #
    # ある文字列を検索し、全体をその位置で再分割する
    #  sep = separator, 文字列可
    #  max = 分割回数
    #  返り値： run, parts = (text, is-separator)
    #
    def runs_separations(self, sep: str, max=-1, range=None) -> Generator[Tuple[Run, List[Tuple[str, bool]]], None, None]:
        range = range or self.range()                    
        hitcnt = max
        for run in self.runs(range.begin, range.end):
            texts = run.text.split(sep, hitcnt)
            hitcnt -= len(texts)-1            
            if len(texts)==1:
                yield run, []
            else:
                newparts = []
                for t in texts:
                    if len(t)>0:
                        newparts.append((t, False))
                    newparts.append((sep, True))
                newparts.pop()
                yield run, newparts
    
    # 分割
    def split(self, sep: str, max=-1, range=None) -> List[RunRange]:
        range = range or self.range()
        runsspl: List[List[Run]] = [[]]
        for run, newparts in self.runs_separations(sep, max, range):
            if len(newparts)>0:
                newrun = run
                for runtext, issep in newparts:
                    newrun = self.insert_run_next(newrun, clone_run(run, text=runtext))
                    if issep:
                        runsspl.append([])
                    else:
                        runsspl[-1].append(newrun)
                self.remove(run)
            else:
                runsspl[-1].append(run)
        
        rets: List[RunRange] = []
        for runlist in runsspl:
            if len(runlist)==0:
                rets.append(self.range())
            else:
                endrun = self.get_next_run(runlist[-1])
                rets.append(self.range(runlist[0], endrun))
        return rets
        
    # splitのようにList[Run]のリストとせず、単にRunを区切ったリストにする。セパレータ文字も削除しない
    def separate(self, sep: str, max=-1, range=None) -> List[Run]:
        runs = []
        for run, newparts in self.runs_separations(sep, max, range):
            if len(newparts)>0:
                newrun = run
                for runtext, _ in newparts:
                    newrun = self.insert_run_next(newrun, clone_run(run, text=runtext))
                    runs.append(newrun)
                self.remove(run)
            else:
                runs.append(run)
        return runs
    
    # 任意の値でランを区分けして返す
    def partitions(self, key: Callable[[Run], Any], range: RunRange=None) -> Generator[Tuple[Any, RunRange], None, None]:
        range = range or self.range()
        curbeg = range.begin
        keyval = None
        for run in self.runs(range.begin, range.end):
            curkeyval = key(run)
            if keyval is not None and keyval != curkeyval:
                yield keyval, self.range(curbeg, run)
                curbeg = run
            keyval = curkeyval
        yield keyval, self.range(curbeg, range.end)
    
    def runs_separations_break(self, breaks: Sequence[Tuple[Run, int]]) -> Optional[RunRange]:
        """
        与えられた位置でランを区切って返す
        Params:
            breaks(Sequence[Tuple[Run, Int]]): 区切り位置を示す[ラン、テキストインデックス]の組のリスト
        Returns:
            RunRange: 区切りが発生した範囲
        """
        runbreaks: Dict[int, Tuple[Run, List[int]]] = {}
        for brkrun, brkpos in breaks:
            key = id(brkrun)
            if key not in runbreaks:
                runbreaks[key] = (brkrun, [])
            runbreaks[key][1].append(brkpos)
        
        newruns: List[Run] = []
        for brkrun, brkposlist in runbreaks.values():
            text: str = brkrun.text
            if not brkposlist:
                continue
            newtextfirst = text[0:brkposlist[0]]

            # テキストを分割する
            newtextothers: List[str] = []
            lastpos = brkposlist[0]
            for brkpos in brkposlist[1:]:
                newtextothers.append(text[lastpos:brkpos])
                lastpos = brkpos
            newtextothers.append(text[lastpos:])

            # 前のランのテキストを縮める
            brkrun.text = newtextfirst
            newruns.append(brkrun)

            # 新しいランを後ろに追加する
            ptrun = brkrun
            for newtext in newtextothers:
                newrun = clone_run(ptrun) # 前のランを複製
                newrun.text = newtext
                self.insert_run_next(ptrun, newrun)
                newruns.append(newrun)
                ptrun = newrun

        if not newruns:
            return None
        return self.range(newruns[0], self.get_next_run(newruns[-1]))

    #
    # テキスト操作（複数ランにまたがっている場合に対応）
    #
    # 検索
    def search_text(self, s: str, range: RunRange=None) -> Optional[RunRange]:
        range = range or self.range()        
        # 先頭から徐々に対象範囲を広げつつ探す
        rng = range
        for tailrun in self.runs(range.begin, range.end):
            endrun = self.get_next_run(tailrun)
            rng = self.range(range.begin, endrun)
            if len(rng)==0:
                continue            
            texthead = rng.text.find(s)
            #print("head sch: ", tailrun.text, view.text, texthead)
            if texthead != -1:
                break
        else:
            return None        
        textend = texthead + len(s)
        
        # 文字列のインデックスを、範囲の先頭基準から、ランの末尾基準に直す
        h = None
        e = None
        offset = len(rng.text)
        for headrun in self.runs_reversed(range.begin, tailrun):
            doffset = offset - len(headrun.text)
            if doffset <= texthead and h is None:
                h = texthead - doffset
            if doffset <= textend and e is None:
                e = textend - doffset
            #print("h index: ", h)
            #print("e index: ", e)
            #print("run text: ", headrun.text)
            offset = doffset
            if None not in (h, e):
                break
        
        # 返り値
        return self.range(headrun, endrun, h, e)
    
    #
    # ランの追加・削除
    #
    # ランを生成し、テキストを割り当て追加する
    def insert(self, ptrun, text, next_=False):
        newrun = self._addrun(text)
        ptrun = ptrun or self.tail
        if next_:
            self.insert_run_next(ptrun, newrun)
        else:
            self.insert_run_prev(ptrun, newrun)
        return newrun
        
    def insert_next(self, ptrun, text):
        return self.insert(ptrun, text, next_=True)
    
    def append(self, text):
        return self.insert(None, text)
        
    # 指定ランの箇所にランを挿入する
    def insert_run_prev(self, ptrun, run):
        if ptrun is None:
            self._parent_elem.append(run._element)
        else:
            parent = ptrun._element.getparent()
            if parent is None:
                raise ValueError("")
            i = parent.index(ptrun._element)
            parent.insert(i, run._element)
        return run

    # 指定ランの箇所にランを挿入する
    def insert_run_next(self, ptrun, run):
        if ptrun is None:
            self._parent_elem.append(run._element)
        else:
            parent = ptrun._element.getparent()
            if parent is None:
                raise ValueError("")
            i = parent.index(ptrun._element)
            parent.insert(i+1, run._element)
        return run

    # ランを挿入し、新しいViewを返す
    #   srcruns : Sequence[Run]
    def insert_runs(self, srcruns, ptrun=None):
        begrun = None
        for srcrun in srcruns:
            self.insert_run_next(ptrun, srcrun)
            ptrun = srcrun
            if begrun is None: 
                begrun = srcrun
        
        if ptrun is not None:
            endrun = self.get_next_run(ptrun)
        else:
            endrun = None
        return self.range(begrun, endrun)
    
    # 指定ランを削除する
    def remove(self, remrun):
        if remrun is None:
            return
        self._parent_elem.remove(remrun._element)
    
    # ビュー内のランのテキストを一つのランにまとめる
    def unify(self, range=None, top=None):
        range = range or self.range()
        top = top or range.begin
        for run, b, e in range.textranges():
            top.text += run.text[b:e]
            self.remove(run)
        return top
    
    # 先頭の文字列を削除する
    def lstrip(self, if_=None, range=None):
        range = range or self.range()
        if_ = if_ or str.isspace
        rems = []
        remend = False
        for run in self.runs(range.begin, range.end):
            text = run.text
            for start, ch in enumerate(text):
                if not if_(ch):
                    run.text = text[start:]
                    remend = True
                    break
            else:
                self.remove(run)            
            if remend:
                break
    
    # 文字列を置き換える
    def replace_placeholders(self, placeholder, *args, range=None):
        newruns = []
        argitr = iter(args)
        for run in self.separate(placeholder, range=range):
            if run.text == placeholder:
                try:
                    arg = next(argitr)
                except GeneratorExit:
                    raise Exception("not enough argument is provided")
                if isinstance(arg, str):
                    run.text = arg
                elif isinstance(arg, Run):
                    newrun = clone_run(arg)
                    self.insert_run_next(run, newrun)
                    self.remove(run)
                    run = newrun
                else:
                    run.text = str(arg)
            newruns.append(run)
        if not newruns:
            return None
        return RunRange(self, newruns[0], self.get_next_run(newruns[-1]))
    
    @property
    def text(self):
        return self.range().text

#
# API
#
def runlist(parentproxy, runadder=None):
    if runadder is None:
        def runadder(paragraph, text):
            return paragraph.add_run(text)
    return RunList(parentproxy, runadder)

    
