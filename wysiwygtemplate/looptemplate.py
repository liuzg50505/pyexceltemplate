from copy import copy

from wysiwygtemplate.base import WrapedDict
from wysiwygtemplate.templatebase import ExcelTemplate, ExcelProcessor, ExcelTemplateContext, ExpressionEvaluator


class ExcelArchetectureCommand:
    pass

class LoopArchetectureCommand(ExcelArchetectureCommand):
    collectionVarName: str
    loopVarName: str
    refVarName: str
    rows:int = 0

    def __init__(self, collectionVarName: str, loopVarName: str) -> None:
        super().__init__()
        self.collectionVarName = collectionVarName
        self.loopVarName = loopVarName
        self.refVarName=None

class ExcelArchetectureTemplate(ExcelTemplate):
    def __init__(self, excelProcessor: ExcelProcessor, context: ExcelTemplateContext, evaluator:ExpressionEvaluator) -> None:
        super().__init__(excelProcessor, context)
        self.evaluator = evaluator
        self.archmap = {}

    def preprocess(self):
        minrow, maxrow, mincol, maxcol = self.excelProcessor.bound()
        self.minrow = minrow
        self.maxrow = maxrow
        self.mincol = mincol
        self.maxcol = maxcol
        for row in range(minrow, maxrow + 1):
            for col in range(mincol, maxcol +1):
                archs = self.parseCellArchetectures(row ,col)
                self.archmap[(row, col)] = archs

        allloops = []
        for archs in self.archmap.values():
            for arch in archs:
                allloops.append(arch)

        for loop1 in allloops:
            for loop2 in allloops:
                if loop1==loop2: continue
                if loop1.refVarName==loop2.loopVarName:
                    loop1.dependsOn = loop2

    def archetectureRowNumbers(self):
        rows = set([r for r,c in self.archmap.keys() if len(self.archmap[(r,c)])>0])
        return sorted(rows, reverse= False)

    def getRowArchetectures(self, row):
        archs = [self.archmap[(r,c)] for r,c in self.archmap.keys() if r==row]
        return archs

    def getRowRootArcheTecture(self, row):
        archs = self.getRowArchetectures(row)
        tops = [i for i in archs if i.dependsOn is None]
        if len(tops)==0: return None
        return tops[0]

    def parseCellArchetectures(self, row, col):
        v = self.excelProcessor.readCellValue(row, col)
        if not isinstance(v, str): return []

        startTag = '<<'
        endTag = '>>'
        text = str(v)

        poslist = []
        pos1 = text.find(startTag)
        pos2 = text.find(endTag, pos1 + len(startTag))
        while pos1 > -1 and pos2 > -1 :
            poslist.append((pos1, pos2 + len(endTag)))
            pos1 = text.find(startTag, pos2)
            pos2 = text.find(endTag, pos1 + len(startTag))

        reserveArray = [0]
        for pos in poslist:
            reserveArray += pos
        reserveArray.append(len(text))

        newtext = ''
        for i in range(len(reserveArray)//2):
            a = reserveArray[2*i]
            b = reserveArray[2*i+1]
            newtext += text[a:b]
        self.excelProcessor.writeCellValue(row, col, newtext)

        commands = []
        for start, end in poslist:
            commands.append(text[start+len(startTag): end-len(endTag)])

        archlist = []
        for cmd in commands:
            if cmd.startswith('loop '):
                arr =  cmd.split(' ')
                collectionVarName = arr[1]
                loopVarName = arr[2]
                arch = LoopArchetectureCommand(collectionVarName, loopVarName)
                if len(arr)==4:
                    refvar = arr[3]
                    arch.refVarName = refvar
                archlist.append(arch)
        return archlist

    def copyRow(self, row):
        rowdata = []
        for i in range(self.mincol, self.maxcol+1):
            v = self.excelProcessor.readCellValue(row, i)
            rowdata.append(v)
        return rowdata

    def process(self):
        self.preprocess()
        rowdata = []
        offsetrows = 0
        for row in self.archetectureRowNumbers():
            vtable, varchtable, vcontexttable = self.processLoopArchetectureTable(row)
            rowdata.append([row, vtable, varchtable, vcontexttable])
            self.excelProcessor.insertRowBefore(vtable, row+offsetrows)
            self.excelProcessor.deleteRow(row+offsetrows+len(vtable))
            for i, vcontextrow in enumerate(vcontexttable):
                for col in range(self.mincol, self.maxcol+1):
                    ctx = vcontextrow[col-self.mincol]
                    self.context.setCellContext(row+ offsetrows + i, col, ctx)

            offsetrows += len(vtable)-1

    def processLoopArchetectureTable(self, row):
        vtablerow = []
        varchrow = []
        vcontextrow = []

        for col in range(self.mincol, self.maxcol+1):
            vtablerow.append(self.excelProcessor.readCellValue(row, col))
            if (row, col) in self.archmap:
                varchrow.append(self.archmap[(row, col)])
            else:
                varchrow.append([])
            vcontextrow.append(WrapedDict(self.context.getCellContext(row,col)))

        vtable = [vtablerow]
        varchtable= [varchrow]
        vcontexttable = [vcontextrow]

        rowidx = 0
        total = len(vtable)
        while rowidx<total:
            r = self.processLoopArchetecture(vtable, varchtable, vcontexttable, rowidx)
            if r:
                rowidx=0
                total = len(vtable)
            else: rowidx+=1

        return vtable, varchtable, vcontexttable

    def processLoopArchetecture(self, vtable:[[any]], varchtable:[[LoopArchetectureCommand or any]], vcontexttable:[[dict]], vrowindex:int):
        vrow = vtable[vrowindex]
        vrowcontext = vcontexttable[vrowindex]
        varchrow = varchtable[vrowindex]
        rootArchetecture = None
        rootArchetectureIdx = -1
        for i,arch in enumerate(varchrow):
            if arch is None: continue
            if len(arch)==0: continue
            for item in arch:
                if isinstance(item, LoopArchetectureCommand):
                    if item.refVarName is None:
                        rootArchetecture = item
                        rootArchetectureIdx = i
                        break

        if rootArchetecture is None: return False
        varname = rootArchetecture.loopVarName
        exp = rootArchetecture.collectionVarName
        context = vrowcontext[rootArchetectureIdx]

        for item in varchrow:
            for i in item:
                if isinstance(i, LoopArchetectureCommand):
                    if i.refVarName == rootArchetecture.loopVarName:
                        i.refVarName=None

        data = self.evaluator.evaluate(exp, context)
        vtable.pop(vrowindex)
        vcontexttable.pop(vrowindex)
        varchtable.pop(vrowindex)

        for i,v in enumerate(data):
            vrowcp = copy(vrow)
            vrowcontextcp = []
            varchrowcp = [ list(i) for i in varchrow]
            for ctx in vrowcontext:
                d = WrapedDict(ctx)
                vrowcontextcp.append(d)
                d.add(varname, v)
                if rootArchetecture in varchrowcp[rootArchetectureIdx]:
                    varchrowcp[rootArchetectureIdx].remove(rootArchetecture)

            vtable.insert(vrowindex+i, vrowcp)
            varchtable.insert(vrowindex+i,varchrowcp)
            vcontexttable.insert(vrowindex+i, vrowcontextcp)

        return True
