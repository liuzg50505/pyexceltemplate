from wysiwygtemplate.base import WrapedDict
from wysiwygtemplate.templatebase import ExcelTemplate, ExcelProcessor, ExcelTemplateContext, ExpressionEvaluator

# 树模板
class TreeNode:
    def __init__(self) -> None:
        super().__init__()
        self.name = None
        self.exp = None
        self.row = None
        self.col = None
        self.parentname = None
        self.isleaf = False

        self.children = []
        self.parent = None

    @staticmethod
    def parse(cmdtext:str):
        idx1 = cmdtext.find(' ')
        idx2 = cmdtext.find(' ',idx1+1)
        if idx1>-1 and idx2>-1:
            type = cmdtext[:idx1]
            node = TreeNode()
            if type!='node': node.isleaf = True
            nodepart = cmdtext[idx1+1:idx2]
            exp = cmdtext[idx2+1:]
            if ':' in nodepart:
                idx = nodepart.find(':')
                nodename = nodepart[:idx]
                parentnodename = nodepart[idx+1:]
                node.parentname = parentnodename
            else:
                nodename = nodepart
            node.name = nodename
            node.exp = exp
            return node

class TreeNodeContext:
    def __init__(self) -> None:
        super().__init__()
        self.ctx = None
        self.treeNode:TreeNode = None
        self.childContext:{str,TreeNodeContext} = {}
        self.parent = None
        self.count = 0
        self.offset = 0


class TreeTemplate(ExcelTemplate):

    def __init__(self, excelProcessor: ExcelProcessor, context: ExcelTemplateContext, evaluator:ExpressionEvaluator) -> None:
        super().__init__(excelProcessor, context)
        self.context = context
        self.evaluator = evaluator
        self.treemap = {}

    def parseTreeNode(self, row, col)->TreeNode or None:
        v = self.excelProcessor.readCellValue(row, col)
        if not isinstance(v, str): return None

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
            cmdtext = text[pos[0]+len(startTag):pos[1]-len(endTag)]
            if cmdtext.startswith('node '):
                reserveArray += pos
        reserveArray.append(len(text))

        newtext = ''
        for i in range(len(reserveArray)//2):
            a = reserveArray[2*i]
            b = reserveArray[2*i+1]
            newtext += text[a:b]
        self.excelProcessor.writeCellValue(row, col, newtext)

        if len(poslist)==0: return None
        node = TreeNode.parse(text[poslist[0][0]+len(startTag):poslist[0][1]-len(endTag)])
        node.row = row
        node.col = col
        return node

    def copyRow(self, row):
        rowdata = []
        for i in range(self.mincol, self.maxcol + 1):
            v = self.excelProcessor.readCellValue(row, i)
            rowdata.append(v)
        return rowdata

    def preprocess(self):
        minrow, maxrow, mincol, maxcol = self.excelProcessor.bound()
        self.minrow = minrow
        self.maxrow = maxrow
        self.mincol = mincol
        self.maxcol = maxcol

        for row in range(minrow, maxrow + 1):
            rownodes = []
            rownodecols = []
            for col in range(mincol, maxcol +1):
                node = self.parseTreeNode(row ,col)
                if node is not None:
                    self.treemap[(row, col)] = node
                    rownodes.append(node)
                    rownodecols.append(col)
            for nodecol, node in enumerate(rownodes):
                node.children = [i for i in rownodes if i.parentname==node.name]
                for c in node.children:
                    c.parent = node

    def calcTreeNodeContext(self, node: TreeNode, ctx):
        exp = node.exp
        datalist = self.evaluator.evaluate(exp ,ctx)
        if node.isleaf:
            nodectx = TreeNodeContext()
            nodectx.ctx = datalist
            nodectx.treeNode = node
            return [nodectx]
        result = []
        for dataitem in datalist:
            nodectx = TreeNodeContext()
            result.append(nodectx)
            nodectx.ctx = dataitem
            nodectx.treeNode = node
            for childnode in node.children:
                childctx = WrapedDict(ctx)
                childctx.add(node.name, dataitem)
                childnodectxlist = self.calcTreeNodeContext(childnode, childctx)
                nodectx.childContext[childnode.name] = childnodectxlist
                for i in childnodectxlist:
                    i.parent = nodectx
        return result

    def calcTreeNodeContextCount(self, nodeContext:TreeNodeContext, offset:int):
        nodeContext.offset = offset
        if len(nodeContext.childContext)==0:
            nodeContext.count = 1
            return 1
        counts = []
        for name,children in nodeContext.childContext.items():
            count = 0
            for child in children:
                count += self.calcTreeNodeContextCount(child, offset+count)
            counts.append(count)
        nodeContext.count = max(counts)
        return nodeContext.count

    def process(self):
        self.preprocess()
        rowrootnodes = []
        for row in range(self.minrow, self.maxrow+1):
            rownodes = []
            for col in range(self.mincol, self.maxcol+1):
                if (row, col) in self.treemap:
                    rownodes.append(self.treemap[(row, col)])
            if len(rownodes)==0: continue

            rootnodes = [i for i in rownodes if i.parent is None]
            rowrootnodes.append([row, rootnodes])

        offsetrows = 0
        for row, rootnodes in rowrootnodes:
            row += offsetrows
            rootnodecontexts = []
            counts = []
            for rootnode in rootnodes:
                count = 0
                rootnodecontextlist = self.calcTreeNodeContext(rootnode, self.context.getCellContext(row, self.mincol))
                rootnodecontexts.append(rootnodecontextlist)
                for rootnodecontext in rootnodecontextlist:
                    self.calcTreeNodeContextCount(rootnodecontext, count)
                    count += rootnodecontext.count
                counts.append(count)
            count = max(counts)

            vtable = []
            vcontextlst = []
            mergedlist = set()
            for i in range(count):
                vtable.append(self.copyRow(row))
                vcontextlst.append({})

            for rootnodecontext in rootnodecontexts:
                self.getDataContext(rootnodecontext, vcontextlst, mergedlist)

            self.excelProcessor.insertRowBefore(vtable, row)
            self.excelProcessor.deleteRow(row+len(vtable))
            for i, ctx in enumerate(vcontextlst):
                for col in range(self.mincol, self.maxcol+1):
                    self.context.setCellContext(row + i, col, ctx)

            for rowoffset, col, rc in mergedlist:
                self.excelProcessor.mergeCell(row+rowoffset, row+rowoffset+rc-1, col, col)

            offsetrows += count-1

    def getDataContext(self, nodecontextlist, vcontextlst, mergedlist):
        for nodecontext in nodecontextlist:
            rowoffset = nodecontext.offset
            col = nodecontext.treeNode.col
            count = nodecontext.count
            if count>1: mergedlist.add((rowoffset, col, count))
            for row in range(nodecontext.count):
                vcontextlst[row+nodecontext.offset][nodecontext.treeNode.name]=nodecontext.ctx
            for name,childnodecontextlist in nodecontext.childContext.items():
                self.getDataContext(childnodecontextlist, vcontextlst, mergedlist)
