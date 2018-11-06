class ExcelProcessor:
    def __init__(self, xlsfilepath) -> None:
        super().__init__()
        self.xlsfilepath = xlsfilepath

    def bound(self)->tuple:
        pass

    def readCellValue(self, row, col):
        pass

    def writeCellValue(self, row, col, val):
        pass

    def insertRowBefore(self, rowdata, row):
        pass

    def insertRowAfter(self, rowdata, row):
        pass

    def insertColumnBefore(self, coldata, row):
        pass

    def insertColumnAfter(self, coldata, row):
        pass

    def deleteRow(self, row):
        pass

    def mergeCell(self, startRow, endRow, startCol, endCol):
        pass

    def insertTemplate(self, template, context, row, col):
        pass

    def save(self, path):
        pass

class ExcelTemplateContext:
    def getCellContext(self, row, col):
        pass

    def setCellContext(self, row, col, context):
        pass

class ExcelTemplate:
    def __init__(self, excelProcessor:ExcelProcessor, context: ExcelTemplateContext) -> None:
        super().__init__()
        self.excelProcessor = excelProcessor
        self.context = context

    def process(self):
        pass

class ExpressionEvaluator:
    def __init__(self) -> None:
        super().__init__()

    def evaluate(self, exp,context):
        pass
