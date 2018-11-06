from wysiwygtemplate.templatebase import ExcelTemplate, ExcelProcessor, ExcelTemplateContext, ExpressionEvaluator


class ExcelRelpaceTemplate(ExcelTemplate):

    def __init__(self, excelProcessor: ExcelProcessor, context: ExcelTemplateContext, expEvaluator: ExpressionEvaluator) -> None:
        super().__init__(excelProcessor, context)
        self.expEvaluator = expEvaluator

    def setExcelContext(self, context: ExcelTemplateContext):
        self.context = context

    def process(self):
        minrow,maxrow,mincol,maxcol = self.excelProcessor.bound()
        for row in range(minrow, maxrow+1):
            for col in range(mincol, maxcol+1):
                v = self.excelProcessor.readCellValue(row, col)
                if isinstance(v, str):
                    cellcontext = self.context.getCellContext(row, col)
                    vv = self.expEvaluator.evaluate(v, cellcontext)
                    self.excelProcessor.writeCellValue(row, col, vv)

    def export(self, path):
        self.excelProcessor.save(path)
