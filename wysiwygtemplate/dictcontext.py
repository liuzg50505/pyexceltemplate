from wysiwygtemplate.templatebase import ExcelTemplateContext

class DictExcelTemplateContext(ExcelTemplateContext):
    def __init__(self, defaultContext) -> None:
        super().__init__()
        self.defaultContext = defaultContext
        self.cellContext = {}

    def getCellContext(self, row, col):
        if (row, col) in self.cellContext:
            return self.cellContext[(row, col)]
        return self.defaultContext

    def setCellContext(self, row, col, context):
        self.cellContext[(row, col)] = context
