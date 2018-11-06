from wysiwygtemplate.templatebase import ExpressionEvaluator


class PyEvaluator(ExpressionEvaluator):
    def __init__(self, printException=True, throwException=False) -> None:
        super().__init__()
        self.printException = printException
        self.throwException = throwException

    def evaluate(self, exp,context):
        try:
            return eval(exp, {},context)
        except Exception as e:
            if self.printException: print(e)
            if self.throwException: raise e
            return ""

class EmbeddedPyEvaluator(ExpressionEvaluator):
    def __init__(self, evaluator:ExpressionEvaluator) -> None:
        super().__init__()
        self.evaluator = evaluator
        self.startTag = '{{'
        self.endTag = '}}'

    def evaluate(self, exp, context):
        startTag = self.startTag
        endTag = self.endTag
        text = exp

        poslist = []
        pos1 = text.find(startTag)
        pos2 = text.find(endTag, pos1 + len(startTag))
        while pos1 > -1 and pos2 > -1 :
            poslist.append((pos1, pos2))
            pos1 = text.find(startTag, pos2)
            pos2 = text.find(endTag, pos1 + len(startTag))

        cur = 0
        values = []
        for start, end in poslist:
            values.append(exp[cur: start])
            v = self.eval(exp[start+len(startTag):end], context)
            values.append(str(v))
            cur = end+len(endTag)

        if cur<len(exp): values.append(exp[cur:])

        return ''.join(values)

    def eval(self, exp, context):
        return self.evaluator.evaluate(exp, context)

if __name__=='__main__':
    text = "name={{age}},age={{name}},{{birth(4)}}"
    ev = EmbeddedPyEvaluator(PyEvaluator())
    r = ev.evaluate(text, {
        'age':22,
        'name':'2sdf',
        'birth': lambda f: f+2
    })
    print(r)

