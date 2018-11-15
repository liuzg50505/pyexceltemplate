import  unittest

from wysiwygtemplate.dictcontext import DictExcelTemplateContext
from wysiwygtemplate.openpyprocessor import OpenpyXlExcelProcessor
from wysiwygtemplate.pyevaluator import PyEvaluator, EmbeddedPyEvaluator
from wysiwygtemplate.replacetemplate import ExcelRelpaceTemplate
from wysiwygtemplate.treetemplate import TreeTemplate


class DivTest(unittest.TestCase):
    def test_div_001(self):
        processor = OpenpyXlExcelProcessor('xlstemplates/template2.xlsx', 'Sheet1')
        processor.deleteRow(3)
        processor.save('out.xlsx')

    def test_div_002(self):
        processor = OpenpyXlExcelProcessor('xlstemplates/template2.xlsx', 'Sheet1')
        processor.insertRowBefore([
            ['a','a','a','a','a','a','a','a','a',],
            ['a','a','a','a','a','a','a','a','a',],
            ['a','a','a','a','a','a','a','a','a',],
            ['a','a','a','a','a','a','a','a','a',],
        ],4)
        processor.save('out1.xlsx')

    def test_div_003(self):
        processor = OpenpyXlExcelProcessor('xlstemplates/template3.xlsx', 'Sheet1')
        evaluator = PyEvaluator()
        embeddedEvaluator = EmbeddedPyEvaluator(evaluator)
        context = DictExcelTemplateContext({
            'cates': [{'name':'Computer','items': [
                {'name': 'Dell', 'items': [
                        {'name': 'Dell 1420 XPS','describe': 'fsdfsdfs','count': 5, 'price':1222.4},
                        {'name': 'Dell Inspron 2in1','describe': 'fsdfsdfs','count': 5, 'price':622.4},
                        {'name': 'Dell XPS 200','describe': 'fsdfsdfs','count': 6, 'price':922.4},
                        {'name': 'Dell L3320','describe': 'fsdfsdfs','count': 10, 'price':2222.4},
                ]},
                {'name': 'Lenovo', 'items': [
                        {'name': 'Z510','describe': 'fsdfsdfs','count': 5, 'price':1409},
                        {'name': 'Z520','describe': 'fsdfsdfs','count': 5, 'price':599.9},
                ]},
                {'name': 'HP', 'items': [
                        {'name': 'HP M320','describe': 'fsdfsdfs','count': 5, 'price':1222.4},
                        {'name': 'HP M340','describe': 'fsdfsdfs','count': 5, 'price':1122.4},
                        {'name': 'HP M380','describe': 'fsdfsdfs','count': 5, 'price':1022.4},
                ]},

            ]},
            {'name':'Mobile Phone','items': [
                {'name': 'Apple', 'items': [
                        {'name': 'IPhone 5S', 'describe': 'fsdfsdfs', 'count': 5, 'price':588.4},
                        {'name': 'IPhone 6', 'describe': 'fsdfsdfs', 'count': 5, 'price':599.4},
                        {'name': 'IPhone 6S', 'describe': 'fsdfsdfs', 'count': 5, 'price':622.4},
                        {'name': 'IPhone 7', 'describe': 'fsdfsdfs', 'count': 5, 'price':722.4},
                ]},
                {'name': 'SAMSUNG', 'items': [
                        {'name': 'Note S5', 'describe': 'fsdfsdfs', 'count': 5, 'price':422.4},
                        {'name': 'Note S6', 'describe': 'fsdfsdfs', 'count': 5, 'price':322.4},
                        {'name': 'Note S7', 'describe': 'fsdfsdfs', 'count': 5, 'price':222.4},
                ]},
            ]}]
        })
        archtemplate = TreeTemplate(processor, context, evaluator)
        rptemplate = ExcelRelpaceTemplate(processor, context, embeddedEvaluator)
        archtemplate.process()
        rptemplate.process()
        processor.save('out5.xlsx')
        pass

    def test_div_004(self):
        processor = OpenpyXlExcelProcessor('xlstemplates/template4.xlsx', 'Sheet1')
        evaluator = PyEvaluator()
        embeddedEvaluator = EmbeddedPyEvaluator(evaluator)
        context = DictExcelTemplateContext({
            "students":[
                {'name':'s1', 'age':12, 'class':'2-A' ,'grade':'A-'},
                {'name':'s2', 'age':13, 'class':'2-B' ,'grade':'B-'},
                {'name':'s3', 'age':14, 'class':'2-C' ,'grade':'C-'},
                {'name':'s4', 'age':15, 'class':'2-D' ,'grade':'E-'},
                {'name':'s5', 'age':16, 'class':'2-E' ,'grade':'F-'},
                {'name':'s6', 'age':17, 'class':'2-F' ,'grade':'A+'},
            ]
        })
        archtemplate = TreeTemplate(processor, context, evaluator)
        rptemplate = ExcelRelpaceTemplate(processor, context, embeddedEvaluator)
        archtemplate.process()
        rptemplate.process()
        processor.save('out6.xlsx')
        pass