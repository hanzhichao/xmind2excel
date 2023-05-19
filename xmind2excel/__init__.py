from pprint import pprint
from configparser import ConfigParser

import openpyxl
from openpyxl.styles import PatternFill
from xmindparser import xmind_to_dict

OWNER = '韩志超'

data = xmind_to_dict('../tests/testcases.xmind')
# pprint(data)

root_topic = data[0]['topic']
testcases = []

for module in root_topic['topics']:
    module_title = module['title']

    for feature in module.get('topics', []):
        feature_title = feature['title']

        for scenario in feature.get('topics', []):
            scenario_title = scenario['title']

            for testcase in scenario.get('topics', []):
                testcase_title = testcase['title']
                priority = ''
                precondition = ''
                owner = OWNER

                steps = []
                excepted = []

                note = testcase.get('note')
                if note:
                    precondition = note  # todo 解析 [预置条件]

                for index, step in enumerate(testcase.get('topics', [])):
                    step_sn = index + 1
                    step_title = step['title']
                    step_excepted = None
                    if 'topics' in step.keys():
                        step_excepted = ' \n'.join([item['title'] for item in step['topics']])
                        excepted.append(f'{step_sn}. {step_excepted}')

                    steps.append(f'{step_sn}. {step_title}')

                markers = testcase.get('markers')
                if markers:
                    priority_markers = [item.startwith('priority') for item in markers]
                    if priority_markers:
                        priority = priority_markers[0]



                test_steps = ' \n'.join(steps)
                test_excepted = ' \n'.join(excepted)
                testcase_data = (module_title, feature_title, scenario_title, testcase_title, precondition, test_steps, test_excepted, owner)
                testcases.append(testcase_data)
                print('测试用例', testcase_data)


excel = openpyxl.Workbook()
sheet = excel.active
title_line = ['模块', '验证功能', '验证场景', '用例标题', '预置条件', '测试步骤', '期望结果', '归属人']
sheet.append(title_line)

color = PatternFill('solid', fgColor="FF00FF")
for i in range(1, 9):
    sheet.cell(1, i).fill = color

for testcase in testcases:
    sheet.append(testcase)

excel.save("case.xlsx")