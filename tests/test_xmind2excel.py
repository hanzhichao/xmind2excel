import os

from xmind2excel import xmind2excel


class TestXmind2ExcelAsLibrary:
    """测试作为库使用xmind2excel功能"""

    def test_xmind2excel(self):
        xmind2excel("./testcases.xmind", "./case.xlsx", owner='Kevin')

    def test_xmind2excel_without_excel_file(self):
        xmind2excel("./testcases.xmind", owner='Kevin')

    def test_xmind2excel_without_excel_file_and_owner(self):
        xmind2excel("./testcases.xmind")


class TestXmind2ExcelAsCommandLineTool:
    def test_xmind2excel(self):
        os.system("python3 ../xmind2excel/main.py ./testcases.xmind ./case.xlsx Kevin")

    def test_xmind2excel_without_owner(self):
        os.system("python3 ../xmind2excel/main.py ./testcases.xmind ./case.xlsx")

    def test_xmind2excel_without_excel_file_and_owner(self):
        os.system("python3 ../xmind2excel/main.py ./testcases.xmind")

    def test_xmind2excel_without_xmind_file(self, capsys):
        os.system("python3 ../xmind2excel/main.py")
        captured = capsys.readouterr()
        print('stdout', captured.out)
        print('stderr', captured.err)
        # assert 'ValueError: 缺少XMind文件路径参数' in stdout