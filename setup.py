
import os

from setuptools import setup, find_packages

this_directory = os.path.abspath(os.path.dirname(__file__))
setup_requirements = []

VERSION = '0.01'


def read_file(filename):
    with open(os.path.join(this_directory, filename), encoding='utf-8') as f:
        long_description = f.read()
    return long_description


setup(
    author="Han Zhichao",
    author_email='superhin@126.com',
    description='XMind TestCase to Excel TestCase',  # 项目描述
    long_description=read_file('README.md'),
    long_description_content_type="text/markdown",
    classifiers=[
        'Programming Language :: Python',
        'Programming Language :: Python :: 3.7',
    ],
    license="MIT license",
    include_package_data=True,
    keywords=[
        'xmind2excel', 'xmind_to_excel',
    ],
    name='xmind2excel2',
    packages=find_packages(include=['xmind2excel2']),
    setup_requires=setup_requirements,
    url='https://github.com/hanzhichao/xmind2excel',  # 项目对应的你的仓库地址
    version=VERSION,
    zip_safe=True,
    install_requires=['openpyxl', 'xmindparser']  # 依赖包
)
