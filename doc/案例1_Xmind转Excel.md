# 案例1：XMind转Excel测试用例

## 需求场景

> 你同事之前都是用XMind来写用例及梳理测试点的，现在组里要求测试完成，测试用例需要都整理到TAPD（腾讯的一个用例及缺陷管理平台）上，TAPD上可以通过Excel批量导入，重写一遍Excel版的用例又非常费力，现在需要你实现一个XMind转Excel测试用例的工具，并制定XMind用例编写规范。

## &#x20;提示

*   需要的能力

    *   Python基础
    *   列表、字典的遍历、数据提取及组装
    *   会pip安装，并参考官方文档使用三方包
*   练习重点

    *   列表、字典数据的提取及组装
    *   Excel文件的读写
*   难点

    *   XMind格式规范的设计
    *   数据缺失及异常处理
*   参考三方包：

    *   xmind（只兼容XMind8以下版本文件）: <https://pypi.org/project/XMind/>
    *   xmindparser: <https://github.com/tobyqin/xmindparser>
    *   openpyxl（可以读/写Excel，只兼容.xlsx，不兼容.xls文件）
*   发布

    *   注册一个GitHub或Gitee，使git将源码推送到自己的仓库中去
    *   注册一个Pypi，将项目发布到Pypi上
    *   项目结构参考

        *   `xmind2excel`  # 项目文件夹

            *   `xmind2excel`  # 包

                *   `__init__.py`  # 简单逻辑可以直接在这里实现
                *   ...
            *   `requirements.txt`  # 三方依赖包列表
            *   `README.md`  # 使用说明
            *   ... 开源协议等
            *   `setup.py`  # 安装脚本
            *   `tests` # 单元测试目录

                *   `__init__.py` &#x20;
                *   ...  # 单元测试脚本
*   优化建议

    *   提供命令行接口
    *   提供一些可自定义的参数或配置，提高灵活性
    *   Excel标题行加一些漂亮的样式
    *   使用tk提供操作界面
    *   使用pyinstaller生成可执行文件



