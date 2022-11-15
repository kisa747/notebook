[![Documentation Status](https://readthedocs.org/projects/kisa747/badge/?version=latest)](https://kisa747.readthedocs.io/zh_CN/latest/?badge=latest)
[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)

# 概述

一个基于 xlwings 库，处理Excel文件的工具。

要求：Python 3.10+

## 安装

```sh
python -m pip install git+https://github.com/kisa747/pyexcel@0.1.0
```

## 快速手册

### 导入模块

```python
import pyexcel
```

### xls 文件转换为 xlsx 文件

```python
# 支持上下文管理，可以自动关闭 Excel 对象
with pyexcel.Excel('test.xls') as ex:
    ex.to_xlsx('test.xlsx')
```

### xlsx 文件导出为pdf 文件

```python
with pyexcel.Excel('test.xlsx') as ex:
    ex.to_pdf('test.pdf')
```

### xlsx 文件去除所有公式

```python
with pyexcel.Excel('test.xlsx') as ex:
    ex.drop_formula('test_without_formula.xlsx')
```

### 批量处理目录内的所有 xlsx 文件

```python
# 批量转换为 xlsx 文件
pyexcel.convert(r'D:\test', 'to_xlsx')

# 批量转换为 pdf 文件
pyexcel.convert(r'D:\test', 'to_xlsx')

# 批量去除所有公式
pyexcel.convert(r'D:\test', 'drop_formula')
```

### 合并 excel 文件

```python
pyexcel.merge(r'D:\test')
```

### 拆分 excel 文件

```python
pyexcel.split('test.xlsx')
```

-------------

[Read the documentation on ReadTheDocs!](https://kisa747.readthedocs.io/zh_CN/latest/)

-------------

# 开发指南

1. 创建虚拟环境。

2. 安装依赖。

```shell
pip insatall requirements.txt
```

3. 测试

```shell
# 静态类型检查
mypy

# 代码测试
pytest

# 代码测试、代码覆盖率测试
tox

# 打包创建 wheel 文件
python -m build -w

# 或是使用 tox 打包创建 wheel 文件
tox -e build
```
