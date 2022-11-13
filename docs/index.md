# Python自制包指南

参考：官方文档 [Packaging Python Projects](https://packaging.python.org/en/latest/tutorials/packaging-projects)

[Setuptools 文档](https://setuptools.pypa.io/en/latest/userguide/quickstart.html)

[OpenAstronomy Python Packaging Guide](https://packaging-guide.openastronomy.org/en/latest/index.html#)

## 1、升级打包工具

```sh
python3 -m pip install --upgrade pip
python3 -m pip install --upgrade setuptools wheel build
```

## 2、创建项目

准备创建一个名为 `pyexcel` 的包，目录结构如下：

```sh
D:\MyProject                       # 项目目录
│  LICENSE                         # 授权协议。可以从 https://choosealicense.com 选择合适的授权协议
│  pyproject.toml                  # 包的配置信息
│  README.md                       # 包的简介
│
├─src                              # 源代码根目录
│   └─pyexcel                      # 包目录在这里，必须，否则安装后没有文件夹
│           converts.py
│           excel.py               # 软件包的 python 文件
│           utils.py
│           __init__.py            # 包的初始化文件
│           __main__.py            # 包可以作为脚本调用，python -m pyexcel
│
└─tests                            # 测试源代码根目录
            test_pyexcel.py        # 测试文件
```

在 pycharm 下，将 `src` 目录标记为：源代码根目录；将 `tests` 目录标记为：测试源代码根目录。python程序中的引用就不会报错了。

配置信息 `pyproject.toml` 内容如下：

```toml
# pyproject.toml
[build-system]
requires = ["setuptools>=61.0"]
build-backend = "setuptools.build_meta"

[project]
name = "pyexcel"
version = "1.0.0"
authors = [
  { name="kevin", email="kevin@example.com" },
]
description = "一个处理Excel文件的工具"
readme = "README.md"
requires-python = ">=3.10"
dependencies = [
    "xlwings>=0.28",
]
classifiers = [
    "Programming Language :: Python :: 3 :: Only",
    "License :: OSI Approved :: GNU General Public License v3 (GPLv3)",
    "Operating System :: OS Independent",
    "Development Status :: 5 - Production/Stable",
]

[project.urls]
"Homepage" = "https://example.com"
"Bug Tracker" = "https://github.com/example"
```

### 额外的依赖包

参考：[pip文档](https://pip.pypa.io/en/stable/cli/pip_install/#examples)、[PEP508](https://peps.python.org/pep-0508/#extras)、[SetupTools文档](https://setuptools.pypa.io/en/latest/userguide/dependency_management.html#optional-dependencies)

有时候包的部分功能需要特定的依赖，但又不是所有人会需要这个功能，可以把这个依赖包作为额外依赖包 `extra package`。black[d] 采用的就是这种写法。

```toml
[project]
name = "pyexcel"

# 额外的包
[project.optional-dependencies]
PDF = ["ReportLab>=1.2", "RXP"]
```

如果要安装额外的包，使用下面的命令：

```shell
python -m pip install pyexcel[PDF]
```

如果另一个包 Package-B 依赖 pyexcel 包且支持PDF额外包，需要这样写：

```toml
[project]
name = "Package-B"

# 依赖的包
dependencies = [
    "pyexcel[PDF]"
]
```

classifiers 解释：https://pypi.org/classifiers/

>Programming Language :: Python :: 3
>
>Programming Language :: Python :: 3 :: Only
>
>Programming Language :: Python :: 3.10
>
>License :: OSI Approved :: BSD License
>
>License :: OSI Approved :: GNU General Public License v3 (GPLv3)
>
>Operating System :: OS Independent
>
>Operating System :: Microsoft :: Windows
>
>Operating System :: POSIX :: Linux
>
>Development Status :: 4 - Beta
>
>Development Status :: 5 - Production/Stable

**注1**：

软件包应该放到 `src` 目录下，setuptool 就能自动查找，否则配置起来很麻烦，而且官方文档标明了是测试功能，参考： [官方文档](https://setuptools.pypa.io/en/latest/userguide/package_discovery.html)  。

## 3、执行 build 命令创建包

```sh
# 注意：pip 设置为默认--user的话会报错，需要先执行 pip config set install.user false
# 在项目目录下执行命令
python3 -m build

# 使用 -w 或 --wheel 参数，仅生成wheel文件
python3 -m build -w

# 运行完可以再改回来
pip config set install.user true
```

执行完后，得到一个 dist 目录：

```sh
pyexcel-1.0.0.tar.gz            # 源码包
pyexcel-1.0.0-py3-none-any.whl  # wheel包
```

## 4、安装创建的包

```sh
# 直接本地安装
pip install pyexcel-1.0.0-py3-none-any.whl
```

## 5、上传包

```sh
git add .
git commit -m '发布版 1.0.0'
# 给上一次 commit 打标签
git tag 1.0.0
# 默认情况下，git push并不会把tag标签传送到远端服务器上，只有通过显式命令才能分享标签到远端仓库。
git push origin 1.0.0
# 查看本地 tag 列表
git tag
```

然后在仓库查看是否推送成功，如果成功了，就可以通过下面命令安装。

```sh
# 自动从仓库检出（checkout），然后打包，再安装至系统。
python -m pip install --user git+https://github.com/kisa747/pyexcel@1.0.0
# 或是下面的方法
python -m pip install --user MyProject@git+https://github.com/kisa747/pyexcel@1.0.0
```

上传至 pypi 网站

```sh
python3 -m pip install --upgrade twine
```
