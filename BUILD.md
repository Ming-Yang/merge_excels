# 打包说明

本文档说明如何将 Excel/CSV文件合并工具打包成可执行文件。

## 前置要求

1. Python 3.7 或更高版本
2. 已安装项目依赖（运行 `pip install -r requirements.txt`）

## 打包方法

### 方法一：使用批处理脚本（推荐，Windows）

直接双击运行 `build.bat` 文件，脚本会自动：
- 检查并安装依赖
- 安装 PyInstaller
- 执行打包

### 方法二：使用 Python 脚本

```bash
python build.py
```

### 方法三：手动使用 PyInstaller

```bash
# 安装 PyInstaller
pip install pyinstaller

# 使用 spec 文件打包
pyinstaller merge_excel_pyside6.spec --clean --noconfirm

# 或直接使用命令打包（目录模式）
pyinstaller --name="Excel合并工具" --windowed --hidden-import=pandas --hidden-import=openpyxl --hidden-import=xlrd --hidden-import=PySide6.QtCore --hidden-import=PySide6.QtGui --hidden-import=PySide6.QtWidgets main.py
```

## 打包输出

打包完成后，可执行文件位于 `dist` 目录中：
- **Windows**: `dist/Excel合并工具/Excel合并工具.exe`
- **Linux/Mac**: `dist/Excel合并工具/Excel合并工具`

打包后会生成一个 `Excel合并工具` 文件夹，其中包含：
- 主可执行文件（`Excel合并工具.exe`）
- 所有依赖的 DLL 和库文件
- Python 运行时文件

## 分发应用

### 目录模式（当前配置）

当前配置为目录模式，生成一个包含所有文件的文件夹。分发时需要将整个文件夹一起分发。

**优点**：
- 启动速度快（无需解压）
- 文件结构清晰，便于调试
- 可以单独更新某些文件

**缺点**：
- 需要分发整个文件夹（通常 100-200MB）
- 文件较多，需要保持文件夹结构

**分发方式**：
1. 将整个 `dist/Excel合并工具` 文件夹压缩成 ZIP 文件
2. 用户解压后，运行其中的 `Excel合并工具.exe` 即可

### 单文件模式（可选）

如果需要单文件模式，可以修改 `merge_excel_pyside6.spec` 文件，将 `exclude_binaries=True` 改为 `False`，并移除 `COLLECT` 部分。

## 常见问题

### 1. 打包后程序无法运行

- 检查是否包含了所有必要的隐藏导入（hiddenimports）
- 尝试在命令行运行可执行文件查看错误信息
- 检查是否有缺失的 DLL 文件

### 2. 文件太大

- 使用 `--exclude-module` 排除不需要的模块
- 考虑使用虚拟环境，只安装必要的依赖

### 3. 杀毒软件误报

PyInstaller 打包的程序有时会被杀毒软件误报。这是正常现象，可以：
- 将程序添加到杀毒软件白名单
- 对程序进行代码签名（需要证书）

### 4. 添加图标

如果需要为可执行文件添加图标：
1. 准备一个 `.ico` 文件（Windows）或 `.icns` 文件（Mac）
2. 在 `merge_excel_pyside6.spec` 文件的 `EXE` 部分，设置 `icon='path/to/icon.ico'`

## 文件说明

- `build.py`: Python 打包脚本，自动检查依赖并执行打包
- `build.bat`: Windows 批处理脚本，一键打包
- `merge_excel_pyside6.spec`: PyInstaller 配置文件，定义打包参数

## 注意事项

1. 打包前建议在虚拟环境中测试
2. 打包后的程序需要在相同或兼容的操作系统上运行
3. 目录模式下启动速度较快，无需等待解压
4. 确保目标机器有足够的磁盘空间（至少 200MB）
5. 分发时需要将整个文件夹一起分发，不能只分发 .exe 文件

