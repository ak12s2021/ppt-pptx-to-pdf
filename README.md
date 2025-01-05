# PPT转PDF批量转换工具
利用`pywin32`库，调用PowerPoint。增加了兼容性和后台窗口也能运行，经测试基本在复杂状态（多后台）下依旧可以运行

## 功能描述

这是一个简单的PPT批量转换PDF工具，支持：

- 批量转换PowerPoint文件（.ppt 和 .pptx）
- 保留原始PPT文件夹结构
- 自动跳过已转换的文件
- 支持中文路径
- 详细转换信息输出

## 环境依赖

1. Python 3.7+
2. Windows操作系统
3. Microsoft PowerPoint
4. 依赖库：
   - `pywin32`
   - `pythoncom`

## 安装步骤

1. 安装Python（建议3.7以上版本）
2. 安装依赖库
```bash
pip install pywin32 pypiwin32


