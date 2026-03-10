# AGENTS.md

This file provides guidance to Qoder (qoder.com) when working with code in this repository.

## 项目概述

这是一个文档转PDF工具，支持图片、Office文档、PDF文件的转换和合并。采用Flask Web框架实现，可通过PyInstaller打包为独立可执行文件。

### 核心功能
- 图片转PDF（JPG, PNG, BMP, GIF, TIFF, WEBP等）
- Office文档转PDF（DOC, DOCX, XLS, XLSX, PPT, PPTX）
- PDF文件合并
- 拖拽排序文件顺序
- Web界面操作

## 开发环境设置

### 安装依赖
```bash
pip install -r requirements.txt
```

### 运行开发服务器
```bash
python app.py
```

### 打包可执行文件
```bash
# Linux
./build.sh

# Windows  
build.bat
```

## 代码架构

### 主要组件

1. **app.py** - Flask Web应用主入口
   - 提供Web界面和API接口
   - 处理文件上传、转换请求
   - 管理文件生命周期

2. **converter.py** - 核心转换逻辑
   - 文件类型检测和分类
   - 图片/PDF/Office文档转换
   - LibreOffice集成支持
   - 后备转换方案（纯Python实现）

3. **前端组件**
   - `templates/index.html` - 主页面模板
   - `static/app.js` - 前端交互逻辑
   - `static/style.css` - 样式文件

### 转换流程

```
用户上传文件 → 文件类型识别 → 分类处理 → 
图片: img2pdf → PDF字节 → 图片提取
Office: LibreOffice/python-docx → PDF → 图片提取  
PDF: 直接图片提取 → 统一合并为PDF输出
```

### 技术栈
- **后端**: Python 3.11+, Flask
- **图像处理**: Pillow, img2pdf, PyMuPDF
- **Office处理**: python-docx, openpyxl, python-pptx, LibreOffice
- **PDF处理**: PyMuPDF, ReportLab
- **前端**: 原生HTML/CSS/JavaScript
- **打包**: PyInstaller

## 常用开发命令

### 测试运行
```bash
python app.py
```

### 查看系统状态
```bash
curl http://localhost:5000/api/status
```

### 手动测试API
```bash
# 上传文件
curl -X POST -F "files=@test.jpg" http://localhost:5000/api/upload

# 转换文件
curl -X POST -H "Content-Type: application/json" \
  -d '{"file_ids": ["file_uuid"]}' \
  http://localhost:5000/api/convert
```

### 打包构建
```bash
# Linux构建
./build.sh

# Windows构建  
build.bat
```

### 常见构建问题及解决方案

#### 1. pathlib包冲突错误
**错误信息**: `The 'pathlib' package is an obsolete backport of a standard library package and is incompatible with PyInstaller`

**解决方案**:
```bash
pip uninstall pathlib -y
```

这是由于第三方pathlib包与Python标准库冲突导致的，卸载第三方包即可解决。

### GitHub Actions自动化构建
推送到带有`v*`标签的分支会触发自动构建和发布。

## 关键设计决策

### 文件处理策略
- 所有文件最终都转换为图片格式再合并为PDF
- 支持LibreOffice作为首选Office转换引擎
- 提供纯Python后备方案确保基本功能

### 性能优化
- 使用临时目录管理中间文件
- 异步处理避免阻塞主线程
- 合理的内存管理和垃圾回收

### 兼容性考虑
- 多平台支持（Windows/Linux）
- 多架构支持（x86_64/ARM64）
- glibc版本兼容性处理

## 注意事项

### 依赖管理
- 保持requirements.txt中依赖版本稳定
- 排除不必要的大型依赖（如AI框架）
- 注意PyInstaller打包时的隐藏导入

### 安全考虑
- 限制上传文件大小（500MB）
- 验证文件类型和扩展名
- 清理临时文件防止磁盘空间耗尽

### 用户体验
- 提供清晰的状态反馈
- 支持拖拽排序调整PDF页面顺序
- 自动检测LibreOffice并给出相应提示