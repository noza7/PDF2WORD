# PDF to Word Converter

## English

### Introduction
This is a PDF to Word converter tool with a user-friendly GUI built using PyQt6. It leverages Microsoft Word's capabilities to convert PDF files to Word documents with high fidelity, while also providing additional formatting improvements.

### Features
- Convert PDF files to Word documents (.docx)
- Automatically fix superscript formatting issues
- Intelligently handle numbered paragraphs
- Maintain consistent formatting throughout the document
- Simple and intuitive user interface
- Progress tracking during conversion

### Requirements
- Windows operating system
- Microsoft Word installed
- Python 3.6+
- PyQt6
- pywin32

### Installation
1. Ensure you have Python installed on your system
2. Install the required packages:
```
pip install PyQt6 pywin32
```
3. Run the application:
```
python pdf_to_word_pyqt.py
```

### Usage
1. Launch the application
2. Click "Browse..." to select a PDF file
3. Choose an output directory (defaults to the same directory as the PDF)
4. Click "Start Conversion"
5. Wait for the conversion to complete
6. The converted Word document will be saved in the specified output directory

### Advanced Features
- **Superscript Fixing**: The tool automatically detects and fixes superscript formatting issues, ensuring that superscript text maintains the same formatting as surrounding text.
- **Numbered Paragraph Handling**: Automatically identifies numbered paragraphs and ensures they are properly formatted with consistent indentation.
- **Dialog Handling**: Automatically handles any dialog boxes that may appear during the conversion process.

---

## 中文

### 简介
这是一个具有用户友好界面的PDF转Word转换工具，使用PyQt6构建。它利用Microsoft Word的功能将PDF文件转换为Word文档，同时提供额外的格式改进。

### 功能特点
- 将PDF文件转换为Word文档（.docx）
- 自动修复上标格式问题
- 智能处理编号段落
- 保持文档格式一致性
- 简单直观的用户界面
- 转换过程中的进度跟踪

### 系统要求
- Windows操作系统
- 安装了Microsoft Word
- Python 3.6+
- PyQt6
- pywin32

### 安装方法
1. 确保您的系统上已安装Python
2. 安装所需的包：
```
pip install PyQt6 pywin32
```
3. 运行应用程序：
```
python pdf_to_word_pyqt.py
```

### 使用方法
1. 启动应用程序
2. 点击"浏览..."选择PDF文件
3. 选择输出目录（默认为PDF文件所在的目录）
4. 点击"开始转换"
5. 等待转换完成
6. 转换后的Word文档将保存在指定的输出目录中

### 高级功能
- **上标修复**：工具自动检测并修复上标格式问题，确保上标文本与周围文本保持相同的格式。
- **编号段落处理**：自动识别编号段落，并确保它们具有适当的格式和一致的缩进。
- **对话框处理**：自动处理转换过程中可能出现的任何对话框。
