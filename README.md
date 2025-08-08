# PDF 转换器工具集

一个功能完整的 PDF 格式转换工具集，支持 PDF 转 DOCX 和 PDF 转 XLSX 两种格式。

## 功能特点

### PDF 转 DOCX

- 批量转换：一次性转换目录中的所有 PDF 文件到 DOCX 格式
- 单文件转换：转换指定的单个 PDF 文件
- 智能覆盖保护：转换前询问是否覆盖已存在的文件
- 详细进度反馈：显示转换进度和结果统计
- 错误处理：单个文件失败不影响其他文件转换

### PDF 转 XLSX（新功能）

- 表格提取：从 PDF 中智能提取表格数据
- 多种提取策略：支持默认、lattice、stream、文本等多种提取方法
- Excel 格式化：生成格式化的 XLSX 文件，包含样式和自动列宽
- 多工作表支持：多个表格自动分配到不同工作表
- 备选方案：支持 pdfplumber 作为备选提取工具

## 使用方法

### PDF 转 DOCX

#### 批量转换（转换当前目录所有 PDF 文件）

```bash
python pdf_converter.py
```

这会扫描当前目录中的所有 PDF 文件并逐个转换为 DOCX 格式。

#### 单文件转换（推荐用于测试）

```bash
python test_single_file.py <PDF文件名>
```

示例：

```bash
python test_single_file.py mytest.pdf
python test_single_file.py dikongjingji.pdf
```

### PDF 转 XLSX

#### 批量转换（转换当前目录所有 PDF 文件）

```bash
python pdf_to_xlsx_converter.py
```

这会扫描当前目录中的所有 PDF 文件，提取其中的表格数据并转换为 XLSX 格式。

#### 单文件转换（推荐用于测试）

```bash
python test_single_xlsx.py <PDF文件名>
```

示例：

```bash
python test_single_xlsx.py pdf2excel.pdf
python test_single_xlsx.py mytest.pdf
```

## 系统要求

- Python 3.7+

## 安装依赖

### PDF 转 DOCX 依赖

```bash
pip install pdf2docx
```

### PDF 转 XLSX 依赖

```bash
pip install tabula-py pandas openpyxl
```

#### 可选依赖（提高表格提取效果）

```bash
pip install pdfplumber
```

## 文件说明

### PDF 转 DOCX 相关

- `pdf_converter.py` - PDF 转 DOCX 批量转换器
- `test_single_file.py` - PDF 转 DOCX 单文件转换工具

### PDF 转 XLSX 相关

- `pdf_to_xlsx_converter.py` - PDF 转 XLSX 批量转换器
- `test_single_xlsx.py` - PDF 转 XLSX 单文件转换工具

### 其他

- `README.md` - 使用说明

## 转换规则

### PDF 转 DOCX

1. 输出文件与输入文件同名，仅扩展名改为`.docx`
2. 输出文件保存在与输入文件相同的目录中
3. 如果目标 DOCX 文件已存在，会询问是否覆盖
4. 支持各种 PDF 文件名格式（包含空格、特殊字符等）

### PDF 转 XLSX

1. 输出文件与输入文件同名，仅扩展名改为`.xlsx`
2. 输出文件保存在与输入文件相同的目录中
3. 如果目标 XLSX 文件已存在，会询问是否覆盖
4. 多个表格会自动分配到不同的工作表中
5. 自动应用表格格式化（表头样式、边框、列宽等）

## 注意事项

### PDF 转 DOCX

- 扫描版 PDF（纯图片）转换效果有限
- 复杂格式的 PDF 可能需要手动调整转换后的 DOCX 文件
- 转换大文件时请耐心等待

### PDF 转 XLSX

- 主要适用于包含表格数据的 PDF 文件
- 扫描版 PDF 需要先进行 OCR 处理
- 复杂表格布局可能需要手动调整
- 程序会尝试多种提取策略以提高成功率
- 如果没有找到表格，会提示相应的错误信息

1. 检查 PDF 文件是否包含实际的表格数据
2. 确认所有依赖都已正确安装
3. 尝试使用不同的 PDF 文件进行测试
4. 查看详细的错误信息以获取具体指导

