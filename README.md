# Words - 英语词汇处理工具集

这是一个基于 Python 的命令行工具集，用于高效处理、清洗、分析和导出英语单词（vocabulary）与词组（phrase）数据。适用于从原始文本中提取词汇、去重、多维度筛选（如难度、重要度），并导出为 Markdown、Word、PDF 等多种格式。

## ✨ 功能特性

- **数据清洗 (`proc_csv.py`)**：
  - 从标准输入流读取文本块。
  - 自动识别表头（支持 `vocabulary` 和 `phrase`）。
  - 支持按首列（去重单词/词组）或整行去重。
  - 自动规范化列数并输出标准 CSV。
- **交互式分析 (`proc_res.py`)**：
  - 内置 `vocab-cli` 交互式命令行。
  - 强大的表达式过滤引擎（支持 `>`, `<`, `contains`, `level()` 等）。
  - 灵活的排序和分组（按考试等级、难度或首字母）。
  - 支持导出为 `CSV, TSV, JSON, Markdown, TXT, HTML, XLSX, DOCX, PDF` 等丰富格式。

## 🛠️ 环境依赖

- Python 3.9 或更高版本
- （可选）导出 Excel 所需：`pip install openpyxl`
- （可选）导出 Word/PDF 所需：需要安装 [Pandoc](https://pandoc.org/) 并配置在系统环境变量中。

## 🚀 快速开始

### 1. 数据清洗 (`proc_csv.py`)

`proc_csv.py` 用于处理粘贴的原始文本或重定向的文本文件。数据块之间使用 `%END%` 或 `~END~` 作为结束分隔符。

**输入示例**：

```text
vocabulary,meanings,difficulty,significance,note
apple,苹果,1,11,
banana,香蕉,1,11,常见水果
%END%
phrase,meanings,note
look up,查阅,动词短语
%END%
```

**运行方式**：

```bash
# 直接从标准输入运行（手动粘贴，输入 ~END~ 结束）
python proc_csv.py

# 或者从文件读取并输出到文件
python proc_csv.py < raw_data.txt > clean_data.csv
```

### 2. 数据分析与导出 (`proc_res.py`)

你可以直接传入清洗好的 `vocab.csv` 和 `phrase.csv` 启动交互式终端：

```bash
python proc_res.py vocab.csv phrase.csv
```

启动后，你将进入 `vocab-cli>` 提示符，可以执行以下常用命令：

- **查看数据**：`show vocab 10` （查看前10个单词）
- **条件筛选**：`filter vocab difficulty > 5 and significance >= 8` （筛选难度大于5且重要度大于等于8的单词）
- **使用内置函数**：`filter vocab contains(meanings, '水果')`
- **排序**：`sort vocab difficulty desc` （按难度降序排列）
- **生成统计**：`stats vocab`
- **去重**：`dedupe vocab first`
- **一键导出**：

  ```bash
  # 导出为带统计信息和按考试分组的 Markdown
  output vocab md vocab_export.md title=核心词汇表 stats=on sections=exam

  # 导出为紧凑排版的 Word 文档
  output vocab docx vocab.docx layout=compact cols=2
  ```

## 📚 进阶命令说明

### `output` 命令支持的参数

- **`format`**: `csv`, `tsv`, `md`, `json`, `txt`, `words`, `html`, `xlsx`, `docx`, `pdf`
- **`title=...`**: 设置文档标题。
- **`stats=on/off`**: 是否在导出文件头部追加数据统计摘要。
- **`sections=...`**: 分组方式，支持 `none`, `auto`, `exam`（按中高考/四六级分组）, `significance`, `difficulty`, `letter`（按首字母）。
- **`columns=...`**: 导出指定的列，如 `columns=vocabulary,meanings,significance`。
- **`layout=compact`**: 专为 Word/PDF 设计的紧凑多列排版。

### 过滤表达式支持的函数

- `contains(a, b)` / `startswith(a, b)` / `endswith(a, b)`
- `matches(text, 'regex')` (正则匹配)
- `lower()` / `upper()` / `strip()` / `len()`
- `level('中考')` / `exam('高考')` (直接筛选指定考试等级)

## 📁 项目目录结构

```text
Words/
├── vocab.csv             # 词汇数据表（由 proc_csv.py 生成）
├── phrase.csv            # 词组数据表（由 proc_csv.py 生成）
├── proc_csv.py           # 文本清洗与格式化脚本
├── proc_res.py           # 数据查询、分析与导出 CLI 脚本
├── README.md             # 本文档
├── LICENSE               # 开源协议
└── Temp/                 # 存放 pandoc 转换 docx/pdf 时的中间 Markdown/HTML 文件（已加入 .gitignore）
```

## 📄 许可证

本项目采用 [GPL-3.0-or-later](LICENSE) 协议开源。
