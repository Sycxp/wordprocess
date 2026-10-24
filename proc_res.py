#!/usr/bin/env python3
import ast
import csv
import html as html_mod
import io
import json
import os
import re
import shlex
import subprocess
import sys
import time
from collections import Counter
from pathlib import Path
from typing import Any

try:
    import openpyxl

    HAS_OPENPYXL = True
except Exception:
    HAS_OPENPYXL = False
from typing import cast

HEADERS = {
    "vocab": ["vocabulary", "meanings", "difficulty", "significance", "note"],
    "phrase": ["phrase", "meanings", "note"],
}

NUMERIC: dict[str, set[str]] = {
    "vocab": {"difficulty", "significance"},
    "phrase": set(),
}

TABLE_ALIASES = {
    "vocab": "vocab",
    "vocabulary": "vocab",
    "word": "vocab",
    "words": "vocab",
    "phrase": "phrase",
    "phrases": "phrase",
}

LEVELS = {
    "小学低年级": 11,
    "小学": 11,
    "低年级": 11,
    "中考": 8,
    "高考": 5,
    "四级": 3,
    "六级": 1,
}

HELP = """命令：
  load <table> [file]           从文件加载；无 file 时粘贴文本，以单独一行 %END% 结束
  tables / info                 查看两张表数量
  show <table> [n]              显示当前视图前 n 条，默认 20
  sort <table> <field> [asc|desc]
  filter <table> <expression>   筛选当前视图
  reset <table>                 恢复原始数据
  output <table> <format> [file] [key=value ...]
  dedupe <table> [first|row]    first 按首列去重，row 按整行去重
  stats [table]                 统计
  exit / quit                   退出

table: vocab 或 phrase
format: csv, tsv, md, json, txt, words
        html, xlsx 已支持；docx/pdf 通过 pandoc 转换
        （中间 Markdown 存于项目目录 Temp/ 下）
  选项: title=...  stats=on/off
        sections=none|auto|exam|significance|difficulty|letter
        columns=default|all|simple|core|study|a,b,c
        layout=compact        紧凑模式（需配合 docx/pdf 使用，单行放多组数据）
        cols=2                每行显示几组数据（默认 2，配合 layout=compact 使用）

表达式示例：
  filter vocab difficulty > 5 and significance >= 8
  filter vocab significance >= level('中考')
  filter vocab contains(meanings, '苹果') or lower(vocabulary) == 'apple'
  filter phrase contains(meanings, '动词') and note != ''

可用函数:
  contains, startswith, endswith, matches,
  lower, upper, strip, len, level/exam
"""

FIELD_LABELS: dict[str, str] = {
    "vocabulary": "单词",
    "phrase": "词组",
    "meanings": "释义",
    "difficulty": "难度",
    "significance": "重要度",
    "note": "备注",
}

DEFAULT_TITLES: dict[str, str] = {
    "vocab": "单词表",
    "phrase": "词组表",
}

EXAM_GROUP_ORDER: list[str] = [
    "小学低年级",
    "中考",
    "高考",
    "四级",
    "六级",
    "未分级",
]

COLUMN_ALIASES: dict[str, str] = {
    "word": "vocabulary",
    "term": "vocabulary",
    "vocab": "vocabulary",
    "phrase": "phrase",
    "meaning": "meanings",
    "meanings": "meanings",
    "释义": "meanings",
    "diff": "difficulty",
    "difficulty": "difficulty",
    "难度": "difficulty",
    "sig": "significance",
    "significance": "significance",
    "importance": "significance",
    "level": "significance",
    "重要度": "significance",
    "note": "note",
    "备注": "note",
    "单词": "vocabulary",
    "词组": "phrase",
}

TEXT_EXPORT_FORMATS = {
    "csv",
    "tsv",
    "json",
    "md",
    "html",
    "txt",
    "words",
}

BINARY_EXPORT_FORMATS = {
    "xlsx",
    "docx",
    "pdf",
}


PROJECT_DIR = Path(__file__).resolve().parent
TEMP_DIR = PROJECT_DIR / "Temp"
PANDOC = os.environ.get("PANDOC", "pandoc")


def ensure_temp_dir() -> Path:
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    return TEMP_DIR


class Table:
    def __init__(self, name):
        self.name = name
        self.header = HEADERS[name]
        self.all = []
        self.view = []

    def load(self, rows):
        self.all = rows
        self.view = rows.copy()

    def reset(self):
        self.view = self.all.copy()


def parse_int(value, default=0):
    s = str(value).strip()
    if not s:
        return default
    try:
        return int(s)
    except ValueError:
        try:
            return int(float(s))
        except ValueError:
            return default


def normalized_fields(fields):
    return [str(x).strip().lstrip("\ufeff").lower() for x in fields]


def get_header_type(fields):
    fs = normalized_fields(fields)
    if fs == [x.lower() for x in HEADERS["vocab"]]:
        return "vocab"
    if fs == [x.lower() for x in HEADERS["phrase"]]:
        return "phrase"
    return None


def normalize_export_format(fmt: str) -> str:
    fmt = str(fmt).strip().lower()

    if fmt in ("markdown", "md"):
        return "md"

    if fmt in ("text", "txt"):
        return "txt"

    if fmt in ("list", "values", "words"):
        return "words"

    if fmt in ("excel", "xlsx"):
        return "xlsx"

    if fmt in ("word", "docx"):
        return "docx"

    if fmt == "pdf":
        return "pdf"

    if fmt == "html":
        return "html"

    if fmt == "csv":
        return "csv"

    if fmt == "tsv":
        return "tsv"

    if fmt == "json":
        return "json"

    return fmt


def parse_export_argv(rest: str):
    """
    解析：
    output <table> <format> [file] [key=value ...]

    例如：
    output vocab md vocab.md title=单词表 stats=on sections=exam
    output vocab xlsx vocab.xlsx columns=vocabulary,meanings,significance
    """
    argv = shlex.split(rest)

    if len(argv) < 2:
        raise ValueError("用法: output <vocab|phrase> <format> [file] [key=value ...]")

    table = resolve_table(argv[0])
    fmt = normalize_export_format(argv[1])

    filename = None
    opts: dict[str, str] = {}

    for token in argv[2:]:
        if "=" in token:
            key, value = token.split("=", 1)
            opts[key.strip().lower()] = value.strip()
        elif filename is None:
            filename = token
        else:
            raise ValueError(f"无法解析参数: {token}")

    return table, fmt, filename, opts


def opt_bool(opts: dict[str, str], keys, default: bool = False) -> bool:
    for key in keys:
        if key in opts:
            value = str(opts[key]).strip().lower()
            return value in ("1", "true", "on", "yes", "y")
    return default


def resolve_columns(table: str, opts: dict[str, str]) -> list[str]:
    """
    支持：
    columns=default
    columns=all
    columns=simple
    columns=core
    columns=study
    columns=vocabulary,meanings,significance
    columns=word,meaning,level
    """
    spec = str(opts.get("columns") or opts.get("cols") or opts.get("fields") or "default").strip()

    low = spec.lower()

    if low in ("", "default", "all", "*", "true"):
        return HEADERS[table][:]

    if low == "simple":
        if table == "vocab":
            return ["vocabulary", "meanings"]
        return ["phrase", "meanings"]

    if low == "core":
        if table == "vocab":
            return ["vocabulary", "meanings", "significance"]
        return ["phrase", "meanings"]

    if low == "study":
        return HEADERS[table][:]

    resolved: list[str] = []

    for raw in spec.split(","):
        raw = raw.strip()
        if not raw:
            continue

        lowname = raw.lower()

        if lowname in ("word", "term"):
            name = HEADERS[table][0]
        elif lowname in COLUMN_ALIASES:
            name = COLUMN_ALIASES[lowname]
        else:
            name = raw.lower()

        if name not in HEADERS[table]:
            raise ValueError(f"{table} 不支持列: {raw}")

        if name not in resolved:
            resolved.append(name)

    if not resolved:
        return HEADERS[table][:]

    return resolved


def get_export_title(table: str, opts: dict[str, str], fmt: str) -> str:
    if "title" in opts:
        return str(opts["title"]).strip()

    if fmt in ("csv", "tsv", "json", "words"):
        return ""

    return DEFAULT_TITLES.get(table, table)


def should_include_stats(fmt: str, opts: dict[str, str]) -> bool:
    default = fmt in ("md", "html", "txt", "docx", "pdf", "xlsx")
    return opt_bool(opts, ("stats", "statistics", "summary"), default)


def get_section_mode(table: str, opts: dict[str, str]):
    """
    sections=none
    sections=auto
    sections=exam
    sections=significance
    sections=difficulty
    sections=letter
    """
    mode = str(opts.get("sections") or opts.get("section") or opts.get("group") or "").strip().lower()

    if mode in ("", "none", "off", "false", "no"):
        return None

    if mode == "auto":
        if table == "vocab":
            return "exam"
        return "letter"

    if mode in ("exam", "level", "levels", "stage", "stages"):
        if table != "vocab":
            raise ValueError("sections=exam 只支持 vocab，phrase 可尝试 sections=letter")
        return "exam"

    if mode in ("sig", "significance"):
        if table != "vocab":
            raise ValueError("sections=significance 只支持 vocab")
        return "significance"

    if mode in ("diff", "difficulty"):
        if table != "vocab":
            raise ValueError("sections=difficulty 只支持 vocab")
        return "difficulty"

    if mode in ("letter", "first", "alpha"):
        return "letter"

    raise ValueError("未知 sections，可用 none/auto/exam/significance/difficulty/letter")


def exam_group_name(significance: Any) -> str:
    s = parse_int(significance, 0)

    if s >= 11:
        return "小学低年级"

    if s >= 8:
        return "中考"

    if s >= 5:
        return "高考"

    if s >= 3:
        return "四级"

    if s >= 1:
        return "六级"

    return "未分级"


def first_letter_group(value: Any) -> str:
    s = str(value).strip()
    if not s:
        return "#"

    ch = s[0].upper()
    if "A" <= ch <= "Z":
        return ch

    return "#"


def group_export_rows(table: str, rows: list[dict[str, Any]], mode: str):
    if mode is None:
        return []

    groups: dict[str, list[dict[str, Any]]] = {}

    if table == "vocab" and mode == "exam":
        for row in rows:
            name = exam_group_name(row.get("significance", 0))
            groups.setdefault(name, []).append(row)

        order = [name for name in EXAM_GROUP_ORDER if name in groups]
        return [(name, groups[name]) for name in order]

    if table == "vocab" and mode == "significance":
        keys = sorted(
            {parse_int(row.get("significance", 0), 0) for row in rows},
            reverse=True,
        )

        for row in rows:
            key = parse_int(row.get("significance", 0), 0)
            name = f"重要度 {key}"
            groups.setdefault(name, []).append(row)

        return [(f"重要度 {key}", groups[f"重要度 {key}"]) for key in keys]

    if table == "vocab" and mode == "difficulty":
        keys = sorted({parse_int(row.get("difficulty", 0), 0) for row in rows})

        for row in rows:
            key = parse_int(row.get("difficulty", 0), 0)
            name = f"难度 {key}"
            groups.setdefault(name, []).append(row)

        return [(f"难度 {key}", groups[f"难度 {key}"]) for key in keys]

    if mode == "letter":
        first_col = HEADERS[table][0]

        for row in rows:
            name = first_letter_group(row.get(first_col, ""))
            groups.setdefault(name, []).append(row)

        keys = sorted(groups.keys(), key=lambda x: (x == "#", x))
        return [(key, groups[key]) for key in keys]

    return []


def vocab_human_stats(rows: list[dict[str, Any]]) -> str:
    n = len(rows)

    if n == 0:
        return "当前视图中没有单词。"

    diffs = [parse_int(row.get("difficulty", 0), 0) for row in rows]
    sigs = [parse_int(row.get("significance", 0), 0) for row in rows]

    avg_d = sum(diffs) / n
    avg_s = sum(sigs) / n

    primary = sum(1 for x in sigs if x >= 11)
    zhongkao = sum(1 for x in sigs if x >= 8)
    gaokao = sum(1 for x in sigs if x >= 5)
    cet4 = sum(1 for x in sigs if x >= 3)
    cet6 = sum(1 for x in sigs if x == 1)

    low = sum(1 for x in diffs if 1 <= x <= 3)
    mid = sum(1 for x in diffs if 4 <= x <= 7)
    high = sum(1 for x in diffs if 8 <= x <= 11)

    diff_counter = Counter(diffs)
    sig_counter = Counter(sigs)

    common_diff = diff_counter.most_common(1)[0]
    common_sig = sig_counter.most_common(1)[0]

    paras = [
        f"当前视图共有 {n} 个单词，平均难度 {avg_d:.1f}，平均重要程度 {avg_s:.1f}。",
        (
            f"按重要程度看，达到小学低年级（11）的有 {primary} 个，"
            f"中考及以上（≥8）的有 {zhongkao} 个，"
            f"高考及以上（≥5）的有 {gaokao} 个，"
            f"四级及以上（≥3）的有 {cet4} 个，"
            f"六级（1）的有 {cet6} 个。"
        ),
        (f"按难度看，低难度（1~3）{low} 个，中难度（4~7）{mid} 个，高难度（8~11）{high} 个。"),
        (
            f"最常见的难度是 {common_diff[0]}，共 {common_diff[1]} 个；"
            f"最常见的重要度是 {common_sig[0]}，共 {common_sig[1]} 个。"
        ),
    ]

    return "\n\n".join(paras)


def phrase_human_stats(rows: list[dict[str, Any]]) -> str:
    n = len(rows)

    if n == 0:
        return "当前视图中没有词组。"

    lengths = [len(str(row.get("phrase", "")).strip()) for row in rows]
    avg_len = sum(lengths) / n

    with_note = sum(1 for row in rows if str(row.get("note", "")).strip())

    letters = Counter(first_letter_group(row.get("phrase", "")) for row in rows)
    common_letter = letters.most_common(1)[0]

    paras = [
        f"当前视图共有 {n} 个词组，平均长度 {avg_len:.1f} 个字符。",
        (f"带备注的词组有 {with_note} 个。开头字母或符号最多的是 {common_letter[0]}，共 {common_letter[1]} 个。"),
    ]

    return "\n\n".join(paras)


def human_stats(table: str, rows: list[dict[str, Any]]) -> str:
    if table == "vocab":
        return vocab_human_stats(rows)
    return phrase_human_stats(rows)


def normalize_row(row, table):
    header = HEADERS[table]
    cells = [str(x).strip() for x in row]

    if len(cells) < len(header):
        cells += [""] * (len(header) - len(cells))
    elif len(cells) > len(header):
        cells = cells[: len(header)]

    d: dict[str, Any] = dict(zip(header, cells))

    for col in NUMERIC[table]:
        d[col] = parse_int(d[col], 0)

    return d


def rows_from_raw_rows(raw_rows, table):
    cleaned = []
    for row in raw_rows:
        if not row:
            continue
        if any(str(cell).strip() for cell in row):
            cleaned.append(row)

    if not cleaned:
        return []

    first_type = get_header_type(cleaned[0])
    data = cleaned[1:] if first_type is not None else cleaned

    rows = []
    for row in data:
        if get_header_type(row) is not None:
            continue
        rows.append(normalize_row(row, table))

    return rows


def rows_from_lines(lines, table):
    text = "\n".join(lines)
    raw_rows = list(csv.reader(io.StringIO(text)))
    return rows_from_raw_rows(raw_rows, table)


def load_file(path, table):
    with open(path, encoding="utf-8-sig", newline="") as f:
        raw_rows = list(csv.reader(f))
    return rows_from_raw_rows(raw_rows, table)


def read_block_from_stdin():
    lines = []
    marker = ""

    while True:
        raw = sys.stdin.readline()
        if raw == "":
            marker = "EOF"
            break

        line = raw.rstrip("\r\n")
        stripped = line.strip()

        if stripped in ("%END%", "~END~"):
            marker = stripped
            break

        lines.append(line)

    return lines, marker


def resolve_table(name):
    key = str(name).strip().lower()
    if key in TABLE_ALIASES:
        return TABLE_ALIASES[key]
    raise ValueError("表名必须是 vocab 或 phrase")


def fn_contains(a, b):
    return str(b).lower() in str(a).lower()


def fn_startswith(a, b):
    return str(a).lower().startswith(str(b).lower())


def fn_endswith(a, b):
    return str(a).lower().endswith(str(b).lower())


def fn_matches(text, pattern):
    try:
        return re.search(str(pattern), str(text), re.IGNORECASE) is not None
    except re.error as e:
        raise ValueError(f"正则表达式错误: {e}")


def fn_level(name):
    s = str(name).strip()
    if s in LEVELS:
        return LEVELS[s]
    for key, value in LEVELS.items():
        if key in s:
            return value
    raise ValueError(f"未知等级/考试: {name}")


FUNCS = {
    "contains": fn_contains,
    "startswith": fn_startswith,
    "endswith": fn_endswith,
    "matches": fn_matches,
    "lower": lambda x: str(x).lower(),
    "upper": lambda x: str(x).upper(),
    "strip": lambda x: str(x).strip(),
    "len": lambda x: len(str(x)),
    "level": fn_level,
    "exam": fn_level,
}


def try_number(value):
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        s = value.strip()
        if not s:
            return None
        try:
            return int(s)
        except ValueError:
            try:
                return float(s)
            except ValueError:
                return None
    return None


def smart_eq(left, right):
    if left == right:
        return True

    ln = try_number(left)
    rn = try_number(right)
    if ln is not None and rn is not None:
        return ln == rn

    return str(left).strip().lower() == str(right).strip().lower()


def smart_order_compare(op: Any, left: Any, right: Any) -> bool:
    ln = try_number(left)
    rn = try_number(right)

    if ln is not None and rn is not None:
        lnum = float(ln)
        rnum = float(rn)

        if isinstance(op, ast.Lt):
            return lnum < rnum
        if isinstance(op, ast.LtE):
            return lnum <= rnum
        if isinstance(op, ast.Gt):
            return lnum > rnum
        if isinstance(op, ast.GtE):
            return lnum >= rnum

        return False

    ltext = str(left).lower()
    rtext = str(right).lower()

    if isinstance(op, ast.Lt):
        return ltext < rtext
    if isinstance(op, ast.LtE):
        return ltext <= rtext
    if isinstance(op, ast.Gt):
        return ltext > rtext
    if isinstance(op, ast.GtE):
        return ltext >= rtext

    return False


def compile_expression(expr):
    expr = expr.strip()
    if not expr:
        raise ValueError("表达式为空")

    try:
        return ast.parse(expr, mode="eval")
    except SyntaxError as e:
        raise ValueError(f"表达式语法错误: {e}")


def eval_expression(compiled, row, table):
    def safe_in(left: Any, right: Any) -> bool:
        try:
            return bool(cast(Any, left) in cast(Any, right))
        except TypeError:
            return str(left) in str(right)

    header = HEADERS[table]
    field_map = {f.lower(): f for f in header}

    def truthy(value):
        return bool(value)

    def eval_node(node):
        if isinstance(node, ast.Expression):
            return eval_node(node.body)

        if isinstance(node, ast.Constant):
            return node.value

        if isinstance(node, ast.Name):
            name = node.id.lower()

            if name in field_map:
                return row[field_map[name]]

            if name == "true":
                return True
            if name == "false":
                return False
            if name == "none":
                return None

            raise ValueError(f"未知字段: {node.id}")

        if isinstance(node, ast.BoolOp):
            if isinstance(node.op, ast.And):
                for value in node.values:
                    if not truthy(eval_node(value)):
                        return False
                return True

            if isinstance(node.op, ast.Or):
                for value in node.values:
                    if truthy(eval_node(value)):
                        return True
                return False

            raise ValueError("不支持的逻辑运算")

        if isinstance(node, ast.UnaryOp):
            if isinstance(node.op, ast.Not):
                return not truthy(eval_node(node.operand))

            if isinstance(node.op, ast.USub):
                value = eval_node(node.operand)
                number = try_number(value)
                return -number if number is not None else 0

            if isinstance(node.op, ast.UAdd):
                return eval_node(node.operand)

            raise ValueError("不支持的一元运算")

        if isinstance(node, ast.BinOp):
            left = eval_node(node.left)
            right = eval_node(node.right)
            ln = try_number(left)
            rn = try_number(right)

            if ln is not None and rn is not None:
                l, r = float(ln), float(rn)
                if isinstance(node.op, ast.Add):
                    return l + r
                if isinstance(node.op, ast.Sub):
                    return l - r
                if isinstance(node.op, ast.Mult):
                    return l * r
                if isinstance(node.op, ast.Div):
                    return l / r
                if isinstance(node.op, ast.FloorDiv):
                    return l // r
                if isinstance(node.op, ast.Mod):
                    return l % r
                if isinstance(node.op, ast.Pow):
                    return l**r
            elif isinstance(node.op, ast.Add):
                return str(left) + str(right)

            raise ValueError("不支持的算术运算")

        if isinstance(node, ast.Compare):
            left = eval_node(node.left)

            for op, comparator in zip(node.ops, node.comparators):
                right = eval_node(comparator)

                if isinstance(op, ast.Eq):
                    ok = smart_eq(left, right)
                elif isinstance(op, ast.NotEq):
                    ok = not smart_eq(left, right)
                elif isinstance(op, (ast.Lt, ast.LtE, ast.Gt, ast.GtE)):
                    ok = smart_order_compare(op, left, right)
                elif isinstance(op, (ast.In, ast.NotIn)):
                    if isinstance(right, str):
                        member = str(left) in right
                    elif isinstance(right, (list, tuple, set)):
                        member = any(smart_eq(left, item) for item in right)
                    else:
                        try:
                            member = safe_in(left, right)
                        except TypeError:
                            member = str(left) in str(right)

                    if isinstance(op, ast.NotIn):
                        member = not member

                    ok = member
                else:
                    raise ValueError("不支持的比较运算")

                if not ok:
                    return False

                left = right

            return True

        if isinstance(node, ast.Call):
            if not isinstance(node.func, ast.Name):
                raise ValueError("只支持普通函数调用")

            if node.keywords:
                raise ValueError("不支持关键字参数")

            fname = node.func.id.lower()
            if fname not in FUNCS:
                raise ValueError(f"不支持函数: {fname}")

            args = [eval_node(x) for x in node.args]

            try:
                return FUNCS[fname](*args)
            except TypeError:
                raise ValueError(f"函数 {fname} 参数错误")

        if isinstance(node, (ast.List, ast.Tuple, ast.Set)):
            return [eval_node(x) for x in node.elts]

        if isinstance(node, ast.Attribute):
            raise ValueError("不支持属性访问")

        raise ValueError("表达式包含不支持的语法")

    return bool(eval_node(compiled))


def render_delimited(table, rows, delimiter=","):
    out = io.StringIO()
    writer = csv.DictWriter(
        out,
        fieldnames=HEADERS[table],
        delimiter=delimiter,
        lineterminator="\n",
        extrasaction="ignore",
    )
    writer.writeheader()

    for row in rows:
        writer.writerow(row)

    return out.getvalue()


def render_md(table, rows):
    header = HEADERS[table]
    lines = [
        "| " + " | ".join(header) + " |",
        "|" + "|".join(["---"] * len(header)) + "|",
    ]

    for row in rows:
        cells = [md_escape(row[col]) for col in header]
        lines.append("| " + " | ".join(cells) + " |")

    return "\n".join(lines)


def render_txt(table, rows):
    lines = []

    for row in rows:
        if table == "vocab":
            parts = [
                row["vocabulary"],
                row["meanings"],
                f"难度:{row['difficulty']}",
                f"重要度:{row['significance']}",
            ]
            if str(row["note"]).strip():
                parts.append(row["note"])
        else:
            parts = [row["phrase"], row["meanings"]]
            if str(row["note"]).strip():
                parts.append(row["note"])

        lines.append(" | ".join(str(x) for x in parts))

    return "\n".join(lines)


def render_words(table, rows):
    first = HEADERS[table][0]
    return "\n".join(str(row[first]) for row in rows)


def render_xlsx(table, rows, columns, title, stats_text, groups) -> bytes:
    if not HAS_OPENPYXL:
        raise ValueError("导出 xlsx 需要安装 openpyxl（pip install openpyxl）")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    if title:
        ws.append([title])
    if stats_text:
        for para in stats_text.split("\n\n"):
            para = para.strip()
            if para:
                ws.append([para])
        ws.append([])

    def write_block(block_rows):
        ws.append([display_field_label(c) for c in columns])
        for r in block_rows:
            ws.append([r.get(c, "") for c in columns])

    if groups:
        for gtitle, grows in groups:
            if not grows:
                continue
            ws.append([gtitle])
            ws.append([f"共 {len(grows)} 条。"])
            write_block(grows)
            ws.append([])
    else:
        write_block(rows)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def render(table, rows, fmt):
    fmt = fmt.lower()

    if fmt == "csv":
        return render_delimited(table, rows, ",")
    if fmt == "tsv":
        return render_delimited(table, rows, "\t")
    if fmt in ("md", "markdown"):
        return render_md(table, rows)
    if fmt == "json":
        return json.dumps(rows, ensure_ascii=False, indent=2)
    if fmt in ("txt", "text"):
        return render_txt(table, rows)
    if fmt in ("words", "list", "values"):
        return render_words(table, rows)
    if fmt == "html":
        return render_html_export(table, rows, HEADERS[table], "", "", None)

    raise ValueError(f"不支持的输出格式: {fmt}")


def write_output(text, filename=None):
    if not text:
        text = "\n"
    elif not text.endswith("\n"):
        text += "\n"

    if filename:
        with open(filename, "w", encoding="utf-8") as f:
            f.write(text)
        print(f"已写入 {filename}")
    else:
        sys.stdout.write(text)


def display_field_label(col: str) -> str:
    return FIELD_LABELS.get(col, col)


def md_escape(value: Any) -> str:
    return str(value).replace("|", "\\|").replace("\n", "<br>")


def md_table(columns: list[str], rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "（无数据）"

    header = [display_field_label(col) for col in columns]

    lines = [
        "| " + " | ".join(header) + " |",
        "|" + "|".join(["---"] * len(columns)) + "|",
    ]

    for row in rows:
        cells = [md_escape(row.get(col, "")) for col in columns]
        lines.append("| " + " | ".join(cells) + " |")

    return "\n".join(lines)


def render_md_export(
    table: str,
    rows: list[dict[str, Any]],
    columns: list[str],
    title: str,
    stats_text: str,
    groups,
) -> str:
    lines: list[str] = []

    if title:
        lines.append(f"# {title}")
        lines.append("")

    if stats_text:
        lines.append(stats_text)
        lines.append("")

    if groups:
        for group_title, group_rows in groups:
            if not group_rows:
                continue

            lines.append(f"## {group_title}")
            lines.append("")
            lines.append(f"共 {len(group_rows)} 条。")
            lines.append("")
            lines.append(md_table(columns, group_rows))
            lines.append("")
    else:
        if rows:
            lines.append(f"共 {len(rows)} 条。")
            lines.append("")

        lines.append(md_table(columns, rows))
        lines.append("")

    return "\n".join(lines).strip() + "\n"


def txt_table(columns: list[str], rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "（无数据）"

    lines = []

    for row in rows:
        parts = []

        for col in columns:
            value = str(row.get(col, "")).strip()
            if value:
                parts.append(f"{display_field_label(col)}: {value}")

        lines.append(" | ".join(parts))

    return "\n".join(lines)


def render_txt_export(
    table: str,
    rows: list[dict[str, Any]],
    columns: list[str],
    title: str,
    stats_text: str,
    groups,
) -> str:
    lines: list[str] = []

    if title:
        lines.append(title)
        lines.append("=" * len(title))
        lines.append("")

    if stats_text:
        lines.extend(stats_text.split("\n\n"))
        lines.append("")

    if groups:
        for group_title, group_rows in groups:
            if not group_rows:
                continue

            lines.append(f"## {group_title}")
            lines.append(f"共 {len(group_rows)} 条。")
            lines.append("")
            lines.append(txt_table(columns, group_rows))
            lines.append("")
    else:
        if rows:
            lines.append(f"共 {len(rows)} 条。")
            lines.append("")

        lines.append(txt_table(columns, rows))

    return "\n".join(lines).strip() + "\n"


def html_table(columns: list[str], rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "<p>（无数据）</p>"

    parts = [
        "<table>",
        "<thead><tr>",
    ]

    for col in columns:
        parts.append(f"<th>{html_mod.escape(display_field_label(col))}</th>")

    parts.append("</tr></thead>")
    parts.append("<tbody>")

    for row in rows:
        parts.append("<tr>")
        for col in columns:
            value = str(row.get(col, ""))
            parts.append(f"<td>{html_mod.escape(value)}</td>")
        parts.append("</tr>")

    parts.append("</tbody>")
    parts.append("</table>")

    return "".join(parts)


def render_html_export(
    table: str,
    rows: list[dict[str, Any]],
    columns: list[str],
    title: str,
    stats_text: str,
    groups,
) -> str:
    parts = [
        "<!DOCTYPE html>",
        '<html lang="zh-CN">',
        "<head>",
        '<meta charset="utf-8">',
    ]

    if title:
        parts.append(f"<title>{html_mod.escape(title)}</title>")

    parts.append(
        """
<style>
body {
    font-family: sans-serif;
    margin: 24px;
}
table {
    border-collapse: collapse;
    width: 100%;
    margin-bottom: 24px;
}
th, td {
    border: 1px solid #999;
    padding: 4px 8px;
    vertical-align: top;
}
th {
    background: #f2f2f2;
}
h1, h2 {
    margin-bottom: 8px;
}
p {
    margin-top: 0;
}
</style>
""".strip()
    )

    parts.append("</head>")
    parts.append("<body>")

    if title:
        parts.append(f"<h1>{html_mod.escape(title)}</h1>")

    if stats_text:
        for para in stats_text.split("\n\n"):
            para = para.strip()
            if para:
                parts.append(f"<p>{html_mod.escape(para)}</p>")

    if groups:
        for group_title, group_rows in groups:
            if not group_rows:
                continue

            parts.append(f"<h2>{html_mod.escape(group_title)}</h2>")
            parts.append(f"<p>共 {len(group_rows)} 条。</p>")
            parts.append(html_table(columns, group_rows))
    else:
        if rows:
            parts.append(f"<p>共 {len(rows)} 条。</p>")

        parts.append(html_table(columns, rows))

    parts.append("</body>")
    parts.append("</html>")

    return "\n".join(parts)


def render_delimited_export(
    rows: list[dict[str, Any]],
    columns: list[str],
    delimiter: str = ",",
) -> str:
    out = io.StringIO()

    writer = csv.DictWriter(
        out,
        fieldnames=columns,
        delimiter=delimiter,
        lineterminator="\n",
        extrasaction="ignore",
    )

    writer.writeheader()

    for row in rows:
        writer.writerow({col: row.get(col, "") for col in columns})

    return out.getvalue()


def render_json_export(
    rows: list[dict[str, Any]],
    columns: list[str],
) -> str:
    data = [{col: row.get(col, "") for col in columns} for row in rows]

    return json.dumps(data, ensure_ascii=False, indent=2)


def render_words_export(
    rows: list[dict[str, Any]],
    columns: list[str],
) -> str:
    if not columns:
        return ""

    first = columns[0]
    return "\n".join(str(row.get(first, "")).strip() for row in rows)


def pandoc_export(table, rows, fmt, filename, columns, title, stats_text, groups, opts=None):
    if opts is None:
        opts = {}
    layout = str(opts.get("layout", "default")).strip().lower()
    cols = parse_int(opts.get("cols", 2), 2)

    if layout == "compact":
        content = render_compact_html(table, rows, columns, title, stats_text, groups, cols)
        input_suffix = ".html"
    else:
        content = render_md_export(table, rows, columns, title, stats_text, groups)
        input_suffix = ".md"

    temp_dir = ensure_temp_dir()
    stamp = int(time.time() * 1000)

    if filename:
        out_path = Path(filename)
        base = out_path.stem or f"export_{stamp}"
    else:
        out_path = None
        base = f"export_{stamp}"

    input_path = temp_dir / f"{base}_{stamp}{input_suffix}"
    input_path.write_text(content, encoding="utf-8")

    tmp_out_path = None
    target_path = out_path
    if target_path is None:
        tmp_out_path = temp_dir / f"{base}.{fmt}"
        target_path = tmp_out_path

    cmd = [PANDOC, str(input_path), "-o", str(target_path)]

    if layout == "compact" and fmt == "pdf":
        print(f"【提示】已生成紧凑 HTML：{input_path}", file=sys.stderr)
        print(
            "为了获得完美的多列排版和紧凑效果，强烈建议直接在浏览器中打开该 HTML 并按 Ctrl+P 打印为 PDF。",
            file=sys.stderr,
        )
        cmd.extend(["--pdf-engine=wkhtmltopdf"])

    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    except FileNotFoundError:
        raise ValueError("未找到 pandoc，请安装 pandoc 或设置环境变量 PANDOC 指向可执行文件")

    if proc.returncode != 0:
        msg = (proc.stderr or proc.stdout or "").strip()
        if layout == "compact" and fmt == "pdf":
            msg += "\n建议：直接打开 Temp/ 目录下的 HTML 文件，在浏览器中打印为 PDF。"
        raise ValueError(f"pandoc 转换失败（{fmt}）: {msg or proc.returncode}")

    if tmp_out_path is not None:
        if not hasattr(sys.stdout, "buffer"):
            raise ValueError("当前 stdout 不支持二进制输出，请提供输出文件名")
        data = tmp_out_path.read_bytes()
        sys.stdout.buffer.write(data)
        sys.stdout.buffer.flush()
        tmp_out_path.unlink(missing_ok=True)
        print(f"（中间文件已保留：{input_path}）", file=sys.stderr)
    else:
        print(f"已写入 {target_path}（中间文件：{input_path}）")


def cmd_load(rest, tables):
    args = rest.split(maxsplit=1)
    if not args:
        raise ValueError("用法: load <vocab|phrase> [file]")

    table = resolve_table(args[0])

    if len(args) == 2:
        path = args[1].strip().strip('"').strip("'")
        rows = load_file(path, table)
    else:
        if sys.stdin.isatty():
            sys.stderr.write(f"请粘贴 {table} 表，然后输入单独一行 %END% 结束：\n")
            sys.stderr.flush()

        lines, _ = read_block_from_stdin()
        rows = rows_from_lines(lines, table)

    tables[table].load(rows)
    print(f"已加载 {table}: {len(tables[table].view)} 条")


def cmd_tables(tables):
    for name in ("vocab", "phrase"):
        t = tables[name]
        print(f"{name}: 当前 {len(t.view)} 条 / 原始 {len(t.all)} 条")


def cmd_show(rest, tables):
    args = rest.split()
    if not args:
        raise ValueError("用法: show <vocab|phrase> [n]")

    table = resolve_table(args[0])
    limit = 20

    if len(args) > 1:
        try:
            limit = max(0, int(args[1]))
        except ValueError:
            raise ValueError("n 必须是整数")

    text = render(table, tables[table].view[:limit], "csv")
    write_output(text)


def cmd_sort(rest, tables):
    args = rest.split()
    if len(args) < 2:
        raise ValueError("用法: sort <vocab|phrase> <field> [asc|desc]")

    table = resolve_table(args[0])
    col = args[1].lower()

    if table == "vocab" and col in ("word", "term"):
        col = "vocabulary"
    if table == "phrase" and col in ("word", "term"):
        col = "phrase"

    if col not in HEADERS[table]:
        raise ValueError(f"{table} 不支持字段: {col}")

    desc = len(args) > 2 and args[2].lower() in ("desc", "d", "descending")
    t = tables[table]

    if col in NUMERIC[table]:

        def numeric_key(row: dict[str, Any]) -> float:
            v = row.get(col, 0)

            if isinstance(v, bool):
                return float(parse_int(v, 0))

            if isinstance(v, int):
                return float(v)

            if isinstance(v, float):
                return v

            return float(parse_int(v, 0))

        t.view.sort(key=numeric_key, reverse=desc)
    else:

        def text_key(row: dict[str, Any]) -> str:
            return str(row.get(col, "")).lower()

        t.view.sort(key=text_key, reverse=desc)

    print(f"已按 {col} {'降序' if desc else '升序'} 排序，当前 {len(t.view)} 条")


def cmd_filter(rest, tables):
    parts = rest.split(maxsplit=1)
    if len(parts) < 2:
        raise ValueError("用法: filter <vocab|phrase> <expression>")

    table = resolve_table(parts[0])
    expr = parts[1].strip()
    compiled = compile_expression(expr)

    t = tables[table]
    old = len(t.view)
    t.view = [row for row in t.view if eval_expression(compiled, row, table)]

    print(f"筛选完成：{old} -> {len(t.view)}")


def cmd_reset(rest, tables):
    if not rest:
        raise ValueError("用法: reset <vocab|phrase>")

    table = resolve_table(rest)
    tables[table].reset()
    print(f"已重置 {table}，当前 {len(tables[table].view)} 条")


def cmd_output(rest, tables):
    table, fmt, filename, opts = parse_export_argv(rest)
    rows = tables[table].view

    columns = resolve_columns(table, opts)
    title = get_export_title(table, opts, fmt)
    stats_text = human_stats(table, rows) if should_include_stats(fmt, opts) else ""
    mode = get_section_mode(table, opts)
    groups = group_export_rows(table, rows, mode)

    if fmt == "md":
        write_output(render_md_export(table, rows, columns, title, stats_text, groups), filename)
        return
    if fmt == "txt":
        write_output(render_txt_export(table, rows, columns, title, stats_text, groups), filename)
        return
    if fmt == "html":
        write_output(render_html_export(table, rows, columns, title, stats_text, groups), filename)
        return
    if fmt == "csv":
        write_output(render_delimited_export(rows, columns, ","), filename)
        return
    if fmt == "tsv":
        write_output(render_delimited_export(rows, columns, "\t"), filename)
        return
    if fmt == "json":
        write_output(render_json_export(rows, columns), filename)
        return
    if fmt == "words":
        write_output(render_words_export(rows, columns), filename)
        return
    if fmt == "xlsx":
        data = render_xlsx(table, rows, columns, title, stats_text, groups)
        if filename:
            with open(filename, "wb") as f:
                f.write(data)
            print(f"已写入 {filename}")
        else:
            sys.stdout.buffer.write(data)
        return
    if fmt in ("docx", "pdf"):
        pandoc_export(table, rows, fmt, filename, columns, title, stats_text, groups, opts)
        return

    raise ValueError(f"不支持的输出格式: {fmt}")


def render_compact_html(table, rows, columns, title, stats_text, groups, cols=2):
    cols = max(cols, 1)

    def build_multicol_table(data_rows):
        parts = ["<table style='border-collapse: collapse; width: 100%; table-layout: fixed; margin-bottom: 4px;'>"]
        parts.append("<thead><tr>")
        for _ in range(cols):
            for col in columns:
                parts.append(
                    f"<th style='border: 1px solid #ccc; padding: 2px 4px; text-align: left; font-weight: bold;'>{display_field_label(col)}</th>"
                )
        parts.append("</tr></thead><tbody>")

        for i in range(0, len(data_rows), cols):
            parts.append("<tr>")
            for c in range(cols):
                idx = i + c
                if idx < len(data_rows):
                    row = data_rows[idx]
                    for col in columns:
                        val = html_mod.escape(str(row.get(col, "")))
                        parts.append(f"<td style='border: 1px solid #ccc; padding: 2px 4px;'>{val}</td>")
                else:
                    for _ in columns:
                        parts.append("<td style='border: none;'></td>")
            parts.append("</tr>")
        parts.append("</tbody></table>")
        return "".join(parts)

    body_parts = []
    if title:
        body_parts.append(f"<h1 style='font-size: 16pt; margin: 4px 0;'>{html_mod.escape(title)}</h1>")
    if stats_text:
        for para in stats_text.split("\n\n"):
            if para.strip():
                body_parts.append(f"<p style='font-size: 10pt; margin: 2px 0;'>{html_mod.escape(para)}</p>")

    if groups:
        for gtitle, grows in groups:
            if not grows:
                continue
            body_parts.append(f"<h2 style='font-size: 14pt; margin: 6px 0 2px 0;'>{html_mod.escape(gtitle)}</h2>")
            body_parts.append(f"<p style='font-size: 10pt; margin: 2px 0;'>共 {len(grows)} 条。</p>")
            body_parts.append(build_multicol_table(grows))
    else:
        if rows:
            body_parts.append(f"<p style='font-size: 10pt; margin: 2px 0;'>共 {len(rows)} 条。</p>")
        body_parts.append(build_multicol_table(rows))

    body = "\n".join(body_parts)

    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<style>
@page {{ margin: 1cm; size: A4; }}
body {{ font-family: "SimSun", "Microsoft YaHei", sans-serif; font-size: 10pt; line-height: 1.2; margin: 0; padding: 0; }}
table {{ page-break-inside: auto; }}
tr {{ page-break-inside: avoid; page-break-after: auto; }}
</style>
</head>
<body>
{body}
</body>
</html>"""


def cmd_dedupe(rest, tables):
    args = rest.split()
    if not args:
        raise ValueError("用法: dedupe <vocab|phrase> [first|row]")

    table = resolve_table(args[0])
    mode = args[1].lower() if len(args) > 1 else "first"

    if mode in ("key", "first"):
        mode = "first"
    elif mode in ("row", "full"):
        mode = "row"
    else:
        raise ValueError("mode 只能是 first 或 row")

    t = tables[table]
    old = len(t.view)
    seen = set()
    out = []
    first_col = HEADERS[table][0]

    for row in t.view:
        if mode == "first":
            key = str(row[first_col]).strip().lower()
        else:
            key = tuple(str(row[col]).strip().lower() for col in HEADERS[table])

        if key not in seen:
            seen.add(key)
            out.append(row)

    t.view = out
    print(f"去重完成：{old} -> {len(t.view)}")


def stats_text(table, t):
    rows = t.view
    n = len(rows)

    if table == "vocab":
        if n == 0:
            return "vocab: 0 条"

        avg_d = sum(row["difficulty"] for row in rows) / n
        avg_s = sum(row["significance"] for row in rows) / n

        diff = Counter(row["difficulty"] for row in rows)
        sig = Counter(row["significance"] for row in rows)

        diff_text = ", ".join(f"{k}:{v}" for k, v in sorted(diff.items())) or "无"
        sig_text = ", ".join(f"{k}:{v}" for k, v in sorted(sig.items())) or "无"

        return "\n".join(
            [
                f"vocab: {n} 条",
                f"平均难度: {avg_d:.2f}",
                f"平均重要度: {avg_s:.2f}",
                f"难度分布: {diff_text}",
                f"重要度分布: {sig_text}",
            ]
        )

    return f"phrase: {n} 条"


def cmd_stats(rest, tables):
    args = rest.split()

    if args:
        names = [resolve_table(args[0])]
    else:
        names = ["vocab", "phrase"]

    blocks = [stats_text(name, tables[name]) for name in names]
    sys.stdout.write("\n\n".join(blocks) + "\n")


def handle_command(line, tables):
    line = line.strip()
    if not line:
        return True

    if line.lower() in ("exit", "quit"):
        return False

    parts = line.split(maxsplit=1)
    cmd = parts[0].lower()
    rest = parts[1].strip() if len(parts) > 1 else ""

    try:
        if cmd == "help":
            print(HELP)
        elif cmd == "load":
            cmd_load(rest, tables)
        elif cmd in ("tables", "info"):
            cmd_tables(tables)
        elif cmd == "show":
            cmd_show(rest, tables)
        elif cmd == "sort":
            cmd_sort(rest, tables)
        elif cmd == "filter":
            cmd_filter(rest, tables)
        elif cmd == "reset":
            cmd_reset(rest, tables)
        elif cmd == "output":
            cmd_output(rest, tables)
        elif cmd == "dedupe":
            cmd_dedupe(rest, tables)
        elif cmd == "stats":
            cmd_stats(rest, tables)
        else:
            print(f"未知命令: {cmd}，输入 help 查看帮助")
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)

    return True


def command_loop(tables):
    prompt = "vocab-cli> " if sys.stdin.isatty() else ""

    while True:
        if prompt:
            sys.stderr.write(prompt)
            sys.stderr.flush()

        raw = sys.stdin.readline()
        if raw == "":
            break

        if not handle_command(raw, tables):
            break


def main():
    args = sys.argv[1:]
    tables = {name: Table(name) for name in HEADERS}

    if len(args) == 2:
        try:
            tables["vocab"].load(load_file(args[0], "vocab"))
            tables["phrase"].load(load_file(args[1], "phrase"))
        except Exception as e:
            print(f"读取文件失败: {e}", file=sys.stderr)
            sys.exit(1)
    elif len(args) != 0:
        print("用法: python vocab_cli.py [vocab.csv phrase.csv]", file=sys.stderr)
        sys.exit(1)
    else:
        if sys.stdin.isatty():
            sys.stderr.write("未提供文件。可用 load 命令加载，输入 help 查看帮助。\n")

    try:
        command_loop(tables)
    except KeyboardInterrupt:
        print()


if __name__ == "__main__":
    main()
