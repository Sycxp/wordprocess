import csv
import io
import sys
from collections import Counter

VOCAB_HEADER = ["vocabulary", "meanings", "difficulty", "significance", "note"]
PHRASE_HEADER = ["phrase", "meanings", "note"]

# True：按首列 vocabulary / phrase 去重
# False：按整行完全相同去重
DEDUPE_BY_FIRST_COLUMN = True


def normalized_fields(fields):
    return [str(x).strip().lstrip("\ufeff").lower() for x in fields]


def get_table_type(fields):
    """
    判断一行是不是表头，返回：
    'vocabulary' / 'phrase' / None
    """
    fs = normalized_fields(fields)
    if fs == VOCAB_HEADER:
        return "vocabulary"
    if fs == PHRASE_HEADER:
        return "phrase"
    return None


def parse_lines_to_rows(lines):
    """
    把一张表的多行文本解析成 CSV 行列表。
    过滤完全空行。
    """
    text = "\n".join(lines)
    reader = csv.reader(io.StringIO(text))
    return [row for row in reader if row and any(cell.strip() for cell in row)]


def normalize_row(row, width):
    """
    将数据行规范成固定列数：
    - 少列补空字符串
    - 多列截断
    """
    row = [cell.strip() for cell in row]
    if len(row) < width:
        row += [""] * (width - len(row))
    elif len(row) > width:
        row = row[:width]
    return tuple(row)


def dedupe(rows):
    """
    去重，保留第一次出现的项。
    """
    seen = set()
    result = []

    for row in rows:
        if DEDUPE_BY_FIRST_COLUMN:
            key = row[0]
        else:
            key = row

        if key not in seen:
            seen.add(key)
            result.append(row)

    return result


def main():
    blocks = []
    buffer = []

    # 读取输入，直到单独一行的 ~END~
    for raw in sys.stdin:
        line = raw.rstrip("\r\n")
        marker = line.strip()

        if marker == "~END~":
            # 如果最后一张表忘了用 %END% 结束，也尝试处理
            if buffer:
                blocks.append(buffer)
            break

        if marker == "%END%":
            if buffer:
                blocks.append(buffer)
            buffer = []
            continue

        buffer.append(line)
    else:
        # 没有 ~END~ 直接 EOF 时，也处理已有输入
        if buffer:
            blocks.append(buffer)

    vocab_rows = []
    phrase_rows = []

    for lines in blocks:
        rows = parse_lines_to_rows(lines)
        if not rows:
            continue

        table_type = get_table_type(rows[0])

        if table_type is not None:
            data_rows = rows[1:]
        else:
            # 容错：如果某张表没有表头，则按列数推断
            width_counter = Counter(len(row) for row in rows)
            most_common_width = width_counter.most_common(1)[0][0]

            if most_common_width == len(VOCAB_HEADER):
                table_type = "vocabulary"
                data_rows = rows
            elif most_common_width == len(PHRASE_HEADER):
                table_type = "phrase"
                data_rows = rows
            else:
                # 无法识别的表，跳过
                continue

        if table_type == "vocabulary":
            header = VOCAB_HEADER
            target = vocab_rows
        else:
            header = PHRASE_HEADER
            target = phrase_rows

        width = len(header)

        for row in data_rows:
            # 跳过可能混入的表头行
            if get_table_type(row) is not None:
                continue
            target.append(normalize_row(row, width))

    vocab_rows = dedupe(vocab_rows)
    phrase_rows = dedupe(phrase_rows)

    writer = csv.writer(sys.stdout, lineterminator="\n")

    # 输出第一张大表
    writer.writerow(VOCAB_HEADER)
    writer.writerows(vocab_rows)

    # 两张表之间空一行分隔
    sys.stdout.write("\n")

    # 输出第二张大表
    writer.writerow(PHRASE_HEADER)
    writer.writerows(phrase_rows)


if __name__ == "__main__":
    main()
