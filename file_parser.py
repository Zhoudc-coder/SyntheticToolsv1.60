import re
import datetime
from pathlib import Path
from dataclasses import dataclass

@dataclass
class ParsedBoxInfo:
    box_code: str
    package_date: str      # YYYY-MM-DD
    sequence_no: int       # 当日序号
    version: str           # 版本号（例如 '11'，不带 v）
    date_from_mtime: bool = False   # 是否使用文件修改时间作为日期


def parse_filename(file_path: Path) -> ParsedBoxInfo | None:
    """
    解析文件名，提取箱码、打包日期、当日序号和版本号。
    如果文件名中找不到日期，则回退使用文件的修改时间（mtime）。
    """
    stem = file_path.stem

    # 1. 提取日期（优先匹配 -YYMMDD- 格式）
    date_match = re.search(r'-(\d{6})-', stem)
    if not date_match:
        date_match = re.search(r'(\d{6})', stem)

    date_from_mtime = False
    package_date = None

    if date_match:
        date_str = date_match.group(1)
        yy, mm, dd = date_str[:2], date_str[2:4], date_str[4:6]
        try:
            year = 2000 + int(yy)
            month, day = int(mm), int(dd)
            package_date = datetime.date(year, month, day).isoformat()
        except ValueError:
            package_date = None

    # 2. 如果日期解析失败，使用文件修改时间
    if package_date is None:
        try:
            mtime = file_path.stat().st_mtime
            dt = datetime.datetime.fromtimestamp(mtime)
            package_date = dt.date().isoformat()
            date_from_mtime = True
        except Exception:
            return None

    # 3. 箱码使用完整文件名（不含扩展名）
    box_code = stem

    # 4. 提取序号（优先从日期后的部分提取，否则从整个文件名提取）
    if date_match:
        suffix = stem[date_match.end():]
    else:
        suffix = stem
    number_sequences = re.findall(r'\d+', suffix)
    if not number_sequences:
        return None

    sequence_no = None
    for num_str in reversed(number_sequences):
        if len(num_str) <= 4:
            try:
                sequence_no = int(num_str)
                break
            except ValueError:
                continue
    if sequence_no is None:
        try:
            sequence_no = int(number_sequences[0])
        except ValueError:
            return None

    # 5. 提取版本号：查找 `(v数字)` 模式
    version_match = re.search(r'\(v(\d+)\)', stem)
    if not version_match:
        return None
    version = version_match.group(1)

    return ParsedBoxInfo(
        box_code=box_code,
        package_date=package_date,
        sequence_no=sequence_no,
        version=version,
        date_from_mtime=date_from_mtime
    )


def validate_sn(sn: str) -> bool:
    """
    校验SN格式：
    前半部分18位码，以ABC开头，以XYZ结尾
    后半部分为八个以+分隔的字段，每个字段可为任意非空字符（字母、数字等）。
    """
    if not sn:
        return False
    sn = sn.strip()

    parts = sn.split('+')
    if len(parts) != 9:          # 1个前缀 + 8个段
        return False

    prefix = parts[0]
    if len(prefix) != 18:
        return False
    if not prefix.startswith('CYV') or not prefix.endswith('H3K'):
        return False

    # 后半部分8个段，每个段必须非空
    for part in parts[1:]:
        if not part:
            return False

    return True