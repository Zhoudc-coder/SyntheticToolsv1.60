import os
import sys
import json
from pathlib import Path

CONFIG_FILENAME = "config.json"
DEFAULT_RETENTION_DAYS = 15
DEFAULT_AUTO_CHECK_PREV_DAYS = 1
DEFAULT_SPECIAL_MODE = 1
DEFAULT_CHECK_DUPLICATES = 0


def _get_default_log_dir() -> Path:
    env_dir = os.environ.get('MERGE_LOG_DIR')
    if env_dir:
        return Path(env_dir)
    if getattr(sys, 'frozen', False):
        base_dir = Path(sys.executable).parent
    else:
        base_dir = Path(__file__).parent
    return base_dir / 'merge_log'


def _find_config_file() -> Path | None:
    if getattr(sys, 'frozen', False):
        exe_dir = Path(sys.executable).parent
    else:
        exe_dir = Path(__file__).parent
    config_path = exe_dir / CONFIG_FILENAME
    if config_path.exists():
        return config_path
    home_config = Path.home() / CONFIG_FILENAME
    if home_config.exists():
        return home_config
    return None


def _load_config() -> dict:
    config_path = _find_config_file()
    if not config_path:
        return {}
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def _normalize_unc_path(path_str: str) -> str:
    path_str = path_str.strip()
    if path_str.startswith('\\\\'):
        return path_str
    if path_str.startswith('\\') and not path_str.startswith('\\\\'):
        return '\\' + path_str
    return path_str


def _load_config_log_dir() -> Path | None:
    config = _load_config()
    log_dir_str = config.get('log_dir')
    if log_dir_str:
        normalized = _normalize_unc_path(log_dir_str)
        return Path(normalized)
    return None


def _load_config_retention_days() -> int:
    config = _load_config()
    retention = config.get('retention_days', DEFAULT_RETENTION_DAYS)
    try:
        days = int(retention)
        if days > 0:
            return days
    except (ValueError, TypeError):
        pass
    return DEFAULT_RETENTION_DAYS


def _load_config_bool_setting(key: str, default: int) -> bool:
    config = _load_config()
    val = config.get(key, default)
    return str(val) == '1'


def _ensure_config_file(default_log_dir: Path):
    config_path = _find_config_file()
    if config_path is None:
        if getattr(sys, 'frozen', False):
            config_dir = Path(sys.executable).parent
        else:
            config_dir = Path(__file__).parent
        config_path = config_dir / CONFIG_FILENAME
        config_data = {
            "log_dir": str(default_log_dir),
            "retention_days": DEFAULT_RETENTION_DAYS,
            "auto_check_prev_days": DEFAULT_AUTO_CHECK_PREV_DAYS,
            "special_mode_default": DEFAULT_SPECIAL_MODE,
            "check_duplicates": DEFAULT_CHECK_DUPLICATES
        }
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=4)
        except Exception:
            home_config = Path.home() / CONFIG_FILENAME
            try:
                with open(home_config, 'w', encoding='utf-8') as f:
                    json.dump(config_data, f, indent=4)
            except Exception:
                pass
    else:
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
            changed = False
            for k, v in [
                ('retention_days', DEFAULT_RETENTION_DAYS),
                ('auto_check_prev_days', DEFAULT_AUTO_CHECK_PREV_DAYS),
                ('special_mode_default', DEFAULT_SPECIAL_MODE),
                ('check_duplicates', DEFAULT_CHECK_DUPLICATES)
            ]:
                if k not in config_data:
                    config_data[k] = v
                    changed = True
            if changed:
                with open(config_path, 'w', encoding='utf-8') as f:
                    json.dump(config_data, f, indent=4)
        except Exception:
            pass


def _get_log_dir() -> Path:
    env_dir = os.environ.get('MERGE_LOG_DIR')
    if env_dir:
        return Path(env_dir)
    config_dir = _load_config_log_dir()
    if config_dir:
        return config_dir
    return _get_default_log_dir()


LOG_DIR = _get_log_dir()
DB_PATH = LOG_DIR / 'merge_log.db'
RETENTION_DAYS = _load_config_retention_days()
AUTO_CHECK_PREV_DAYS = _load_config_bool_setting('auto_check_prev_days', DEFAULT_AUTO_CHECK_PREV_DAYS)
SPECIAL_MODE_DEFAULT = _load_config_bool_setting('special_mode_default', DEFAULT_SPECIAL_MODE)
CHECK_DUPLICATES_DEFAULT = _load_config_bool_setting('check_duplicates', DEFAULT_CHECK_DUPLICATES)
XLSX_SUFFIX = '.xlsx'
CSV_SUFFIX = '.csv'


def ensure_log_dir() -> Path:
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        raise RuntimeError(
            f"无法创建日志目录 {LOG_DIR}，请检查：\n"
            f"1. 路径是否正确（Z 盘是否已连接？）\n"
            f"2. 是否有写入权限\n"
            f"3. 可修改配置文件 config.json 中的 log_dir 项，或设置环境变量 MERGE_LOG_DIR\n"
            f"原始错误：{e}"
        )
    _ensure_config_file(LOG_DIR)
    return LOG_DIR