import configparser
from pathlib import Path

CONFIG_FILE_NAME = "../configs/config.ini"


def _get_config_path():
    return Path(__file__).resolve().parent / CONFIG_FILE_NAME


def load_config(config_path=None):
    parser = configparser.ConfigParser()
    path = Path(config_path) if config_path is not None else _get_config_path()

    if not path.is_file():
        raise FileNotFoundError(f"Config file not found at: {path}")

    parser.read(path)
    return parser


def get_config_value(
    section,
    option,
    fallback=None,
    config_path=None,
):
    config = load_config(config_path)

    if not config.has_section(section) or not config.has_option(section, option):
        if fallback is not None:
            return fallback
        raise KeyError(f"Missing config value: [{section}] {option}")

    return config.get(section, option)


def get_full_config(config_path=None):
    config = load_config(config_path)
    return {section: dict(config[section]) for section in config.sections()}
