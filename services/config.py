from __future__ import annotations

from copy import deepcopy
from pathlib import Path
import yaml


def _deep_merge(base: dict, override: dict) -> dict:
    result = deepcopy(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = deepcopy(value)
    return result


def load_config(path: str | Path | None = None) -> dict:
    config_path = Path(path) if path else Path(__file__).resolve().parents[1] / "config" / "rules.yaml"
    with config_path.open("r", encoding="utf-8") as f:
        config = yaml.safe_load(f) or {}

    # 本机专用配置不进入GitHub。若存在，则覆盖公开基础配置。
    local_path = config_path.with_name("rules.local.yaml")
    if local_path.exists():
        with local_path.open("r", encoding="utf-8") as f:
            local_config = yaml.safe_load(f) or {}
        config = _deep_merge(config, local_config)
    return config
