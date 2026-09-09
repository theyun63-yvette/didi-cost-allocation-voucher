from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA_ROOT = ROOT.parent
sys.path.insert(0, str(ROOT))
from services.config import load_config
from services.engine import process_files, save_outputs
from services.utils import for_display

config = load_config()
result = process_files(
    DATA_ROOT / "人员信息表 8月.xlsx",
    DATA_ROOT / "研发工时成本分摊表-2026年1-7月 V5.xlsx",
    DATA_ROOT / "滴滴7月网约车.xlsx",
    DATA_ROOT / "工作簿1.xlsx",
    "2026-07",
    config,
)
out = ROOT / "output" / f"sample-2026-07-v{config['version']}"
paths = save_outputs(result, out)
summary = {
    "target_month": result.target_month,
    "recognized_sheets": result.recognized_sheets,
    "metrics": for_display(result.metrics),
    "blocking_reasons": sorted({f"{x.code}: {x.reason}" for x in result.blocking_issues}),
    "output_files": [str(x) for x in paths],
}
(out / "运行结果.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps(summary, ensure_ascii=False, indent=2))
