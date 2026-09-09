from __future__ import annotations
from decimal import Decimal


def for_display(value):
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, dict):
        return {k: for_display(v) for k, v in value.items()}
    if isinstance(value, list):
        return [for_display(v) for v in value]
    return value
