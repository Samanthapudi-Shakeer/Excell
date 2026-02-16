from __future__ import annotations

import dataclasses
from datetime import datetime
from typing import Optional


@dataclasses.dataclass
class TranslationLogEntry:
    file_name: str
    sheet_name: str
    object_id: str
    original_text: str
    translated_text: str
    engine: str
    status: str
    error: Optional[str] = None
    timestamp: str = dataclasses.field(default_factory=lambda: datetime.utcnow().isoformat() + "Z")


def log_to_dict(entry: TranslationLogEntry) -> dict:
    return dataclasses.asdict(entry)
