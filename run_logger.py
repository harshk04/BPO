from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from threading import Lock
from typing import Any


@dataclass
class RunLogger:
    file_path: Path
    _lock: Lock = field(default_factory=Lock, repr=False)

    @classmethod
    def create(cls, logs_dir: str = "logs", prefix: str = "run") -> "RunLogger":
        logs_path = Path(logs_dir).expanduser().resolve()
        logs_path.mkdir(parents=True, exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_path = logs_path / f"{prefix}_{timestamp}.txt"

        logger = cls(file_path=file_path)
        logger.log("=== Run Started ===")
        logger.log(f"started_at={datetime.now().isoformat(timespec='seconds')}")
        return logger

    def log(self, message: str) -> None:
        text = str(message)
        with self._lock:
            with self.file_path.open("a", encoding="utf-8") as handle:
                handle.write(text)
                if not text.endswith("\n"):
                    handle.write("\n")

    def section(self, title: str) -> None:
        self.log(f"\n=== {title} ===")

    def log_json(self, title: str, payload: Any) -> None:
        self.section(title)
        try:
            dumped = json.dumps(payload, indent=2, ensure_ascii=False, default=str)
        except TypeError:
            dumped = repr(payload)
        self.log(dumped)

    def close(self) -> None:
        self.log(f"finished_at={datetime.now().isoformat(timespec='seconds')}")
        self.log("=== Run Finished ===")
