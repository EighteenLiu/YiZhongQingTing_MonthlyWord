"""面向前端展示的轻量日志收集器。"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List


@dataclass
class AppLogger:
    """收集运行过程中的提示、警告和错误。"""

    messages: List[str] = field(default_factory=list)

    def info(self, message: str) -> None:
        self.messages.append(f"[INFO] {message}")

    def warning(self, message: str) -> None:
        self.messages.append(f"[WARNING] {message}")

    def error(self, message: str) -> None:
        self.messages.append(f"[ERROR] {message}")

    def extend(self, messages: List[str]) -> None:
        self.messages.extend(messages)
