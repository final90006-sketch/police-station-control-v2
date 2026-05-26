"""
framework.py — 高強度測試框架（6 大類）

每頁設計完成後必跑：
  1. 公式語法驗證
  2. 邊界資料測試
  3. 性能基準
  4. 跨平台相容
  5. 真實使用情境模擬
  6. 錯誤恢復
"""
from dataclasses import dataclass, field
from typing import Callable, List
from enum import Enum


class Category(str, Enum):
    SYNTAX = "1. 公式語法"
    BOUNDARY = "2. 邊界資料"
    PERFORMANCE = "3. 性能基準"
    CROSS_PLATFORM = "4. 跨平台相容"
    REAL_WORLD = "5. 真實情境"
    ERROR_RECOVERY = "6. 錯誤恢復"


@dataclass
class TestResult:
    name: str
    category: Category
    passed: bool
    detail: str = ""
    measured: str = ""

    def __str__(self):
        icon = "✓" if self.passed else "✗"
        line = f"  [{icon}] {self.category.value} | {self.name}"
        if self.detail:
            line += f" :: {self.detail}"
        if self.measured:
            line += f" (measured={self.measured})"
        return line


@dataclass
class TestSuite:
    """單頁測試套件 — 依 6 大類組織"""
    page_name: str
    results: List[TestResult] = field(default_factory=list)

    def add(self, name: str, category: Category, passed: bool,
            detail: str = "", measured: str = ""):
        self.results.append(TestResult(name, category, passed, detail, measured))

    def assert_eq(self, name: str, category: Category, actual, expected,
                  detail: str = ""):
        passed = actual == expected
        self.add(name, category, passed,
                 detail=detail or f"expected={expected!r}",
                 measured=str(actual)[:60])
        return passed

    def assert_true(self, name: str, category: Category, condition: bool,
                    detail: str = ""):
        self.add(name, category, bool(condition), detail=detail)
        return bool(condition)

    def assert_in_range(self, name: str, category: Category, value, lo, hi,
                        detail: str = ""):
        passed = lo <= value <= hi
        self.add(name, category, passed,
                 detail=detail or f"range=[{lo}, {hi}]",
                 measured=str(value))
        return passed

    def assert_le(self, name: str, category: Category, value, threshold,
                  detail: str = ""):
        passed = value <= threshold
        self.add(name, category, passed,
                 detail=detail or f"<= {threshold}",
                 measured=str(value))
        return passed

    def run_check(self, name: str, category: Category, fn: Callable[[], bool],
                  detail: str = ""):
        try:
            passed = bool(fn())
            self.add(name, category, passed, detail=detail)
        except Exception as e:
            self.add(name, category, False, detail=f"exception: {e}")

    @property
    def passed_count(self):
        return sum(1 for r in self.results if r.passed)

    @property
    def failed_count(self):
        return sum(1 for r in self.results if not r.passed)

    @property
    def total(self):
        return len(self.results)

    def by_category(self):
        out = {c: [] for c in Category}
        for r in self.results:
            out[r.category].append(r)
        return out

    def report(self):
        lines = [f"\n{'='*60}",
                 f" 測試報告：{self.page_name}",
                 f"{'='*60}"]
        for cat, items in self.by_category().items():
            if not items:
                continue
            cat_pass = sum(1 for r in items if r.passed)
            lines.append(f"\n[{cat.value}] {cat_pass}/{len(items)}")
            for r in items:
                lines.append(str(r))
        lines.append(f"\n{'-'*60}")
        lines.append(f" 總計：{self.passed_count}/{self.total} 通過 "
                     f"({self.failed_count} 失敗)")
        lines.append(f"{'='*60}\n")
        return "\n".join(lines)

    def is_all_pass(self):
        return self.failed_count == 0 and self.total > 0
