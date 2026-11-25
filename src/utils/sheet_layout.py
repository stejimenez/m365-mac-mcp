"""
Layout-Agnostic 2D Semantic Multi-Table Analyzer & Smart Magnitude Engine (IMP-003 + IMP-001).

Pure-Python, read-only structural analyzer for openpyxl worksheets.
NEVER assumes table headers live on Row 1 or row labels live in Column A.
Handles:
  - Arbitrary (row, col) table origins (e.g., Row 6 under 4-line headers, indented Col B, ERP C8)
  - Stacked and side-by-side multi-table worksheets (even of different heights)
  - Institutional title blocks (Rows 1-4) and section banners
  - 2-way sensitivity matrices on any sheet name
  - Mid-table section sub-headers (e.g., CURRENT ASSETS / OPERATING ACTIVITIES)
  - Account-code + label columns (ERP / IFRS trial balances)
  - Formula-AST + multilingual (EN/FR/ES) subtotal and grand-total detection
  - Smart Magnitude Option-A currency scaling ($ / $1,000's / $ Millions / $ Billions)
    with automatic per-share / per-unit exact-currency protection.
"""

from __future__ import annotations

import re
from collections import deque
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Sequence, Set, Tuple

from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Currency & Smart Magnitude Helpers (IMP-001 — Option A, Whole Numbers)
# ---------------------------------------------------------------------------

CURRENCY_SYMBOLS: Dict[str, str] = {
    "USD": "$",
    "EUR": "€",
    "GBP": "£",
    "CAD": "$",
    "AUD": "$",
    "MXN": "$",
    "JPY": "¥",
    "CNY": "¥",
    "CHF": "CHF ",
    "INR": "₹",
}

SYMBOL_TO_CURRENCY: Dict[str, str] = {
    "$": "USD",
    "€": "EUR",
    "£": "GBP",
    "¥": "JPY",
    "₹": "INR",
}

PER_UNIT_KEYWORDS: Tuple[str, ...] = (
    "share price", "price per share", "per share", "eps", "dps", "nav per",
    "unit price", "price per unit", "cost per unit", "per unit", "arpu",
    "asp", "selling price", "average salary", "salary", "hourly rate", "wage",
)

PCT_KEYWORDS: Tuple[str, ...] = (
    "%", "margin", "rate", "growth", "wacc", "irr", "mirr", "cagr",
    "yield", "scale", "discount rate", "tax rate", "cost of equity",
    "cost of debt", "erp", "risk premium", "risk-free", "tasa",
    "crecimiento", "margen", "taux", "croissance", "marge",
)

_PCT_WORD_RE = re.compile(
    r"\b(?:" + "|".join(re.escape(k) for k in PCT_KEYWORDS if k != "%") + r")\b",
    re.IGNORECASE,
)


def is_pct_label(label: str) -> bool:
    """Return True if a row label denotes a percentage/rate (word-bounded so 'erp' never matches 'Enterprise')."""
    if not label:
        return False
    core = re.sub(r"\([^)]*\)", "", str(label)).strip() or str(label)
    if "%" in core:
        return True
    return bool(_PCT_WORD_RE.search(core))

MULT_KEYWORDS: Tuple[str, ...] = (
    "multiple", "profitability index", "index", "ev/", "p/e", "moic", "tvpi", "dpi",
)

FACTOR_KEYWORDS: Tuple[str, ...] = (
    "disc. factor", "discount factor", "pv factor", "factor de descuento",
)

PAYBACK_KEYWORDS: Tuple[str, ...] = (
    "payback", "recuperación", "délai de récupération",
)

INT_COUNT_KEYWORDS: Tuple[str, ...] = (
    "life", "horizon", "project life", "shares", "headcount", "employees", "units", "count",
)

NON_CURRENCY_SCALE_KEYWORDS: Tuple[str, ...] = (
    "payback", "life", "horizon", "multiple", "index", "year", "factor",
    "shares", "headcount", "employees", "units", "count", "beta",
)

GRAND_TOTAL_KEYWORDS: Tuple[str, ...] = (
    "total assets", "total liabilities", "total equity", "total liabilities & equity",
    "total liabilities and equity", "net income", "net profit", "ufcf",
    "unlevered free cash flow", "free cash flow", "levered free cash flow",
    "cum. ufcf", "cum. pv", "cumulative ufcf", "cumulative pv",
    "npv", "net present value", "enterprise value", "equity value",
    "implied share price", "share price", "balance check", "ending cash",
    "net cash", "total actif", "total passif", "résultat net",
    "utilidad neta", "flujo de caja libre", "total activo", "total pasivo",
)

SUBTOTAL_KEYWORDS: Tuple[str, ...] = (
    "total", "subtotal", "gross profit", "ebitda", "ebit", "operating income",
    "nopat", "ebt", "pre-tax income", "operating cash flow", "investing cash flow",
    "financing cash flow", "bénéfice brut", "utilidad bruta", "total ",
)

KEY_VALUE_HEADER_WORDS: Set[str] = {
    "value", "values", "input", "inputs", "amount", "amounts", "metric", "metrics",
    "kpi", "kpis", "parameter", "parameters", "assumption", "assumptions",
    "driver", "drivers", "unit", "units", "notes", "note", "description", "result",
    "results", "base", "bull", "bear", "total", "year", "period", "line item",
}


@dataclass
class MagnitudeSpec:
    """Resolved currency and display magnitude for Option-A Excel formatting."""

    currency_code: str = "USD"
    currency_symbol: str = "$"
    scale_name: str = "exact"       # 'exact' | 'thousands' | 'millions' | 'billions'
    scale_divisor: int = 1          # 1 | 1000 | 1_000_000 | 1_000_000_000
    display_scaled: bool = False    # True when Option-A trailing commas scale exact base numbers
    header_label: str = "In USD $"
    currency_format: str = '$#,##0;($#,##0);"-"'
    per_unit_format: str = '$#,##0.00;($#,##0.00);"-"'


def normalize_currency(raw: Optional[str]) -> Tuple[str, str]:
    """Return (ISO_CODE, symbol) from any user currency string ('USD', '$', 'EUR €', 'GBP')."""
    if not raw or not str(raw).strip():
        return "USD", "$"
    s = str(raw).strip()
    up = s.upper()
    for code, sym in CURRENCY_SYMBOLS.items():
        if re.search(rf"\b{code}\b", up):
            return code, sym
    for sym, code in SYMBOL_TO_CURRENCY.items():
        if sym in s:
            return code, sym
    return "USD", "$"


def build_magnitude_spec(
    currency: Optional[str] = "USD",
    scale_name: str = "exact",
    display_scaled: bool = True,
    decimals_if_stored_millions: bool = False,
    has_per_unit_rows: bool = False,
) -> MagnitudeSpec:
    """Construct exact Option-A (trailing-comma scaled) or stored-unit MagnitudeSpec."""
    code, sym = normalize_currency(currency)
    sname = (scale_name or "exact").strip().lower()
    if sname in ("k", "thousand", "thousands", "000", "000s", "1000", "1,000", "1,000's"):
        sname, divisor, commas, unit_phrase = "thousands", 1_000, ",", f"{code} {sym}1,000's"
    elif sname in ("m", "mm", "mn", "million", "millions", "1000000"):
        sname, divisor, commas, unit_phrase = "millions", 1_000_000, ",,", f"{code} {sym} Millions"
    elif sname in ("b", "bn", "billion", "billions", "1000000000"):
        sname, divisor, commas, unit_phrase = "billions", 1_000_000_000, ",,,", f"{code} {sym} Billions"
    else:
        sname, divisor, commas, unit_phrase = "exact", 1, "", f"{code} {sym}"

    # Format strings: Option-A uses Excel trailing commas (whole numbers per developer decision)
    if display_scaled and divisor > 1:
        cur_fmt = f'{sym}#,##0{commas};({sym}#,##0{commas});"-"'
    elif decimals_if_stored_millions:
        cur_fmt = f'{sym}#,##0.00;({sym}#,##0.00);"-"'
    else:
        cur_fmt = f'{sym}#,##0;({sym}#,##0);"-"'

    per_unit_fmt = f'{sym}#,##0.00;({sym}#,##0.00);"-"'
    hdr = f"In {unit_phrase.strip()}"
    if has_per_unit_rows and divisor > 1:
        hdr += " (except per-share / per-unit data)"

    return MagnitudeSpec(
        currency_code=code,
        currency_symbol=sym,
        scale_name=sname,
        scale_divisor=divisor,
        display_scaled=bool(display_scaled and divisor > 1),
        header_label=hdr,
        currency_format=cur_fmt,
        per_unit_format=per_unit_fmt,
    )


def resolve_smart_magnitude(
    max_macro_amount: float,
    currency: Optional[str] = "USD",
    requested_magnitude: Optional[str] = "auto",
    has_stored_millions_cue: bool = False,
    has_per_unit_rows: bool = False,
) -> MagnitudeSpec:
    """Resolve Smart Magnitude tier from workbook monetary scale (Option A — whole numbers)."""
    req = (requested_magnitude or "auto").strip().lower()
    if req not in ("", "auto", "smart"):
        return build_magnitude_spec(
            currency=currency,
            scale_name=req,
            display_scaled=(not has_stored_millions_cue),
            decimals_if_stored_millions=has_stored_millions_cue,
            has_per_unit_rows=has_per_unit_rows,
        )

    # Legacy stored-$M workbooks (e.g. CapEx=25 with '($M)' label): keep 2-decimal unscaled format
    if has_stored_millions_cue and 0 < max_macro_amount < 10_000:
        return build_magnitude_spec(
            currency=currency,
            scale_name="millions",
            display_scaled=False,
            decimals_if_stored_millions=True,
            has_per_unit_rows=has_per_unit_rows,
        )

    # Smart Magnitude Auto-Tiers (exact base storage -> Option-A trailing-comma display)
    if max_macro_amount >= 1_000_000_000_000:
        tier = "billions"
        use_scaled = True
    elif max_macro_amount >= 1_000_000_000:
        tier = "millions"
        use_scaled = True
    elif max_macro_amount >= 1_000_000:
        tier = "thousands"
        use_scaled = True
    elif 0 < max_macro_amount < 1_000:
        # Small stored numbers (<1000, e.g. $M model without explicit header) -> 2 decimals unscaled
        return build_magnitude_spec(
            currency=currency,
            scale_name="exact",
            display_scaled=False,
            decimals_if_stored_millions=True,
            has_per_unit_rows=has_per_unit_rows,
        )
    else:
        tier = "exact"
        use_scaled = False

    return build_magnitude_spec(
        currency=currency,
        scale_name=tier,
        display_scaled=use_scaled,
        decimals_if_stored_millions=False,
        has_per_unit_rows=has_per_unit_rows,
    )


def is_per_unit_label(label: str) -> bool:
    low = (label or "").strip().lower()
    return bool(low) and any(kw in low for kw in PER_UNIT_KEYWORDS)


def parse_magnitude_from_text(text: str) -> Optional[Tuple[str, str]]:
    """Extract (currency_code, scale_name) from an institutional header line like 'In USD $1,000's' or 'In EUR € Millions'."""
    if not text or not isinstance(text, str):
        return None
    s = text.strip()
    low = s.lower()
    has_currency_cue = any(
        tok in low for tok in ("usd", "eur", "gbp", "cad", "aud", "mxn", "jpy", "chf", "$", "€", "£", "¥")
    )
    if not has_currency_cue:
        return None
    code, _ = normalize_currency(s)
    if any(kw in low for kw in ("billion", "$b", "€b", "£b", "bn")):
        return code, "billions"
    if any(kw in low for kw in ("million", "$m", "€m", "£m", "($m)", "mm")):
        return code, "millions"
    if any(kw in low for kw in ("thousand", "1,000", "000's", "000s", "($k)", "$k")):
        return code, "thousands"
    if re.search(r"\bin\s+(?:usd|eur|gbp|cad|aud|mxn|jpy|chf|\$|€|£)", low):
        return code, "exact"
    return None


# ---------------------------------------------------------------------------
# 2D Semantic Multi-Table Layout Data Structures (IMP-003)
# ---------------------------------------------------------------------------

@dataclass
class TableBlock:
    """A semantically classified 2D rectangular region on a worksheet."""

    kind: str                       # 'TITLE_BLOCK' | 'SECTION_BANNER' | 'DATA_TABLE' | 'MATRIX_TABLE'
    min_row: int
    max_row: int
    min_col: int
    max_col: int
    label_col: int                  # Column holding row labels (never hardcoded to Col A)
    code_col: Optional[int] = None  # Optional ERP account-code column left of label_col
    header_rows: List[int] = field(default_factory=list)
    subheader_rows: Set[int] = field(default_factory=set)
    subtotal_rows: Set[int] = field(default_factory=set)
    total_rows: Set[int] = field(default_factory=set)
    is_assumptions_block: bool = False
    matrix_header_row: Optional[int] = None
    matrix_axis_col: Optional[int] = None

    @property
    def range_ref(self) -> str:
        return f"{get_column_letter(self.min_col)}{self.min_row}:{get_column_letter(self.max_col)}{self.max_row}"

    @property
    def value_start_col(self) -> int:
        return self.label_col + 1

    @property
    def primary_header_row(self) -> Optional[int]:
        if self.matrix_header_row is not None:
            return self.matrix_header_row
        return self.header_rows[-1] if self.header_rows else None

    def contains(self, row: int, col: int) -> bool:
        return self.min_row <= row <= self.max_row and self.min_col <= col <= self.max_col

    def to_summary_dict(self) -> Dict[str, Any]:
        out: Dict[str, Any] = {
            "kind": self.kind,
            "range": self.range_ref,
            "label_col": get_column_letter(self.label_col),
        }
        if self.header_rows:
            out["header_rows"] = list(self.header_rows)
        elif self.matrix_header_row is not None:
            out["header_rows"] = [self.matrix_header_row]
        if self.code_col is not None:
            out["code_col"] = get_column_letter(self.code_col)
        return out


@dataclass
class SheetLayout:
    """Complete 2D layout map of a single openpyxl worksheet."""

    sheet_name: str
    blocks: List[TableBlock] = field(default_factory=list)
    title_block: Optional[TableBlock] = None
    overflow_banner_cells: Set[Tuple[int, int]] = field(default_factory=set)
    recommended_freeze_cell: Optional[str] = None
    declared_currency: Optional[str] = None
    declared_magnitude: Optional[str] = None

    def block_at(self, row: int, col: int) -> Optional[TableBlock]:
        for blk in self.blocks:
            if blk.contains(row, col):
                return blk
        return None

    def data_blocks(self) -> List[TableBlock]:
        return [b for b in self.blocks if b.kind in ("DATA_TABLE", "MATRIX_TABLE")]

    def row_label_for(self, ws, row: int, col: int) -> str:
        blk = self.block_at(row, col)
        lcol = blk.label_col if blk is not None else 1
        val = ws.cell(row=row, column=lcol).value
        return str(val).strip() if val is not None else ""

    def col_header_for(self, ws, row: int, col: int) -> Optional[Any]:
        blk = self.block_at(row, col)
        if blk is not None and blk.primary_header_row is not None:
            return ws.cell(row=blk.primary_header_row, column=col).value
        return ws.cell(row=1, column=col).value


# ---------------------------------------------------------------------------
# Internal Structural Helpers
# ---------------------------------------------------------------------------

def _is_formula(v: Any) -> bool:
    return isinstance(v, str) and v.strip().startswith("=")


def _is_number(v: Any) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _is_text(v: Any) -> bool:
    return isinstance(v, str) and bool(v.strip()) and not v.strip().startswith("=")


def _is_monotonic_period_sequence(nums: Sequence[float]) -> bool:
    """True for year/period axes like [0,1,2,3,4,5,6], [1,2,3,4,5], [2020,2021,2022,2023]."""
    if len(nums) < 2:
        return False
    if not all(float(x).is_integer() for x in nums):
        return False
    ints = [int(x) for x in nums]
    diffs = [ints[i + 1] - ints[i] for i in range(len(ints) - 1)]
    if all(d == 1 for d in diffs):
        if (0 <= ints[0] <= 20) or (1900 <= ints[0] <= 2150):
            return True
    return False


def _is_monotonic_numeric_axis(nums: Sequence[float]) -> bool:
    """True for sensitivity axis steps like [0.085,0.095,0.105,0.115,0.125] or [0.8,0.9,1.0,1.1,1.2]."""
    if len(nums) < 2:
        return False
    inc = all(nums[i + 1] > nums[i] for i in range(len(nums) - 1))
    dec = all(nums[i + 1] < nums[i] for i in range(len(nums) - 1))
    return inc or dec


def _formula_is_vertical_rollup(formula: str, row: int, col: int) -> bool:
    """Detect vertical SUM/SUBTOTAL or >=3-term same-column vertical addition (language-agnostic total signal)."""
    if not _is_formula(formula):
        return False
    s = formula.strip()
    col_l = get_column_letter(col)
    # 1. =SUM(C2:C10) or =SUBTOTAL(9,C2:C10) in same column above current row
    m = re.search(
        rf"(?:SUM|SUBTOTAL)\s*\((?:\d+\s*,\s*)?\$?{col_l}\$?(\d+)\s*:\s*\$?{col_l}\$?(\d+)\)",
        s,
        re.IGNORECASE,
    )
    if m:
        r1, r2 = int(m.group(1)), int(m.group(2))
        if min(r1, r2) < row and max(r1, r2) < row and abs(r2 - r1) >= 1:
            return True
    # 2. Additive chain of >=3 cells in same column above current row (e.g., =C8+C9+C10+C11+C12)
    if "/" not in s and "*" not in s:
        refs = re.findall(rf"(?<![A-Za-z0-9_!])\$?{col_l}\$?(\d+)(?![A-Za-z0-9_])", s, re.IGNORECASE)
        above_rows = {int(r_s) for r_s in refs if int(r_s) < row}
        if len(above_rows) >= 3:
            return True
    return False


def _find_connected_components(ws) -> List[Tuple[int, int, int, int]]:
    """Segment occupied worksheet cells (including merged spans) into 2D connected bounding rectangles.

    Uses openpyxl's sparse `ws._cells` dictionary so worksheets with ghost formatting at
    row 1,048,576 run in O(populated_cells) without allocating millions of empty Cell objects.
    """
    occupied: Set[Tuple[int, int]] = set()
    cells_dict = getattr(ws, "_cells", None)
    if isinstance(cells_dict, dict):
        for (r, c), cell_obj in list(cells_dict.items()):
            if cell_obj is not None and cell_obj.value is not None:
                occupied.add((r, c))
    else:
        max_r = ws.max_row or 0
        max_c = ws.max_column or 0
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                if ws.cell(row=r, column=c).value is not None:
                    occupied.add((r, c))

    # Include cells inside merged ranges that have a non-empty anchor so merged headers stay connected
    for mrange in getattr(ws.merged_cells, "ranges", []):
        if (mrange.min_row, mrange.min_col) in occupied:
            for r in range(mrange.min_row, mrange.max_row + 1):
                for c in range(mrange.min_col, mrange.max_col + 1):
                    occupied.add((r, c))

    if not occupied:
        return []

    visited: Set[Tuple[int, int]] = set()
    rects: List[Tuple[int, int, int, int]] = []

    for cell in sorted(occupied):
        if cell in visited:
            continue
        q = deque([cell])
        visited.add(cell)
        min_r = max_r_b = cell[0]
        min_c = max_c_b = cell[1]
        while q:
            cr, cc = q.popleft()
            if cr < min_r:
                min_r = cr
            if cr > max_r_b:
                max_r_b = cr
            if cc < min_c:
                min_c = cc
            if cc > max_c_b:
                max_c_b = cc
            for nr, nc in ((cr - 1, cc), (cr + 1, cc), (cr, cc - 1), (cr, cc + 1)):
                if (nr, nc) in occupied and (nr, nc) not in visited:
                    visited.add((nr, nc))
                    q.append((nr, nc))
        rects.append((min_r, max_r_b, min_c, max_c_b))

    # Merge any overlapping bounding boxes (rare L-shaped layouts) cleanly
    merged = True
    while merged:
        merged = False
        out: List[Tuple[int, int, int, int]] = []
        used = [False] * len(rects)
        for i in range(len(rects)):
            if used[i]:
                continue
            r1, r2, c1, c2 = rects[i]
            used[i] = True
            for j in range(i + 1, len(rects)):
                if used[j]:
                    continue
                jr1, jr2, jc1, jc2 = rects[j]
                if not (r2 < jr1 or jr2 < r1 or c2 < jc1 or jc2 < c1):
                    r1, r2 = min(r1, jr1), max(r2, jr2)
                    c1, c2 = min(c1, jc1), max(c2, jc2)
                    used[j] = True
                    merged = True
            out.append((r1, r2, c1, c2))
        rects = out

    rects.sort(key=lambda b: (b[0], b[2]))
    return rects


def _peel_top_banners(
    ws, rect: Tuple[int, int, int, int]
) -> List[Tuple[int, int, int, int]]:
    """If a multi-column block has 1..4 top single-column text rows (no blank spacer), peel them off."""
    r1, r2, c1, c2 = rect
    if c2 <= c1 or r2 <= r1:
        return [rect]

    peeled_end = r1 - 1
    for r in range(r1, min(r1 + 4, r2)):
        row_non_empty_cols = [c for c in range(c1, c2 + 1) if ws.cell(row=r, column=c).value is not None]
        if len(row_non_empty_cols) == 1 and row_non_empty_cols[0] == c1 and _is_text(ws.cell(row=r, column=c1).value):
            peeled_end = r
        else:
            break

    # Only peel if the remaining block below actually has >= 2 populated columns in its first row
    if peeled_end >= r1:
        next_r = peeled_end + 1
        next_cols = [c for c in range(c1, c2 + 1) if ws.cell(row=next_r, column=c).value is not None]
        if len(next_cols) >= 2:
            # Don't peel a 1-row banner on Row 1 if it's a compact 2-column sheet title (like ["Assumptions", None] on A1:B13)
            # unless there are >=2 title rows (Rows 1..4 institutional block) OR a section banner below Row 1.
            if (peeled_end - r1 + 1) >= 2 or r1 > 1:
                return [(r1, peeled_end, c1, c1), (next_r, r2, c1, c2)]

    return [rect]


def _classify_block(ws, rect: Tuple[int, int, int, int], sheet_title: str) -> TableBlock:
    """Semantically classify a single 2D rectangular block on a worksheet."""
    r1, r2, c1, c2 = rect
    height = r2 - r1 + 1
    width = c2 - c1 + 1
    title_low = (sheet_title or "").lower()

    # 1. Single-column text blocks: TITLE_BLOCK (top of sheet) or SECTION_BANNER
    if width == 1:
        all_text = all(
            _is_text(ws.cell(row=r, column=c1).value)
            for r in range(r1, r2 + 1)
            if ws.cell(row=r, column=c1).value is not None
        )
        if all_text:
            kind = "TITLE_BLOCK" if (r1 <= 5 and height <= 5) else "SECTION_BANNER"
            return TableBlock(
                kind=kind,
                min_row=r1,
                max_row=r2,
                min_col=c1,
                max_col=c2,
                label_col=c1,
            )

    # 2. Detect 2-way MATRIX_TABLE (Sensitivity table on ANY sheet name, ANY row/col)
    if width >= 3 and height >= 3:
        top_nums = [
            ws.cell(row=r1, column=c).value
            for c in range(c1 + 1, c2 + 1)
            if ws.cell(row=r1, column=c).value is not None
        ]
        left_nums = [
            ws.cell(row=r, column=c1).value
            for r in range(r1 + 1, r2 + 1)
            if ws.cell(row=r, column=c1).value is not None
        ]
        if (
            len(top_nums) >= 2
            and len(left_nums) >= 2
            and all(_is_number(v) for v in top_nums)
            and all(_is_number(v) for v in left_nums)
            and _is_monotonic_numeric_axis([float(v) for v in top_nums])
            and _is_monotonic_numeric_axis([float(v) for v in left_nums])
        ):
            body_cells = [
                ws.cell(row=r, column=c).value
                for r in range(r1 + 1, r2 + 1)
                for c in range(c1 + 1, c2 + 1)
                if ws.cell(row=r, column=c).value is not None
            ]
            fml_ratio = (
                sum(1 for v in body_cells if _is_formula(v)) / len(body_cells)
                if body_cells
                else 0.0
            )
            if fml_ratio >= 0.7:
                return TableBlock(
                    kind="MATRIX_TABLE",
                    min_row=r1,
                    max_row=r2,
                    min_col=c1,
                    max_col=c2,
                    label_col=c1,
                    header_rows=[r1],
                    matrix_header_row=r1,
                    matrix_axis_col=c1,
                )

    # 3. DATA_TABLE — Detect optional ERP account-code column + label column
    code_col: Optional[int] = None
    label_col = c1
    if width >= 3 and height >= 3:
        c1_body = [ws.cell(row=r, column=c1).value for r in range(r1 + 1, r2 + 1) if ws.cell(row=r, column=c1).value is not None]
        c2_body = [ws.cell(row=r, column=c1 + 1).value for r in range(r1 + 1, r2 + 1) if ws.cell(row=r, column=c1 + 1).value is not None]
        if (
            len(c1_body) >= 2
            and len(c2_body) >= 2
            and all(_is_number(v) and float(v).is_integer() and 100 <= abs(float(v)) <= 999999 for v in c1_body)
            and all(_is_text(v) for v in c2_body)
        ):
            code_col = c1
            label_col = c1 + 1

    # 4. Detect header_rows (0, 1, or 2 stacked header rows at top of block)
    header_rows: List[int] = []
    if height >= 2:
        for r in range(r1, min(r1 + 2, r2)):
            vals = [ws.cell(row=r, column=c).value for c in range(c1, c2 + 1)]
            non_empty = [v for v in vals if v is not None]
            if not non_empty:
                break
            if any(_is_formula(v) for v in non_empty):
                break

            val_cols = [ws.cell(row=r, column=c).value for c in range(label_col + 1, c2 + 1)]
            val_non_empty = [v for v in val_cols if v is not None]
            num_val_cols = [float(v) for v in val_non_empty if _is_number(v)]

            # Case H0: 1-cell banner title directly on top of a 2-col table (e.g., ["Assumptions", None] on A1:B13)
            if r == r1 and len(non_empty) == 1 and _is_text(vals[0]):
                header_rows.append(r)
                continue

            # Case H1: All-text header row (skip headerless key-value rows like ["Accounting Basis", "US GAAP"])
            if len(non_empty) >= 2 and all(_is_text(v) for v in non_empty):
                if width == 2 and r == r1:
                    col2_low = str(vals[1]).strip().lower()
                    col1_low = str(vals[0]).strip().lower()
                    # Check if Row 1 is a genuine header ("Parameter | Value") vs first data row ("Basis | US GAAP")
                    next_col2 = ws.cell(row=r + 1, column=c1 + 1).value
                    is_header_word = (
                        col2_low in KEY_VALUE_HEADER_WORDS
                        or col1_low in KEY_VALUE_HEADER_WORDS
                        or any(w in col2_low for w in ("value", "amount", "input", "metric", "result"))
                    )
                    if is_header_word or (_is_number(next_col2) or _is_formula(next_col2)):
                        if is_header_word:
                            header_rows.append(r)
                        else:
                            # Headerless key-value table (Row 1 is text data like Accounting Basis -> US GAAP)
                            break
                    else:
                        break
                else:
                    header_rows.append(r)
                continue

            # Case H2: Numeric monotonic period/year axis header (e.g., ["Line Item ($)", 0, 1, 2, 3, 4, 5, 6] or [2020..2026])
            if len(num_val_cols) >= 2 and _is_monotonic_period_sequence(num_val_cols):
                if vals[0] is None or _is_text(vals[0]):
                    header_rows.append(r)
                    continue

            break

    # 5. Detect whether this block is an Assumptions / pure-input block
    body_start_r = (header_rows[-1] + 1) if header_rows else r1
    formula_count = 0
    num_input_count = 0
    for r in range(body_start_r, r2 + 1):
        for c in range(label_col + 1, c2 + 1):
            v = ws.cell(row=r, column=c).value
            if _is_formula(v):
                formula_count += 1
            elif _is_number(v):
                num_input_count += 1
    is_assumptions = (
        any(kw in title_low for kw in ("assumption", "input", "driver"))
        or (formula_count == 0 and num_input_count >= 3 and width <= 4)
    )

    # 6. Detect mid-table SUBHEADER rows, SUBTOTAL rows, and GRAND TOTAL rows
    subheader_rows: Set[int] = set()
    subtotal_rows: Set[int] = set()
    total_rows: Set[int] = set()

    for r in range(body_start_r, r2 + 1):
        lbl_raw = ws.cell(row=r, column=label_col).value
        lbl = str(lbl_raw).strip() if lbl_raw is not None else ""
        lbl_low = lbl.lower()
        row_vals = [ws.cell(row=r, column=c).value for c in range(label_col + 1, c2 + 1)]
        row_non_empty = [v for v in row_vals if v is not None]

        # Mid-table section sub-header (e.g. "CURRENT ASSETS" with blank value columns)
        if _is_text(lbl_raw) and not row_non_empty:
            subheader_rows.add(r)
            continue

        if is_assumptions or not lbl_low:
            continue

        if is_pct_label(lbl_low):
            continue

        # Hybrid Signal A: Language-agnostic formula AST vertical SUM / 3+ row vertical addition
        has_vertical_rollup = any(
            _formula_is_vertical_rollup(ws.cell(row=r, column=c).value, r, c)
            for c in range(label_col + 1, c2 + 1)
        )

        # Hybrid Signal B: Multilingual financial grand-total vs subtotal label keywords
        is_grand_kw = any(
            lbl_low == kw or lbl_low.startswith(kw) or f" {kw}" in lbl_low
            for kw in GRAND_TOTAL_KEYWORDS
        )
        is_sub_kw = any(
            lbl_low == kw or lbl_low.startswith(kw) or f" {kw}" in lbl_low
            for kw in SUBTOTAL_KEYWORDS
        )

        if is_grand_kw:
            total_rows.add(r)
        elif has_vertical_rollup:
            # Bottom-row vertical rollup or "total ..." -> Grand Total; intermediate rollup -> Subtotal
            if r == r2 or lbl_low.startswith("total"):
                total_rows.add(r)
            else:
                subtotal_rows.add(r)
        elif is_sub_kw:
            subtotal_rows.add(r)

    return TableBlock(
        kind="DATA_TABLE",
        min_row=r1,
        max_row=r2,
        min_col=c1,
        max_col=c2,
        label_col=label_col,
        code_col=code_col,
        header_rows=header_rows,
        subheader_rows=subheader_rows,
        subtotal_rows=subtotal_rows,
        total_rows=total_rows,
        is_assumptions_block=is_assumptions,
    )


def analyze_worksheet_layout(ws) -> SheetLayout:
    """Analyze any openpyxl worksheet into semantic 2D blocks without assuming Row 1 or Column A."""
    layout = SheetLayout(sheet_name=ws.title)
    raw_rects = _find_connected_components(ws)
    if not raw_rects:
        return layout

    refined_rects: List[Tuple[int, int, int, int]] = []
    for rect in raw_rects:
        refined_rects.extend(_peel_top_banners(ws, rect))

    for rect in refined_rects:
        blk = _classify_block(ws, rect, ws.title)
        layout.blocks.append(blk)

    # Identify top-most TITLE_BLOCK and parse any declared currency / magnitude in it
    for blk in layout.blocks:
        if blk.kind == "TITLE_BLOCK" and layout.title_block is None:
            layout.title_block = blk
        if blk.kind in ("TITLE_BLOCK", "SECTION_BANNER"):
            for r in range(blk.min_row, blk.max_row + 1):
                txt = ws.cell(row=r, column=blk.min_col).value
                # Mark as overflow banner cell ONLY if EVERY cell to the right on that row is empty
                # (if C3 has data while B3 is blank, A3 cannot overflow past B3 and must size Col A!)
                row_has_data_right = any(
                    ws.cell(row=r, column=rc).value is not None
                    for rc in range(blk.min_col + 1, (ws.max_column or blk.min_col) + 1)
                )
                if not row_has_data_right:
                    layout.overflow_banner_cells.add((r, blk.min_col))
                parsed = parse_magnitude_from_text(str(txt or ""))
                if parsed:
                    layout.declared_currency, layout.declared_magnitude = parsed

    # Also check single-cell Row 1 banner inside a DATA_TABLE (e.g. ["CapEx Project Analysis", None])
    for blk in layout.blocks:
        if blk.kind == "DATA_TABLE" and blk.header_rows:
            h0 = blk.header_rows[0]
            non_empty_cols = [c for c in range(blk.min_col, blk.max_col + 1) if ws.cell(row=h0, column=c).value is not None]
            if len(non_empty_cols) == 1 and non_empty_cols[0] == blk.min_col:
                parsed = parse_magnitude_from_text(str(ws.cell(row=h0, column=blk.min_col).value or ""))
                if parsed and layout.declared_magnitude is None:
                    layout.declared_currency, layout.declared_magnitude = parsed

    # Compute recommended freeze_panes cell from primary DATA_TABLE / MATRIX_TABLE
    data_blks = layout.data_blocks()
    if data_blks:
        # Choose the widest / largest primary table on the sheet
        primary = max(data_blks, key=lambda b: (b.max_row - b.min_row + 1) * (b.max_col - b.min_col + 1))
        freeze_row = (primary.primary_header_row + 1) if primary.primary_header_row is not None else primary.min_row
        freeze_col = primary.label_col + 1
        layout.recommended_freeze_cell = f"{get_column_letter(freeze_col)}{freeze_row}"

    return layout


def resolve_workbook_magnitude(
    wb,
    layouts: Optional[Dict[str, SheetLayout]] = None,
    header_metadata: Optional[Dict[str, Any]] = None,
) -> MagnitudeSpec:
    """Resolve consistent Smart Magnitude across an entire workbook (Option A — whole numbers)."""
    layouts = layouts or {ws.title: analyze_worksheet_layout(ws) for ws in wb.worksheets}
    meta = header_metadata or {}

    declared_currency: Optional[str] = meta.get("currency")
    declared_magnitude: Optional[str] = meta.get("magnitude")
    has_stored_millions_cue = False
    has_per_unit = False
    max_macro_num = 0.0

    for ws in wb.worksheets:
        lay = layouts.get(ws.title) or analyze_worksheet_layout(ws)
        if not declared_currency and lay.declared_currency:
            declared_currency = lay.declared_currency
        if not declared_magnitude and lay.declared_magnitude:
            declared_magnitude = lay.declared_magnitude

        stitle = (ws.title or "").lower()
        if any(kw in stitle for kw in ("($m)", "$m", "millions")):
            has_stored_millions_cue = True

        for blk in lay.data_blocks():
            hdr_set = set(blk.header_rows)
            for r in range(blk.min_row, blk.max_row + 1):
                lbl = str(ws.cell(row=r, column=blk.label_col).value or "").strip().lower()
                if any(kw in lbl for kw in ("($m)", "$m", "_m", "millions")):
                    has_stored_millions_cue = True
                if is_per_unit_label(lbl):
                    has_per_unit = True
                    continue
                if r in hdr_set or r in blk.subheader_rows:
                    continue
                if any(kw in lbl for kw in NON_CURRENCY_SCALE_KEYWORDS) or any(kw in lbl for kw in PCT_KEYWORDS):
                    continue
                for c in range(blk.value_start_col, blk.max_col + 1):
                    v = ws.cell(row=r, column=c).value
                    if _is_number(v) and abs(float(v)) >= 2:
                        max_macro_num = max(max_macro_num, abs(float(v)))

    return resolve_smart_magnitude(
        max_macro_amount=max_macro_num,
        currency=declared_currency or "USD",
        requested_magnitude=declared_magnitude or "auto",
        has_stored_millions_cue=has_stored_millions_cue,
        has_per_unit_rows=has_per_unit,
    )
