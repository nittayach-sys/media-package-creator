"""
pricing_engine.py
==================
Layer C of the architecture (DETERMINISTIC PRICING ENGINE).

Pure arithmetic, no AI, no estimation. Every formula here matches the
confirmed rules:

    Total Media Cost       = Agency Rate x Duration (Month)
    Total Media Investment = Total Media Cost + Production
    TAX3 (Q1-Q4)            = looked up from the database, never computed

A LineItem is what one row of Sales' selections looks like before it is
written into the Excel template.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, asdict
from typing import Optional

from data_loader import MediaRecord, NO_DATA


def _to_number(val) -> Optional[float]:
    """Best-effort parse of a numeric value out of a database cell. Returns
    None (never 0, never a guess) if the cell isn't a clean number — e.g.
    '155,000 / Escalator' or '1 Side : 12,936...' can't be safely reduced to
    a single number without inventing a choice, so pricing must stop and
    surface that to the Sales user rather than silently pick one."""
    if val is None or val == NO_DATA:
        return None
    s = str(val).strip()
    s_clean = s.replace(",", "")
    m = re.fullmatch(r"-?\d+(\.\d+)?", s_clean)
    if m:
        return float(s_clean)
    return None


@dataclass
class LineItem:
    record: MediaRecord
    period: str            # Sales-entered, free text
    duration_months: int   # Sales-selected, 1-12 (or custom)

    def _num(self, field_value: str) -> Optional[float]:
        return _to_number(field_value)

    @property
    def normal_rate(self) -> Optional[float]:
        return self._num(self.record.rate_card)

    @property
    def agency_rate(self) -> Optional[float]:
        return self._num(self.record.agency_rate)

    @property
    def production(self) -> float:
        # Rule: Production is used exactly as given, never multiplied by
        # duration. If absent/unparseable, it contributes 0 to totals but
        # displays as "-" in Excel (handled by excel_generator).
        v = self._num(self.record.production)
        return v if v is not None else 0.0

    @property
    def total_media_cost(self) -> Optional[float]:
        ar = self.agency_rate
        if ar is None:
            return None
        return ar * self.duration_months

    @property
    def total_media_investment(self) -> Optional[float]:
        tmc = self.total_media_cost
        if tmc is None:
            return None
        return tmc + self.production

    @property
    def needs_manual_rate_review(self) -> bool:
        """True if the Rate Card / Agency Rate cell couldn't be reduced to a
        clean number (e.g. tiered text like '1 Side : ... / 2 Sides : ...').
        Sales must resolve this manually before the total can be trusted —
        the system must never guess which tier applies."""
        return self.agency_rate is None

    def tax3(self) -> dict:
        return {
            "Q1": self.record.tax3_q1,
            "Q2": self.record.tax3_q2,
            "Q3": self.record.tax3_q3,
            "Q4": self.record.tax3_q4,
        }

    def to_row_dict(self) -> dict:
        """Everything needed to write one Excel row, already computed."""
        rec = self.record
        tmc = self.total_media_cost
        tmi = self.total_media_investment
        return {
            "type_of_media": rec.type_of_media,
            "media": rec.media,
            "quantity": rec.quantity,
            "detail": rec.min_loop if rec.min_loop != NO_DATA else rec.size,
            "period": self.period,
            "duration": self.duration_months,
            "normal_rate": self.normal_rate if self.normal_rate is not None else NO_DATA,
            "agency_rate_secondary": self.agency_rate if self.agency_rate is not None else NO_DATA,  # mirrors J, per confirmed rule
            "agency_rate": self.agency_rate if self.agency_rate is not None else NO_DATA,
            "total_media_cost": tmc if tmc is not None else NO_DATA,
            "production": self.production if self._num(rec.production) is not None else NO_DATA,
            "total_media_investment": tmi if tmi is not None else NO_DATA,
            "remark": rec.remark,
            "tax3_q1": rec.tax3_q1,
            "tax3_q2": rec.tax3_q2,
            "tax3_q3": rec.tax3_q3,
            "tax3_q4": rec.tax3_q4,
            "media_proposal_url": rec.media_proposal_url,
            "material_guideline_url": rec.material_guideline_url,
            "needs_manual_rate_review": self.needs_manual_rate_review,
        }


def build_package_totals(line_items: list[LineItem]) -> dict:
    """Sum every numeric column for the package Total row. Any line whose
    rate couldn't be resolved is excluded from the sum and flagged
    separately rather than silently treated as zero."""
    fields = ["normal_rate", "agency_rate", "total_media_cost", "production",
              "total_media_investment", "tax3_q1", "tax3_q2", "tax3_q3", "tax3_q4"]
    totals = {f: 0.0 for f in fields}
    unresolved = []
    for li in line_items:
        row = li.to_row_dict()
        if li.needs_manual_rate_review:
            unresolved.append(li)
            continue
        for f in ["normal_rate", "agency_rate", "total_media_cost", "production", "total_media_investment"]:
            v = row[f]
            if isinstance(v, (int, float)):
                totals[f] += v
        for q in ["tax3_q1", "tax3_q2", "tax3_q3", "tax3_q4"]:
            v = _to_number(row[q])
            if v is not None:
                totals[q] += v
    return {"totals": totals, "unresolved": unresolved}
