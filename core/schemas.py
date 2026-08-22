"""
core/schemas.py

Pydantic v2 data models defining the strict JSON schema that Gemini must
return.  Moved here from app.py so every module can import them without
pulling in Streamlit.

Validators handle the messy edges of LLM output:
  - string numbers coerced to float
  - empty / whitespace-only names rejected
  - None propagated gracefully for missing fields
"""

from __future__ import annotations

from pydantic import BaseModel, Field, field_validator
from typing import List, Optional


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _coerce_numeric(v):
    """Attempt to coerce a value to float; return None on failure."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        cleaned = v.strip().replace(",", "").replace(" ", "")
        if not cleaned or cleaned in ("-", "N/A", "n/a", "NA", "null", "None"):
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


# ---------------------------------------------------------------------------
# Base field wrapper — every extracted value has a source citation
# ---------------------------------------------------------------------------

class FieldBase(BaseModel):
    """A single extracted value with provenance metadata."""
    value: str | int | float | None = Field(
        description="The extracted value. Use None if not found."
    )
    source: str = Field(
        description="Exact page number and section where this was found."
    )
    flag: Optional[str] = Field(
        None,
        description="Any judgment calls, assumptions, or notes for the human reviewer.",
    )


# ---------------------------------------------------------------------------
# AE Revenue Split
# ---------------------------------------------------------------------------

class AERevenueSplitValue(BaseModel):
    domestic_ae_fy25: float | int | None = None
    export_ae_fy25: float | int | None = None
    domestic_third_party_fy25: float | int | None = None
    export_third_party_fy25: float | int | None = None

    @field_validator("*", mode="before")
    @classmethod
    def coerce_numbers(cls, v):
        return _coerce_numeric(v)


class AERevenueSplit(BaseModel):
    value: AERevenueSplitValue
    source: str
    flag: Optional[str] = None


# ---------------------------------------------------------------------------
# Shareholding
# ---------------------------------------------------------------------------

class ShareholderRow(BaseModel):
    name: str = Field(description="Name of the shareholder/promoter category")
    shares_fy25: int | float | None = None
    shares_fy24: int | float | None = None

    @field_validator("name", mode="before")
    @classmethod
    def name_must_not_be_empty(cls, v):
        if not v or not str(v).strip():
            raise ValueError("Shareholder name must not be empty")
        return str(v).strip()

    @field_validator("shares_fy25", "shares_fy24", mode="before")
    @classmethod
    def coerce_shares(cls, v):
        return _coerce_numeric(v)


class Shareholding(BaseModel):
    rows: List[ShareholderRow]
    total_shares_fy25: int | float | None = None
    total_shares_fy24: int | float | None = None
    source: str


# ---------------------------------------------------------------------------
# Related Party Transactions
# ---------------------------------------------------------------------------

class RPTItem(BaseModel):
    label: str = Field(description="Exact transaction label from the AR")
    value_fy25: float | int | None = Field(
        None,
        description="Positive for outflow/cost, Negative for inflow/receivable",
    )

    @field_validator("label", mode="before")
    @classmethod
    def label_must_not_be_empty(cls, v):
        if not v or not str(v).strip():
            raise ValueError("RPT label must not be empty")
        return str(v).strip()

    @field_validator("value_fy25", mode="before")
    @classmethod
    def coerce_value(cls, v):
        return _coerce_numeric(v)


class RelatedPartyTransactions(BaseModel):
    items: List[RPTItem]
    source: str


# ---------------------------------------------------------------------------
# Litigation
# ---------------------------------------------------------------------------

class LitigationItem(BaseModel):
    nature_of_dues: str | None = None
    amount_demanded_lakhs: float | int | None = None
    amount_paid_lakhs: float | int | None = None
    period: str | None = None
    forum: str | None = None

    @field_validator("amount_demanded_lakhs", "amount_paid_lakhs", mode="before")
    @classmethod
    def coerce_amounts(cls, v):
        return _coerce_numeric(v)


class Litigation(BaseModel):
    items: List[LitigationItem]
    source: str


# ---------------------------------------------------------------------------
# Top-level extraction schema
# ---------------------------------------------------------------------------

class ExtractedFields(BaseModel):
    hq_india_entity: FieldBase
    hq_group: FieldBase
    company_description: FieldBase
    group_description: FieldBase
    standalone_turnover_fy25_lakhs: FieldBase
    standalone_turnover_fy24_lakhs: FieldBase
    standalone_total_cost_fy25_lakhs: FieldBase
    standalone_total_cost_fy24_lakhs: FieldBase
    consolidated_summary: FieldBase
    statutory_auditors: FieldBase
    ae_revenue_split: AERevenueSplit
    cash_fy25_lakhs: FieldBase
    cash_fy24_lakhs: FieldBase
    ae_trade_receivables_fy25_lakhs: FieldBase
    ae_trade_receivables_fy24_lakhs: FieldBase
    pe_investment: FieldBase
    shareholding: Shareholding
    related_party_transactions_lakhs: RelatedPartyTransactions
    countries_presence: FieldBase
    litigation: Litigation
    website: FieldBase
    linkedin: FieldBase


class CompanyData(BaseModel):
    """Root schema returned by Gemini for one company."""
    entity: str = Field(description="Name of the company")
    fields: ExtractedFields
