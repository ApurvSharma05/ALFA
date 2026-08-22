"""
tests/test_schemas.py

Tests for Pydantic schema parsing, validation, and edge-case handling.
"""

import pytest
from core.schemas import (
    FieldBase,
    AERevenueSplitValue,
    ShareholderRow,
    RPTItem,
    LitigationItem,
    CompanyData,
    _coerce_numeric,
)


# ---------------------------------------------------------------------------
# _coerce_numeric helper
# ---------------------------------------------------------------------------

class TestCoerceNumeric:
    def test_none_returns_none(self):
        assert _coerce_numeric(None) is None

    def test_int_passthrough(self):
        assert _coerce_numeric(42) == 42

    def test_float_passthrough(self):
        assert _coerce_numeric(3.14) == 3.14

    def test_string_number(self):
        assert _coerce_numeric("1234.56") == 1234.56

    def test_string_with_commas(self):
        assert _coerce_numeric("1,234,567.89") == 1234567.89

    def test_string_dash_returns_none(self):
        assert _coerce_numeric("-") is None

    def test_string_na_returns_none(self):
        assert _coerce_numeric("N/A") is None
        assert _coerce_numeric("n/a") is None
        assert _coerce_numeric("NA") is None

    def test_string_null_returns_none(self):
        assert _coerce_numeric("null") is None
        assert _coerce_numeric("None") is None

    def test_empty_string_returns_none(self):
        assert _coerce_numeric("") is None
        assert _coerce_numeric("   ") is None

    def test_garbage_string_returns_none(self):
        assert _coerce_numeric("not a number") is None


# ---------------------------------------------------------------------------
# FieldBase
# ---------------------------------------------------------------------------

class TestFieldBase:
    def test_valid_with_all_fields(self):
        fb = FieldBase(value="Test", source="Page 5", flag="high confidence")
        assert fb.value == "Test"
        assert fb.source == "Page 5"
        assert fb.flag == "high confidence"

    def test_value_none(self):
        fb = FieldBase(value=None, source="Page 1")
        assert fb.value is None
        assert fb.flag is None

    def test_value_numeric(self):
        fb = FieldBase(value=12345.67, source="P&L")
        assert fb.value == 12345.67


# ---------------------------------------------------------------------------
# AERevenueSplitValue — validators
# ---------------------------------------------------------------------------

class TestAERevenueSplitValue:
    def test_normal_numbers(self):
        v = AERevenueSplitValue(
            domestic_ae_fy25=100.0,
            export_ae_fy25=200.0,
            domestic_third_party_fy25=300.0,
            export_third_party_fy25=400.0,
        )
        assert v.domestic_ae_fy25 == 100.0
        assert v.export_ae_fy25 == 200.0

    def test_string_coercion(self):
        v = AERevenueSplitValue(
            domestic_ae_fy25="1,234.56",
            export_ae_fy25="N/A",
            domestic_third_party_fy25=None,
            export_third_party_fy25="",
        )
        assert v.domestic_ae_fy25 == 1234.56
        assert v.export_ae_fy25 is None
        assert v.domestic_third_party_fy25 is None
        assert v.export_third_party_fy25 is None


# ---------------------------------------------------------------------------
# ShareholderRow — validators
# ---------------------------------------------------------------------------

class TestShareholderRow:
    def test_valid(self):
        sr = ShareholderRow(name="John Doe", shares_fy25=10000, shares_fy24=9000)
        assert sr.name == "John Doe"

    def test_empty_name_rejected(self):
        with pytest.raises(Exception):
            ShareholderRow(name="", shares_fy25=100, shares_fy24=100)

    def test_whitespace_name_rejected(self):
        with pytest.raises(Exception):
            ShareholderRow(name="   ", shares_fy25=100, shares_fy24=100)

    def test_name_trimmed(self):
        sr = ShareholderRow(name="  Padded Name  ", shares_fy25=100, shares_fy24=100)
        assert sr.name == "Padded Name"

    def test_string_shares_coerced(self):
        sr = ShareholderRow(name="Test", shares_fy25="1,000", shares_fy24="N/A")
        assert sr.shares_fy25 == 1000.0
        assert sr.shares_fy24 is None


# ---------------------------------------------------------------------------
# RPTItem — validators
# ---------------------------------------------------------------------------

class TestRPTItem:
    def test_valid(self):
        item = RPTItem(label="Business Support Services", value_fy25=-500.0)
        assert item.label == "Business Support Services"
        assert item.value_fy25 == -500.0

    def test_empty_label_rejected(self):
        with pytest.raises(Exception):
            RPTItem(label="", value_fy25=100)

    def test_label_trimmed(self):
        item = RPTItem(label="  Rent Paid  ", value_fy25=200)
        assert item.label == "Rent Paid"

    def test_string_value_coerced(self):
        item = RPTItem(label="Test", value_fy25="1,234.56")
        assert item.value_fy25 == 1234.56


# ---------------------------------------------------------------------------
# LitigationItem — validators
# ---------------------------------------------------------------------------

class TestLitigationItem:
    def test_all_none(self):
        item = LitigationItem()
        assert item.nature_of_dues is None
        assert item.amount_demanded_lakhs is None

    def test_string_amounts_coerced(self):
        item = LitigationItem(
            nature_of_dues="Income Tax",
            amount_demanded_lakhs="5,432.10",
            amount_paid_lakhs="null",
        )
        assert item.amount_demanded_lakhs == 5432.10
        assert item.amount_paid_lakhs is None


# ---------------------------------------------------------------------------
# CompanyData — full schema round-trip
# ---------------------------------------------------------------------------

class TestCompanyData:
    def _make_field(self, value=None, source="Test"):
        return {"value": value, "source": source}

    def test_minimal_valid(self):
        """The minimum viable CompanyData should parse without error."""
        data = {
            "entity": "Test Corp",
            "fields": {
                "hq_india_entity": self._make_field("Mumbai"),
                "hq_group": self._make_field("N/A"),
                "company_description": self._make_field("A test company"),
                "group_description": self._make_field("N/A"),
                "standalone_turnover_fy25_lakhs": self._make_field(1000),
                "standalone_turnover_fy24_lakhs": self._make_field(900),
                "standalone_total_cost_fy25_lakhs": self._make_field(800),
                "standalone_total_cost_fy24_lakhs": self._make_field(700),
                "consolidated_summary": self._make_field("N/A"),
                "statutory_auditors": self._make_field("Deloitte"),
                "ae_revenue_split": {
                    "value": {
                        "domestic_ae_fy25": 100,
                        "export_ae_fy25": 200,
                        "domestic_third_party_fy25": 300,
                        "export_third_party_fy25": 400,
                    },
                    "source": "Page 45",
                },
                "cash_fy25_lakhs": self._make_field(500),
                "cash_fy24_lakhs": self._make_field(400),
                "ae_trade_receivables_fy25_lakhs": self._make_field(50),
                "ae_trade_receivables_fy24_lakhs": self._make_field(40),
                "pe_investment": self._make_field("None"),
                "shareholding": {
                    "rows": [
                        {"name": "Director A", "shares_fy25": 5000, "shares_fy24": 5000},
                        {"name": "Others", "shares_fy25": 5000, "shares_fy24": 5000},
                    ],
                    "total_shares_fy25": 10000,
                    "total_shares_fy24": 10000,
                    "source": "Page 30",
                },
                "related_party_transactions_lakhs": {
                    "items": [
                        {"label": "Rent", "value_fy25": 100},
                    ],
                    "source": "Page 50",
                },
                "countries_presence": self._make_field("India"),
                "litigation": {
                    "items": [],
                    "source": "No litigation noted",
                },
                "website": self._make_field("https://example.com"),
                "linkedin": self._make_field(None),
            },
        }
        cd = CompanyData(**data)
        assert cd.entity == "Test Corp"
        assert len(cd.fields.shareholding.rows) == 2
        assert len(cd.fields.related_party_transactions_lakhs.items) == 1
