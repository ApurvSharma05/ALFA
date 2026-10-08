"""
tests/test_excel_engine.py

Tests for the excel_engine module: fill_factsheet with mock data and templates.
"""

import os
import json
import tempfile
import pytest
from openpyxl import Workbook, load_workbook

from core.excel_engine import fill_factsheet, FillResult
from config.template_mapping import FactSheetMapping


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

def _make_field(value=None, source="Test Source"):
    return {"value": value, "source": source}


def _make_minimal_data(
    entity: str = "Test Corp",
    num_shareholders: int = 2,
    num_rpt: int = 1,
    num_litigation: int = 0,
) -> dict:
    """Build a minimal CompanyData-shaped dict."""
    shareholders = [
        {"name": f"Shareholder {i+1}", "shares_fy25": 1000 * (i+1), "shares_fy24": 900 * (i+1)}
        for i in range(num_shareholders)
    ]
    rpt_items = [
        {"label": f"Transaction Type {i+1}", "value_fy25": 100.0 * (i+1)}
        for i in range(num_rpt)
    ]
    lit_items = [
        {
            "nature_of_dues": f"Tax Dispute {i+1}",
            "amount_demanded_lakhs": 500.0 * (i+1),
            "amount_paid_lakhs": 50.0 * (i+1),
            "period": "FY2023-24",
            "forum": "ITAT",
        }
        for i in range(num_litigation)
    ]

    return {
        "entity": entity,
        "fields": {
            "hq_india_entity": _make_field("123 Test Street, Mumbai"),
            "hq_group": _make_field("N/A"),
            "company_description": _make_field("A test company."),
            "group_description": _make_field("N/A"),
            "standalone_turnover_fy25_lakhs": _make_field(10000),
            "standalone_turnover_fy24_lakhs": _make_field(9000),
            "standalone_total_cost_fy25_lakhs": _make_field(8000),
            "standalone_total_cost_fy24_lakhs": _make_field(7000),
            "consolidated_summary": _make_field("N/A"),
            "statutory_auditors": _make_field("Deloitte Haskins"),
            "ae_revenue_split": {
                "value": {
                    "domestic_ae_fy25": 100,
                    "export_ae_fy25": 200,
                    "domestic_third_party_fy25": 300,
                    "export_third_party_fy25": 400,
                },
                "source": "Page 45",
                "flag": "Cross-checked",
            },
            "cash_fy25_lakhs": _make_field(5000),
            "cash_fy24_lakhs": _make_field(4000),
            "ae_trade_receivables_fy25_lakhs": _make_field(500),
            "ae_trade_receivables_fy24_lakhs": _make_field(400),
            "pe_investment": _make_field("None", "Page 60"),
            "shareholding": {
                "rows": shareholders,
                "total_shares_fy25": sum(s["shares_fy25"] for s in shareholders),
                "total_shares_fy24": sum(s["shares_fy24"] for s in shareholders),
                "source": "Page 30",
            },
            "related_party_transactions_lakhs": {
                "items": rpt_items,
                "source": "Page 50",
            },
            "countries_presence": _make_field("India"),
            "litigation": {
                "items": lit_items,
                "source": "Page 55",
            },
            "website": _make_field("https://example.com"),
            "linkedin": _make_field("https://linkedin.com/company/test"),
        },
    }


def _create_minimal_template(path: str, sheet_name: str = "Sheet1"):
    """
    Create a bare-bones Excel template with the expected structure.
    The key requirement is a 'Total' cell in column C within the
    shareholding range.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    # Header
    ws["A1"] = "Fact Sheet Template"

    # Shareholding markers
    ws["C40"] = "Shareholding"
    ws["C42"] = "Row 1"
    ws["C43"] = "Row 2"
    ws["C44"] = "Total"  # shareholding_default_capacity = 2 (rows 42-43)
    ws["D44"] = 0
    ws["F44"] = 0

    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

class TestFillFactsheet:
    def test_basic_fill(self):
        """fill_factsheet with minimal data should succeed."""
        with tempfile.TemporaryDirectory() as td:
            template = _create_minimal_template(os.path.join(td, "template.xlsx"))
            data = _make_minimal_data()
            out = os.path.join(td, "out.xlsx")

            result = fill_factsheet(data, template, out, "Sheet1")

            assert result.success
            assert result.company_name == "Test Corp"
            assert result.shareholding_rows_written == 2
            assert result.rpt_rows_written == 1
            assert result.rpt_rows_dropped == 0

            # Verify the output file was created
            assert os.path.exists(out)

            # Verify some cells
            wb = load_workbook(out)
            ws = wb[data["entity"][:31]]
            assert ws["A1"].value == "Fact Sheet - Test Corp"
            assert ws["C3"].value == "123 Test Street, Mumbai"
            assert ws["D12"].value == 10000

    def test_source_comments_attached(self):
        """Every filled cell should have a source-citing comment."""
        with tempfile.TemporaryDirectory() as td:
            template = _create_minimal_template(os.path.join(td, "template.xlsx"))
            data = _make_minimal_data()
            out = os.path.join(td, "out.xlsx")

            fill_factsheet(data, template, out, "Sheet1")

            wb = load_workbook(out)
            ws = wb[data["entity"][:31]]

            # HQ cell should have a comment
            assert ws["C3"].comment is not None
            assert "Source:" in ws["C3"].comment.text

    def test_rpt_all_items_written_after_expansion(self):
        """
        After expand_dynamic_tables() pre-expands the template, fill_factsheet()
        should write ALL RPT items with zero drops.
        This validates the end-to-end pipeline that replaced the old silent-drop behaviour.
        """
        from core.excel_engine import expand_dynamic_tables

        with tempfile.TemporaryDirectory() as td:
            template = _create_minimal_template(os.path.join(td, "template.xlsx"))
            data = _make_minimal_data(num_rpt=10)  # exceeds default capacity of 7
            expanded = os.path.join(td, "expanded.xlsx")
            out = os.path.join(td, "out.xlsx")

            # Step 1: pre-expand
            template_to_use, delta = expand_dynamic_tables(
                data=data,
                template_path=template,
                out_path=expanded,
                sheet_name="Sheet1",
            )
            assert delta == 3  # 10 - 7 = 3 rows added

            # Step 2: fill
            result = fill_factsheet(data, template_to_use, out, "Sheet1")

            assert result.success
            assert result.rpt_rows_written == 10
            assert result.rpt_rows_dropped == 0
            assert result.warnings == []  # no overflow warnings

    def test_rpt_fill_without_expansion_no_crash(self):
        """
        fill_factsheet() called alone (without pre-expansion) on a template that
        is too small should still succeed — it will write as many items as the
        default capacity allows and produce zero drops because fill no longer
        enforces the old cap (that responsibility moved to expand_dynamic_tables).
        This ensures backwards-compatibility for callers that bypass expansion.
        """
        with tempfile.TemporaryDirectory() as td:
            template = _create_minimal_template(os.path.join(td, "template.xlsx"))
            data = _make_minimal_data(num_rpt=10)
            out = os.path.join(td, "out.xlsx")

            # Should not raise — fill writes all items into available rows.
            result = fill_factsheet(data, template, out, "Sheet1")
            assert result.success
            assert result.rpt_rows_dropped == 0  # no silent drops

    def test_unused_rows_hidden(self):
        """When fewer items than capacity, unused rows should be hidden."""
        with tempfile.TemporaryDirectory() as td:
            template = _create_minimal_template(os.path.join(td, "template.xlsx"))
            data = _make_minimal_data(num_rpt=3)  # 3 < 7 capacity
            out = os.path.join(td, "out.xlsx")

            result = fill_factsheet(data, template, out, "Sheet1")
            assert result.success

            wb = load_workbook(out)
            ws = wb[data["entity"][:31]]

            # Some rows should be hidden
            hidden_count = sum(
                1 for r in range(1, ws.max_row + 1)
                if ws.row_dimensions[r].hidden
            )
            assert hidden_count > 0

    def test_custom_mapping(self):
        """A custom FactSheetMapping should be respected."""
        mapping = FactSheetMapping(hq_india_cell="C5")  # moved from C3 to C5

        with tempfile.TemporaryDirectory() as td:
            template = _create_minimal_template(os.path.join(td, "template.xlsx"))
            data = _make_minimal_data()
            out = os.path.join(td, "out.xlsx")

            result = fill_factsheet(data, template, out, "Sheet1", mapping=mapping)
            assert result.success

            wb = load_workbook(out)
            ws = wb[data["entity"][:31]]
            # HQ should be at C5 now, not C3
            assert ws["C5"].value == "123 Test Street, Mumbai"

    def test_fill_result_metadata(self):
        """FillResult should contain accurate metadata."""
        with tempfile.TemporaryDirectory() as td:
            template = _create_minimal_template(os.path.join(td, "template.xlsx"))
            data = _make_minimal_data(num_shareholders=2, num_rpt=5, num_litigation=1)
            out = os.path.join(td, "out.xlsx")

            result = fill_factsheet(data, template, out, "Sheet1")

            assert result.success
            assert result.shareholding_rows_written == 2
            assert result.rpt_rows_written == 5
            assert result.rpt_rows_dropped == 0
            assert result.lit_rows_written == 1
            assert result.lit_rows_dropped == 0
