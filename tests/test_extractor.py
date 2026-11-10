"""
tests/test_extractor.py

Tests for GeminiExtractor validation and retry policy. The Gemini client is
mocked; no network or paid API calls are made.
"""

import json
from unittest.mock import MagicMock, patch

import httpx
import pytest
from google.genai import errors as genai_errors

from core.extractor import GeminiExtractor, _is_transient
from tests.test_excel_engine import _make_minimal_data


@pytest.fixture
def extractor():
    with patch("core.extractor.genai.Client"), \
         patch("core.extractor.get_system_prompt", return_value="prompt"):
        ex = GeminiExtractor(api_key="test-key")
    # Make retries instantaneous.
    for fn in (ex._call_text, ex._call_vision):
        fn.retry.wait = lambda *a, **k: 0
    return ex


def _respond(ex, text):
    ex.client.models.generate_content.return_value = MagicMock(text=text)


class TestValidation:
    def test_valid_payload_returned_validated(self, extractor):
        data = _make_minimal_data()
        data["fields"]["cash_fy25_lakhs"]["value"] = "1,234.5"  # str stays str in FieldBase
        data["fields"]["ae_revenue_split"]["value"]["domestic_ae_fy25"] = "1,000"
        _respond(extractor, json.dumps(data))
        res = extractor.extract_from_text("text", "Test Corp")
        assert res.success
        # Coerced: numeric string in a numeric-typed field becomes a number.
        assert res.data["fields"]["ae_revenue_split"]["value"]["domestic_ae_fy25"] == 1000.0
        # Optional keys are filled in by the model dump.
        assert "flag" in res.data["fields"]["hq_group"]

    def test_missing_required_field_reports_error(self, extractor):
        data = _make_minimal_data()
        del data["fields"]["shareholding"]
        _respond(extractor, json.dumps(data))
        res = extractor.extract_from_text("text", "Test Corp")
        assert not res.success
        assert res.data is None
        assert "Schema validation failed" in res.error
        assert "fields.shareholding" in res.error

    def test_empty_shareholder_name_reports_error(self, extractor):
        data = _make_minimal_data()
        data["fields"]["shareholding"]["rows"][0]["name"] = "  "
        _respond(extractor, json.dumps(data))
        res = extractor.extract_from_text("text", "Test Corp")
        assert not res.success
        assert "name" in res.error

    def test_wrong_json_shape_reports_error(self, extractor):
        _respond(extractor, json.dumps({"entity": "X"}))
        res = extractor.extract_from_text("text", "X")
        assert not res.success and "Schema validation failed" in res.error

    def test_malformed_json_still_reports_parse_error(self, extractor):
        _respond(extractor, "{not json")
        res = extractor.extract_from_text("text", "X")
        assert not res.success and res.error.startswith("JSON parse error")

    def test_vision_mode_validates_and_cleans_up(self, extractor):
        data = _make_minimal_data()
        del data["entity"]
        _respond(extractor, json.dumps(data))
        extractor.client.files.upload.return_value = MagicMock(name="f")
        res = extractor.extract_from_pdf("x.pdf", "Test Corp")
        assert not res.success and res.mode == "vision"
        assert "Schema validation failed" in res.error
        extractor.client.files.delete.assert_called_once()


class TestRetryPolicy:
    def test_is_transient_classification(self):
        assert _is_transient(genai_errors.ClientError(429, {}))
        assert _is_transient(genai_errors.ClientError(408, {}))
        assert _is_transient(genai_errors.ServerError(503, {}))
        assert _is_transient(httpx.ConnectError("boom"))
        assert _is_transient(TimeoutError())
        assert not _is_transient(genai_errors.ClientError(400, {}))
        assert not _is_transient(genai_errors.ClientError(401, {}))
        assert not _is_transient(genai_errors.ClientError(403, {}))
        assert not _is_transient(ValueError("nope"))

    def test_auth_error_not_retried(self, extractor):
        extractor.client.models.generate_content.side_effect = genai_errors.ClientError(401, {})
        res = extractor.extract_from_text("text", "X")
        assert not res.success
        assert extractor.client.models.generate_content.call_count == 1

    def test_rate_limit_retried_then_succeeds(self, extractor):
        ok = MagicMock(text=json.dumps(_make_minimal_data()))
        extractor.client.models.generate_content.side_effect = [
            genai_errors.ClientError(429, {}),
            genai_errors.ServerError(503, {}),
            ok,
        ]
        res = extractor.extract_from_text("text", "Test Corp")
        assert res.success
        assert extractor.client.models.generate_content.call_count == 3

    def test_persistent_server_error_gives_up_after_5(self, extractor):
        extractor.client.models.generate_content.side_effect = genai_errors.ServerError(500, {})
        res = extractor.extract_from_text("text", "X")
        assert not res.success
        assert extractor.client.models.generate_content.call_count == 5


def test_validated_output_feeds_fill_factsheet(extractor, tmp_path):
    """The coerced dict returned by the extractor must work downstream."""
    from core.excel_engine import fill_factsheet
    from tests.test_excel_engine import _create_minimal_template

    _respond(extractor, json.dumps(_make_minimal_data()))
    res = extractor.extract_from_text("text", "Test Corp")
    assert res.success
    template = _create_minimal_template(str(tmp_path / "t.xlsx"))
    out = str(tmp_path / "out.xlsx")
    assert fill_factsheet(res.data, template, out, "Sheet1").success
