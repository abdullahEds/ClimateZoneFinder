"""Pytest test suite for the FastAPI report API (report_api.py)."""

import pathlib
import pytest
import httpx

from report_api import app

_EPW_PATH = pathlib.Path(__file__).parent / "IND_DL_New.Delhi-Safdarjung.AP.421820_ISHRAE2014.epw"
_EPW_BYTES = _EPW_PATH.read_bytes() if _EPW_PATH.exists() else None

_PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

_SKIP_NO_FILE = pytest.mark.skipif(
    _EPW_BYTES is None,
    reason="Bundled EPW file not found",
)

pytestmark = pytest.mark.anyio


async def _get(path: str, **kwargs):
    async with httpx.AsyncClient(
        transport=httpx.ASGITransport(app=app),
        base_url="http://testserver",
    ) as client:
        return await client.get(path, **kwargs)


async def _post(path: str, **kwargs):
    async with httpx.AsyncClient(
        transport=httpx.ASGITransport(app=app),
        base_url="http://testserver",
    ) as client:
        return await client.post(path, **kwargs)


# ── Health / docs ─────────────────────────────────────────────────────────────

async def test_health():
    resp = await _get("/api/health")
    assert resp.status_code == 200
    assert resp.json()["status"] == "ok"


async def test_docs_endpoint():
    resp = await _get("/api/docs")
    assert resp.status_code == 200
    body = resp.json()
    assert "title" in body
    assert "endpoints" in body


# ── Stations ──────────────────────────────────────────────────────────────────

async def test_rainfall_stations():
    resp = await _get("/api/rainfall/stations")
    assert resp.status_code == 200
    data = resp.json()
    assert "stations" in data
    assert len(data["stations"]) >= 10


# ── Climate report ────────────────────────────────────────────────────────────

@_SKIP_NO_FILE
async def test_climate_report_valid_epw():
    resp = await _post(
        "/api/reports/climate-analysis",
        files={"file": ("test.epw", _EPW_BYTES, "application/octet-stream")},
    )
    assert resp.status_code == 200
    assert resp.headers["content-type"].startswith(_PPTX_MIME)
    cd = resp.headers.get("content-disposition", "")
    assert ".pptx" in cd


async def test_climate_report_invalid_file():
    resp = await _post(
        "/api/reports/climate-analysis",
        files={"file": ("bad.txt", b"this is not an epw file", "text/plain")},
    )
    assert resp.status_code == 400


# ── Combined report ───────────────────────────────────────────────────────────

@_SKIP_NO_FILE
async def test_combined_report_valid():
    resp = await _post(
        "/api/reports/combined-analysis",
        files={"file": ("test.epw", _EPW_BYTES, "application/octet-stream")},
    )
    assert resp.status_code == 200
    assert resp.headers["content-type"].startswith(_PPTX_MIME)


@_SKIP_NO_FILE
async def test_combined_report_invalid_rainfall_station():
    resp = await _post(
        "/api/reports/combined-analysis",
        data={"rainfall_station_name": "INVALID_STATION", "rainfall_year": "2023"},
        files={"file": ("test.epw", _EPW_BYTES, "application/octet-stream")},
    )
    assert resp.status_code == 400
    body = resp.json()
    assert "detail" in body


# ── Wind report ───────────────────────────────────────────────────────────────

@_SKIP_NO_FILE
async def test_wind_report_invalid_sectors():
    resp = await _post(
        "/api/reports/wind-analysis",
        data={"n_sectors": "7"},
        files={"file": ("test.epw", _EPW_BYTES, "application/octet-stream")},
    )
    assert resp.status_code == 400
    body = resp.json()
    assert "detail" in body


# ── Error response schema ─────────────────────────────────────────────────────

async def test_error_response_schema_400():
    resp = await _post(
        "/api/reports/climate-analysis",
        files={"file": ("bad.txt", b"not epw", "text/plain")},
    )
    assert resp.status_code == 400
    body = resp.json()
    detail = body.get("detail")
    assert isinstance(detail, dict), f"Expected dict detail, got: {type(detail)}"
    assert "error" in detail
    assert "detail" in detail
