import io
import json

import pytest

from security.session import SessionManager
from server import MAX_REQUEST_SIZE, McpHandler, SUPPORTED_PROTOCOL_VERSION
from tools.manager import PptxTools


class HarnessHandler(McpHandler):
    def __init__(self, *, tmp_path, path, headers=None, body=b"", token="secret-token"):
        self.path = path
        self.headers = headers or {}
        self.rfile = io.BytesIO(body)
        self.wfile = io.BytesIO()
        self.status_code = None
        self.error_message = None
        self.sent_headers = {}
        self.session_manager = SessionManager()
        self.tools = PptxTools(self.session_manager, str(tmp_path))
        self.token = token

    def send_response(self, code, message=None):
        self.status_code = code

    def send_header(self, keyword, value):
        self.sent_headers[keyword] = value

    def end_headers(self):
        return None

    def send_error(self, code, message=None, explain=None):
        self.status_code = code
        self.error_message = message


def build_post_handler(tmp_path, payload, *, headers=None):
    raw_body = json.dumps(payload).encode("utf-8")
    request_headers = {
        "Authorization": "Bearer secret-token",
        "Content-Length": str(len(raw_body)),
    }
    if headers:
        request_headers.update(headers)

    return HarnessHandler(
        tmp_path=tmp_path,
        path="/mcp",
        headers=request_headers,
        body=raw_body,
    )


def test_initialize_notification_returns_accepted_without_body(tmp_path):
    initialize_handler = build_post_handler(
        tmp_path,
        {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {"protocolVersion": SUPPORTED_PROTOCOL_VERSION},
        },
    )

    initialize_handler.do_POST()
    body = json.loads(initialize_handler.wfile.getvalue().decode("utf-8"))

    assert initialize_handler.status_code == 200
    assert body["result"]["protocolVersion"] == SUPPORTED_PROTOCOL_VERSION

    notification_handler = build_post_handler(
        tmp_path,
        {
            "jsonrpc": "2.0",
            "method": "notifications/initialized",
            "params": {},
        },
        headers={"MCP-Protocol-Version": SUPPORTED_PROTOCOL_VERSION},
    )

    notification_handler.do_POST()

    assert notification_handler.status_code == 202
    assert notification_handler.wfile.getvalue() == b""


def test_ping_returns_empty_result(tmp_path):
    handler = build_post_handler(
        tmp_path,
        {"jsonrpc": "2.0", "id": 7, "method": "ping"},
        headers={"MCP-Protocol-Version": SUPPORTED_PROTOCOL_VERSION},
    )

    handler.do_POST()
    body = json.loads(handler.wfile.getvalue().decode("utf-8"))

    assert handler.status_code == 200
    assert body["result"] == {}


def test_missing_session_returns_session_error(tmp_path):
    handler = build_post_handler(
        tmp_path,
        {
            "jsonrpc": "2.0",
            "id": 8,
            "method": "tools/call",
            "params": {
                "name": "pptx_info",
                "arguments": {"session_id": "missing"},
            },
        },
        headers={"MCP-Protocol-Version": SUPPORTED_PROTOCOL_VERSION},
    )

    handler.do_POST()
    body = json.loads(handler.wfile.getvalue().decode("utf-8"))

    assert handler.status_code == 200
    assert body["error"]["code"] == -32003


def test_rejects_oversized_request_bodies(tmp_path):
    payload = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "tools/call",
        "params": {
            "name": "pptx_create",
            "arguments": {"name": "x" * MAX_REQUEST_SIZE},
        },
    }
    raw_body = json.dumps(payload).encode("utf-8")

    handler = HarnessHandler(
        tmp_path=tmp_path,
        path="/mcp",
        headers={
            "Authorization": "Bearer secret-token",
            "Content-Length": str(len(raw_body)),
        },
        body=raw_body,
    )

    handler.do_POST()

    assert handler.status_code == 413
    assert handler.error_message == "Request too large"


def test_stats_endpoint_requires_authentication(tmp_path):
    handler = HarnessHandler(tmp_path=tmp_path, path="/stats", headers={}, body=b"", token="secret-token")

    handler.do_GET()

    assert handler.status_code == 401
    assert handler.error_message == "Unauthorized"


def test_mismatched_protocol_header_is_rejected(tmp_path):
    handler = build_post_handler(
        tmp_path,
        {"jsonrpc": "2.0", "id": 9, "method": "ping"},
        headers={"MCP-Protocol-Version": "2025-03-26"},
    )

    handler.do_POST()

    assert handler.status_code == 400
    assert handler.error_message == "Unsupported MCP-Protocol-Version"
