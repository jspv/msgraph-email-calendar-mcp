"""The test suite must never make a real outbound HTTP request.

A test in an earlier batch reached the live Microsoft Graph API using the
developer's cached token. ``AGENTS.md`` now forbids network access in tests;
this file proves the prohibition is enforced rather than merely written down.

The URLs below point at a closed local port, so this test makes no real network
request whether or not the guard is installed.
"""
from __future__ import annotations

import httpx
import pytest

DEAD_URL = "http://127.0.0.1:1/"


class TestNetworkGuard:
    def test_direct_client_request_is_blocked(self):
        with pytest.raises(RuntimeError, match="real network call"):
            httpx.Client(timeout=0.01).get(DEAD_URL)

    def test_module_level_helper_is_blocked(self):
        # httpx.get() builds its own Client internally, so the guard has to sit
        # on the Client method rather than on the module function.
        with pytest.raises(RuntimeError, match="real network call"):
            httpx.get(DEAD_URL, timeout=0.01)

    def test_error_names_the_offending_url(self):
        with pytest.raises(RuntimeError, match="127.0.0.1"):
            httpx.Client(timeout=0.01).get(DEAD_URL)
