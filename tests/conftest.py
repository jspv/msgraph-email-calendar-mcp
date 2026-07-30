"""Shared test fixtures.

The one rule enforced here: tests never touch the network. A test in an earlier
batch reached the live Microsoft Graph API using the developer's cached token
and got back "Id is malformed" -- harmless that time, purely by luck. Tests that
need HTTP patch ``GraphClient``, which sits well above this layer, so the guard
only ever fires on the mistake.
"""
from __future__ import annotations

import httpx
import pytest


@pytest.fixture(autouse=True, scope="session")
def _block_network():
    """Make any real outbound HTTP request raise instead of leaving the machine.

    ``httpx.Client.request`` is the single funnel: the module-level helpers
    (``httpx.get`` and friends) construct a ``Client`` internally, and
    ``GraphClient`` calls ``client.request(...)`` directly, so patching this one
    method covers every path a test could take.
    """
    real_request = httpx.Client.request

    def _refuse(self, method, url, *args, **kwargs):
        raise RuntimeError(
            f"Test attempted a real network call: {method} {url}. "
            "Patch msgraph_mcp.graph.GraphClient (or httpx) in the test instead."
        )

    httpx.Client.request = _refuse
    try:
        yield
    finally:
        httpx.Client.request = real_request
