from m365_mcp.errors import GraphError, M365Error, PolicyError, ReauthRequired


def test_graph_error_carries_status_and_code():
    err = GraphError(403, "ErrorAccessDenied", "Access is denied.")
    assert err.status == 403
    assert err.code == "ErrorAccessDenied"
    assert "403" in str(err) and "ErrorAccessDenied" in str(err)
    assert isinstance(err, M365Error)


def test_reauth_message_mentions_relogin():
    assert "Neuanmeldung" in str(ReauthRequired())
    assert "invalid_grant" in str(ReauthRequired("invalid_grant"))


def test_policy_error_is_m365_error():
    assert isinstance(PolicyError("x"), M365Error)


def test_graph_error_str_omits_raw_message():
    err = GraphError(403, "ErrorAccessDenied", "Zugriff für geheim@example.org verweigert")
    assert str(err) == "Graph 403 ErrorAccessDenied"
    assert "geheim@example.org" not in str(err)
    assert err.message == "Zugriff für geheim@example.org verweigert"


def test_transport_graph_error_text_is_fixed():
    err = GraphError(0, "transport", "https://graph.microsoft.com/secret?x=1")
    assert "Ergebnis unklar" in str(err)
    assert "graph.microsoft.com" not in str(err)
