import pytest

from m365_mcp import cli
from m365_mcp.errors import ReauthRequired

CONFIG = """
tenant_id = "tid"
client_id = "cid"
account = "olivia@example.org"
internal_domains = ["example.org"]
[external_mail]
reply_to = "chef@example.org"
cc = ["chef@example.org"]
"""


@pytest.fixture
def config_path(tmp_path):
    p = tmp_path / "config.toml"
    p.write_text(CONFIG, encoding="utf-8")
    return p


def test_check_ok(config_path, monkeypatch, capsys):
    monkeypatch.setattr(cli, "check", lambda cfg: "olivia@example.org")
    assert cli.main(["check", "--config", str(config_path)]) == 0
    assert capsys.readouterr().out.strip() == "OK olivia@example.org"


def test_check_reauth_exit_3(config_path, monkeypatch, capsys):
    def boom(cfg):
        raise ReauthRequired("invalid_grant")

    monkeypatch.setattr(cli, "check", boom)
    assert cli.main(["check", "--config", str(config_path)]) == 3
    assert f"m365-mcp login --config {config_path}" in capsys.readouterr().err


def test_usage_error_is_not_reauth_code():
    with pytest.raises(SystemExit) as exc:
        cli.main(["check"])
    assert exc.value.code == 2


def test_bad_config_exit_1(tmp_path, capsys):
    assert cli.main(["check", "--config", str(tmp_path / "fehlt.toml")]) == 1
    assert "nicht lesbar" in capsys.readouterr().err
