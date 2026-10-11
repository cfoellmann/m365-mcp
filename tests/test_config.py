import pytest

from m365_mcp.config import load_config
from m365_mcp.errors import ConfigError

VALID = """
tenant_id = "tid"
client_id = "cid"
account = "Olivia@Example.org"
internal_domains = ["Example.org"]

[external_mail]
reply_to = "chef@example.org"
cc = ["chef@example.org"]

[calendars]
christian = "christian@example.org"

[planner.plans]
familie = "plan1"

[planner.assignees]
christian = "uid-c"
"""


def write(tmp_path, text):
    p = tmp_path / "config.toml"
    p.write_text(text, encoding="utf-8")
    return p


def test_valid_config(tmp_path):
    cfg = load_config(write(tmp_path, VALID))
    assert cfg.internal_domains == ("example.org",)
    assert cfg.account == "olivia@example.org"
    assert cfg.calendars == {"christian": "christian@example.org"}
    assert cfg.plans == {"familie": "plan1"}
    assert cfg.assignees == {"christian": "uid-c"}
    assert cfg.timezone == "Europe/Berlin"
    assert cfg.max_body_chars == 8000
    assert cfg.token_cache == tmp_path / "token_cache.json"
    assert cfg.token_lock == tmp_path / "token_cache.lock"


def test_missing_file(tmp_path):
    with pytest.raises(ConfigError, match="nicht lesbar"):
        load_config(tmp_path / "fehlt.toml")


@pytest.mark.parametrize(
    "old,new,match",
    [
        ('tenant_id = "tid"', 'tenant_id = ""', "tenant_id"),
        ('internal_domains = ["Example.org"]', "internal_domains = []", "internal_domains"),
        ('reply_to = "chef@example.org"', 'reply_to = "chef@extern.example"', "intern"),
        ('cc = ["chef@example.org"]', 'cc = ["x@extern.example"]', "intern"),
        ('cc = ["chef@example.org"]', "cc = []", "nicht leere"),
        ('cc = ["chef@example.org"]', "", "nicht leere"),
        ('account = "Olivia@Example.org"', 'account = "olivia@extern.example"', "intern"),
        ('christian = "christian@example.org"', 'christian = "c@extern.example"', "intern"),
        ("[calendars]\nchristian", "[calendars]\nself", "self"),
        ("[planner.plans]\nfamilie", "[planner.plans]\nFamilie", "Schlüssel"),
        ('familie = "plan1"', 'familie = ""', "leer"),
    ],
)
def test_invalid_configs(tmp_path, old, new, match):
    assert old in VALID
    with pytest.raises(ConfigError, match=match):
        load_config(write(tmp_path, VALID.replace(old, new)))


def test_max_body_chars_must_be_positive(tmp_path):
    with pytest.raises(ConfigError, match="max_body_chars"):
        load_config(write(tmp_path, "max_body_chars = 0\n" + VALID))
