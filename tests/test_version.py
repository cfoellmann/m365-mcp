import tomllib
from pathlib import Path

import m365_mcp


def test_version_matches_pyproject():
    pyproject = tomllib.loads((Path(__file__).parents[1] / "pyproject.toml").read_text(encoding="utf-8"))
    assert m365_mcp.__version__ == pyproject["project"]["version"]
