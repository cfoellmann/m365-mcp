"""Befehle: login (einmalig, interaktiv), serve (stdio, von Hermes gestartet), check (Health)."""

import argparse
import logging
import sys
from pathlib import Path

from m365_mcp.auth import TokenProvider, ensure_owner, login
from m365_mcp.config import Config, load_config
from m365_mcp.context import Ctx
from m365_mcp.errors import M365Error, ReauthRequired
from m365_mcp.graph import GraphClient


def check(cfg: Config) -> str:
    me = GraphClient(TokenProvider(cfg)).get("/me", params={"$select": "userPrincipalName"})
    upn = me.get("userPrincipalName", "")
    if upn.lower() != cfg.account.lower():
        raise M365Error(f"Token gehört {upn}, erwartet {cfg.account}")
    return upn


def serve(cfg: Config) -> None:
    from m365_mcp.server import build_server

    build_server(Ctx(cfg=cfg, graph=GraphClient(TokenProvider(cfg)))).run("stdio")


def configure_logging() -> None:
    """Nur nach stderr. Fremde Logger leise: httpx loggt URLs (/users/<adresse>), das MCP-SDK
    Werkzeugfehler samt Text (Adressen aus Richtlinienmeldungen)."""
    logging.basicConfig(stream=sys.stderr, level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    for noisy in ("httpx", "httpcore", "msal", "mcp"):
        logging.getLogger(noisy).setLevel(logging.WARNING)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="m365-mcp")
    sub = parser.add_subparsers(dest="command", required=True)
    for name in ("login", "serve", "check"):
        sub.add_parser(name).add_argument("--config", required=True, type=Path)
    args = parser.parse_args(argv)

    configure_logging()

    try:
        cfg = load_config(args.config)
        ensure_owner(cfg.directory)
        if args.command == "login":
            print(f"Angemeldet als {login(cfg)}")
        elif args.command == "check":
            print(f"OK {check(cfg)}")
        else:
            serve(cfg)
    except ReauthRequired as exc:
        print(f"{exc}: m365-mcp login --config {args.config}", file=sys.stderr)
        return 2
    except M365Error as exc:
        print(str(exc), file=sys.stderr)
        return 1
    return 0
