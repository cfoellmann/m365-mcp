from dataclasses import dataclass

from m365_mcp.config import Config
from m365_mcp.graph import GraphClient


@dataclass(frozen=True)
class Ctx:
    cfg: Config
    graph: GraphClient
