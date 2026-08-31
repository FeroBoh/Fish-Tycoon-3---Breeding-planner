"""
inventory.py
------------
Saves/loads the list of fish the user currently owns in their aquarium,
kept separately per map. Stored in a plain JSON file next to the script.
"""

from __future__ import annotations

import json
import os
from typing import Dict, List, Tuple

Fish = Tuple[str, str]

DEFAULT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "aquarium_inventory.json")


def load_inventory(path: str = DEFAULT_PATH) -> Dict[str, List[Fish]]:
    if not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as f:
        raw = json.load(f)
    result: Dict[str, List[Fish]] = {}
    for map_name, fish_list in raw.items():
        result[map_name] = [tuple(item) for item in fish_list]
    return result


def save_inventory(data: Dict[str, List[Fish]], path: str = DEFAULT_PATH) -> None:
    serializable = {map_name: [list(f) for f in fish_list] for map_name, fish_list in data.items()}
    with open(path, "w", encoding="utf-8") as f:
        json.dump(serializable, f, ensure_ascii=False, indent=2)
