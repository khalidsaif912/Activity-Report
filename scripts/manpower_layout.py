"""Manpower section ordering and special roster placements."""

from __future__ import annotations

import re
from typing import Iterable

# Serial numbers → move into Inventory / Support Team from other sections when on shift.
# 990737 (Said Al Amri): listed under Inventory when roster shows him on that shift (not under Supervisor).
INVENTORY_SN_IDS: frozenset[str] = frozenset({"82592", "990737"})
SUPPORT_TEAM_SN_IDS: frozenset[str] = frozenset({"82565", "82653"})
INVENTORY_SN_ORDER: list[str] = ["82592", "990737"]
SUPPORT_SN_ORDER: list[str] = ["82565", "82653"]

PREFIX_TITLES: list[str] = [
    "Supervisor",
    "Load Control",
    "Export Checker",
    "Export Operators",
    "Flight Dispatch",
]

TAIL_TITLES: list[str] = [
    "CTU Staff On Duty",
    "Inventory",
    "Support Team",
    "Sick Leave / No Show / Others",
    "Annual Leave / Course / Off in Lieu",
    "Trainee",
    "Overtime Justification",
]

ALL_MANPOWER_TITLES: list[str] = PREFIX_TITLES + TAIL_TITLES


def normalize_sn_id(item: str) -> str | None:
    m = re.match(r"^SN(\d+)\s+", item.strip(), re.IGNORECASE)
    return m.group(1) if m else None


def _sn_sort_key(item: str, order: list[str]) -> tuple:
    sid = normalize_sn_id(item) or ""
    if sid in order:
        return (0, order.index(sid))
    return (1, item)


def reassign_special_roster_staff(roster: dict[str, list[str]]) -> None:
    """Move configured SNs from roster sections into Inventory / Support Team."""
    inv = roster.setdefault("Inventory", [])
    sup = roster.setdefault("Support Team", [])

    skip_scan = frozenset({"Inventory", "Support Team", "Flight Dispatch"})

    for title in list(roster.keys()):
        if title in skip_scan:
            continue
        items = roster.get(title) or []
        kept: list[str] = []
        for item in items:
            sid = normalize_sn_id(item)
            if sid in INVENTORY_SN_IDS:
                if item not in inv:
                    inv.append(item)
            elif sid in SUPPORT_TEAM_SN_IDS:
                if item not in sup:
                    sup.append(item)
            else:
                kept.append(item)
        roster[title] = kept

    inv.sort(key=lambda it: _sn_sort_key(it, INVENTORY_SN_ORDER))
    sup.sort(key=lambda it: _sn_sort_key(it, SUPPORT_SN_ORDER))


def finalize_manpower_sections(
    roster: dict[str, list[str]],
    flight_dispatch_items: Iterable[str],
) -> list[dict]:
    """
    Fixed top: Supervisor … Flight Dispatch (dispatch JSON for Flight Dispatch).
    Remaining sections sorted by headcount (desc), then TAIL_TITLES order on ties.
    """
    flight_list = list(flight_dispatch_items)

    r: dict[str, list[str]] = {}
    for t in ALL_MANPOWER_TITLES:
        r[t] = list(roster.get(t) or [])
    for title, items in roster.items():
        if title not in r and title != "Flight Dispatch":
            r[title] = list(items)

    reassign_special_roster_staff(r)

    result: list[dict] = []
    for t in PREFIX_TITLES:
        if t == "Flight Dispatch":
            items = flight_list
        else:
            items = list(r.get(t) or [])
        result.append({"title": t, "items": items})

    tail_blocks: list[dict] = []
    for t in TAIL_TITLES:
        tail_blocks.append({"title": t, "items": list(r.get(t) or [])})

    tail_blocks.sort(
        key=lambda b: (-len(b["items"]), TAIL_TITLES.index(b["title"]))
    )
    result.extend(tail_blocks)
    return result
