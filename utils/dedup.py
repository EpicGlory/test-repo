"""Address normalization and property deduplication."""

import re
from typing import List

from models.property import Property


# Common address abbreviation mappings
ABBREVIATIONS = {
    r"\bstreet\b": "ST",
    r"\bstr\b": "ST",
    r"\bavenue\b": "AVE",
    r"\bav\b": "AVE",
    r"\bboulevard\b": "BLVD",
    r"\bdrive\b": "DR",
    r"\bcourt\b": "CT",
    r"\bcircle\b": "CIR",
    r"\blane\b": "LN",
    r"\broad\b": "RD",
    r"\bplace\b": "PL",
    r"\bway\b": "WAY",
    r"\btrail\b": "TRL",
    r"\bpath\b": "PATH",
    r"\bnorth\b": "N",
    r"\bsouth\b": "S",
    r"\beast\b": "E",
    r"\bwest\b": "W",
    r"\bnortheast\b": "NE",
    r"\bnorthwest\b": "NW",
    r"\bsoutheast\b": "SE",
    r"\bsouthwest\b": "SW",
    r"\bapartment\b": "APT",
    r"\bsuite\b": "STE",
    r"\bunit\b": "UNIT",
}


def normalize_address(address: str) -> str:
    """Normalize an address for comparison purposes."""
    if not address:
        return ""

    addr = address.upper().strip()

    # Remove punctuation except hyphens in house numbers
    addr = re.sub(r"[.,#]", "", addr)

    # Standardize abbreviations
    for pattern, replacement in ABBREVIATIONS.items():
        addr = re.sub(pattern, replacement, addr, flags=re.IGNORECASE)

    # Collapse multiple spaces
    addr = re.sub(r"\s+", " ", addr).strip()

    return addr


def _address_similarity(addr1: str, addr2: str) -> float:
    """Calculate similarity between two normalized addresses."""
    if not addr1 or not addr2:
        return 0.0

    if addr1 == addr2:
        return 1.0

    # Extract house number and street name for comparison
    num1 = re.match(r"^(\d+)", addr1)
    num2 = re.match(r"^(\d+)", addr2)

    # If house numbers differ, not the same property
    if num1 and num2 and num1.group(1) != num2.group(1):
        return 0.0

    # Simple token-based similarity
    tokens1 = set(addr1.split())
    tokens2 = set(addr2.split())

    if not tokens1 or not tokens2:
        return 0.0

    intersection = tokens1 & tokens2
    union = tokens1 | tokens2

    return len(intersection) / len(union)


# Source priority for merging (higher = more trusted)
SOURCE_PRIORITY = {
    "County": 4,  # county assessor/permits (any county)
    "Builder": 3,
    "Luxury Realty": 2,
    "Listing": 1,
}


def _get_source_priority(source: str) -> int:
    """Get the priority level for a source."""
    for key, priority in SOURCE_PRIORITY.items():
        if key.lower() in source.lower():
            return priority
    return 0


def deduplicate_properties(properties: List[Property], similarity_threshold: float = 0.75) -> List[Property]:
    """Deduplicate properties by normalized address.

    When duplicates are found, merges data preferring higher-priority sources.
    """
    if not properties:
        return []

    # Group by normalized address
    address_groups: dict[str, List[Property]] = {}

    for prop in properties:
        norm_addr = normalize_address(prop.address)
        if not norm_addr:
            # Keep properties without addresses as-is
            address_groups[f"__no_addr_{id(prop)}"] = [prop]
            continue

        matched = False
        for existing_addr in list(address_groups.keys()):
            if existing_addr.startswith("__no_addr_"):
                continue
            if _address_similarity(norm_addr, existing_addr) >= similarity_threshold:
                address_groups[existing_addr].append(prop)
                matched = True
                break

        if not matched:
            address_groups[norm_addr] = [prop]

    # Merge each group into a single property
    merged = []
    for group in address_groups.values():
        if len(group) == 1:
            merged.append(group[0])
            continue

        # Sort by source priority (highest first)
        group.sort(key=lambda p: _get_source_priority(p.source), reverse=True)

        # Start with the highest priority property
        base = group[0]
        for other in group[1:]:
            base.merge_with(other)

        merged.append(base)

    return merged
