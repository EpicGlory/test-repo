"""Cross-source data enrichment for properties."""

import logging
from typing import List, Dict

from models.property import Property

logger = logging.getLogger(__name__)


def enrich_properties(
    properties: List[Property],
    county_data: Dict[str, dict] = None,
    config: dict = None,
) -> List[Property]:
    """Enrich property data from multiple sources.

    Args:
        properties: List of properties to enrich.
        county_data: Dict mapping normalized addresses to county record data.
        config: Optional config for community/zip lookup tables.

    Returns:
        Enriched list of properties.
    """
    if county_data:
        _enrich_from_county(properties, county_data)

    _infer_community(properties, config or {})
    _apply_door_details(properties)
    _filter_completed(properties)

    return properties


def _enrich_from_county(properties: List[Property], county_data: Dict[str, dict]):
    """Enrich properties with county records data."""
    from utils.dedup import normalize_address

    for prop in properties:
        norm_addr = normalize_address(prop.address)
        if norm_addr in county_data:
            record = county_data[norm_addr]
            # County data is highest priority for these fields
            if record.get("phase") and not prop.phase:
                prop.phase = record["phase"]
            if record.get("year_built"):
                prop.year_built = record["year_built"]
            if record.get("sqft") and not prop.sqft:
                prop.sqft = record["sqft"]
            if record.get("lot_size") and not prop.lot_size:
                prop.lot_size = record["lot_size"]
            if record.get("owner") and not prop.additional_info:
                prop.additional_info = f"Owner: {record['owner']}"
            logger.info(f"Enriched {prop.address} with county data")


def _infer_community(properties: List[Property], config: dict):
    """Try to infer community from text mentions if not already set."""
    community_names = []
    for region in config.get("communities", {}).values():
        for comm in region:
            community_names.append(comm.get("name", ""))
    # Sort longest-first for multi-word preference
    community_names = sorted([c for c in community_names if c], key=len, reverse=True)

    for prop in properties:
        if prop.community:
            continue
        details_lower = (prop.details or "").lower() + " " + (prop.address or "").lower()
        for name in community_names:
            if name.lower() in details_lower:
                prop.community = name
                break


def _apply_door_details(properties: List[Property]):
    """Apply door_details logic to all properties."""
    for prop in properties:
        prop.apply_door_details()


def _filter_completed(properties: List[Property]) -> List[Property]:
    """Remove properties that appear to be completed homes.

    Modifies the list in-place and returns it.
    """
    to_remove = []
    for i, prop in enumerate(properties):
        if prop.is_complete_home():
            logger.info(f"Filtering out completed home: {prop.address}")
            to_remove.append(i)

    for i in reversed(to_remove):
        properties.pop(i)

    return properties
