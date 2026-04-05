"""Cross-source data enrichment for properties."""

import logging
from typing import List, Dict

from models.property import Property

logger = logging.getLogger(__name__)


def enrich_properties(
    properties: List[Property],
    county_data: Dict[str, dict] = None,
) -> List[Property]:
    """Enrich property data from multiple sources.

    Args:
        properties: List of properties to enrich.
        county_data: Dict mapping normalized addresses to county record data.

    Returns:
        Enriched list of properties.
    """
    if county_data:
        _enrich_from_county(properties, county_data)

    _infer_community(properties)
    _apply_door_details(properties)
    _filter_completed(properties)

    return properties


def _enrich_from_county(properties: List[Property], county_data: Dict[str, dict]):
    """Enrich properties with Maricopa County records data."""
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


# Known community zip code / area mappings
COMMUNITY_ZIP_MAP = {
    "85255": ["Silverleaf", "DC Ranch", "Desert Mountain", "Troon North", "Estancia"],
    "85262": ["Desert Mountain", "Pinnacle Peak Estates"],
    "85266": ["Whisper Rock", "Desert Mountain"],
    "85259": ["McDowell Mountain Ranch", "DC Ranch"],
    "85260": ["McDowell Mountain Ranch"],
    "85253": ["Paradise Valley"],
}


def _infer_community(properties: List[Property]):
    """Try to infer community from address/zip if not already set."""
    for prop in properties:
        if prop.community:
            continue

        # Check if community name appears in details or address
        details_lower = (prop.details or "").lower() + " " + (prop.address or "").lower()
        community_names = [
            "Silverleaf", "Desert Mountain", "DC Ranch", "Estancia",
            "Whisper Rock", "Pinnacle Peak", "McDowell Mountain Ranch",
            "Troon North", "Paradise Valley", "Arcadia",
            "Mummy Mountain", "Camelback Mountain",
        ]
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
