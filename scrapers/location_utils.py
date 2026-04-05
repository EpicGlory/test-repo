"""Location-awareness helpers driven by config.

Centralizes target city/zip/community/keyword logic so the scrapers can be
reused across markets (Scottsdale, Summit County, etc.) by swapping config.
"""

from typing import List


def get_target_cities(config: dict) -> List[str]:
    """List of target city names (lower-cased) from config.search.locations."""
    cities = []
    for loc in config.get("search", {}).get("locations", []):
        name = loc.get("name", "")
        if name:
            cities.append(name.lower())
    return cities


def get_target_zips(config: dict) -> List[str]:
    """List of target zip codes from config.search.locations."""
    zips = []
    for loc in config.get("search", {}).get("locations", []):
        zips.extend(loc.get("zip_codes", []))
    return zips


def get_target_state(config: dict) -> str:
    """Primary state code from config."""
    for loc in config.get("search", {}).get("locations", []):
        state = loc.get("state")
        if state:
            return state
    return ""


def get_target_communities(config: dict) -> List[str]:
    """List of community names from config.communities (original casing)."""
    names = []
    for region in config.get("communities", {}).values():
        for comm in region:
            names.append(comm["name"])
    return names


def get_target_subdivisions(config: dict) -> List[str]:
    """List of subdivision names from config.communities."""
    subs = []
    for region in config.get("communities", {}).values():
        for comm in region:
            for sub in comm.get("subdivisions", []) or []:
                subs.append(sub)
    return subs


def get_location_keywords(config: dict) -> List[str]:
    """All lower-cased tokens (cities + zips + communities + subdivisions)."""
    kws = set()
    kws.update(get_target_cities(config))
    kws.update(z.lower() for z in get_target_zips(config))
    kws.update(c.lower() for c in get_target_communities(config))
    kws.update(s.lower() for s in get_target_subdivisions(config))
    return [k for k in kws if k]


def get_city_display_names(config: dict) -> List[str]:
    """Target city names with original casing for display/matching."""
    return [loc.get("name", "") for loc in config.get("search", {}).get("locations", []) if loc.get("name")]


def detect_city(text: str, config: dict) -> str:
    """Return the target city mentioned in text, or empty string.

    Matches longest names first so multi-word cities (e.g. 'Paradise Valley',
    'Park City') match before shorter substrings.
    """
    text_lower = (text or "").lower()
    cities = sorted(get_city_display_names(config), key=len, reverse=True)
    for city in cities:
        if city.lower() in text_lower:
            return city
    return ""


def detect_community(text: str, config: dict) -> str:
    """Return the community mentioned in text, or empty string."""
    text_lower = (text or "").lower()
    names = sorted(get_target_communities(config), key=len, reverse=True)
    for name in names:
        if name.lower() in text_lower:
            return name
    return ""


def detect_subdivision(text: str, config: dict) -> str:
    """Return the subdivision mentioned in text, or empty string."""
    text_lower = (text or "").lower()
    subs = sorted(get_target_subdivisions(config), key=len, reverse=True)
    for sub in subs:
        if sub.lower() in text_lower:
            return sub
    return ""


def text_mentions_target(text: str, config: dict) -> bool:
    """Check whether text mentions any target city/zip/community."""
    text_lower = (text or "").lower()
    return any(kw in text_lower for kw in get_location_keywords(config))


def build_city_regex_alt(config: dict) -> str:
    """Regex alternation of target city display names (escaped)."""
    import re as _re
    cities = get_city_display_names(config)
    if not cities:
        return r"[A-Za-z ]+"
    # Sort longest first to prefer multi-word matches
    cities_sorted = sorted(cities, key=len, reverse=True)
    return "|".join(_re.escape(c) for c in cities_sorted)
