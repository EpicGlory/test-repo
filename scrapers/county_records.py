"""Scraper for county assessor/permit records (config-driven county)."""

import logging
import re
import json
from typing import List, Dict, Optional
from datetime import datetime

from bs4 import BeautifulSoup

from models.property import Property
from scrapers.base import BaseScraper
from scrapers import location_utils as lu

logger = logging.getLogger(__name__)


class CountyRecordsScraper(BaseScraper):
    """Scrapes the configured county's assessor API and permit records.

    This is the most reliable source for:
    - Determining if a property is truly under construction vs completed
    - Getting accurate lot size, owner info, and parcel data
    - Building permit status and construction phase
    """

    def __init__(self, config: dict):
        super().__init__(config)
        county_cfg = config.get("county_records", {})
        self.county_name = county_cfg.get("name", "County")
        self.assessor_api = county_cfg.get("assessor_api", "")
        self.assessor_search_url = county_cfg.get("assessor_search_url", "")
        self.permit_viewer_url = county_cfg.get("permit_viewer", "")
        self.accela_url = county_cfg.get("accela_portal", "")
        self.target_zips = lu.get_target_zips(config)
        self.target_cities = lu.get_target_cities(config)

    def scrape(self) -> List[Property]:
        """Scrape county records for new construction properties.

        This scraper primarily enriches data from other scrapers,
        but can also discover properties through permit searches.
        """
        all_properties = []

        # Search for new residential construction permits
        logger.info(f"Searching {self.county_name} permit records...")
        permit_props = self._search_permits()
        all_properties.extend(permit_props)

        return all_properties

    def lookup_properties(self, addresses: List[str]) -> Dict[str, dict]:
        """Look up existing addresses in county records for enrichment.

        Args:
            addresses: List of property addresses to look up.

        Returns:
            Dict mapping normalized addresses to county data dicts.
        """
        from utils.dedup import normalize_address

        county_data = {}

        for address in addresses:
            try:
                data = self._lookup_address(address)
                if data:
                    norm_addr = normalize_address(address)
                    county_data[norm_addr] = data
                    logger.info(f"Found county data for: {address}")
            except Exception as e:
                logger.warning(f"Error looking up {address}: {e}")

        return county_data

    def _lookup_address(self, address: str) -> Optional[dict]:
        """Look up a single address via the configured county assessor API."""
        if not self.assessor_api:
            return self._lookup_address_alternative(address)
        try:
            # The Assessor API accepts address searches
            search_url = f"{self.assessor_api.rstrip('/')}/v1/search"
            params = {
                "address": address,
                "type": "address",
            }

            self._delay()
            response = self.session.get(
                search_url,
                headers=self._get_headers(),
                params=params,
                timeout=self.timeout,
            )

            if response.status_code != 200:
                # Try alternative endpoint format
                return self._lookup_address_alternative(address)

            data = response.json()
            if not data:
                return None

            # Parse the response
            results = data if isinstance(data, list) else data.get("results", [data])
            if not results:
                return None

            record = results[0] if isinstance(results, list) else results
            return self._parse_assessor_record(record)

        except Exception as e:
            logger.debug(f"Assessor API lookup failed for {address}: {e}")
            return self._lookup_address_alternative(address)

    def _lookup_address_alternative(self, address: str) -> Optional[dict]:
        """Try alternative lookup via the Assessor website search URL."""
        if not self.assessor_search_url:
            return None
        try:
            search_url = self.assessor_search_url
            params = {"q": address}

            self._delay()
            response = self.session.get(
                search_url,
                headers=self._get_headers(),
                params=params,
                timeout=self.timeout,
            )

            if response.status_code != 200:
                return None

            # Try to parse JSON response
            try:
                data = response.json()
                if isinstance(data, list) and data:
                    return self._parse_assessor_record(data[0])
            except json.JSONDecodeError:
                pass

            # Fall back to HTML parsing
            soup = BeautifulSoup(response.text, "lxml")
            return self._parse_assessor_html(soup)

        except Exception as e:
            logger.debug(f"Alternative assessor lookup failed: {e}")
            return None

    def _parse_assessor_record(self, record: dict) -> dict:
        """Parse an assessor API record into our enrichment format."""
        return {
            "parcel_number": str(record.get("parcel", "") or record.get("apn", "")),
            "owner": str(record.get("owner", "") or record.get("ownerName", "")),
            "year_built": str(record.get("yearBuilt", "") or record.get("year_built", "")),
            "sqft": str(record.get("livingArea", "") or record.get("sqft", "")),
            "lot_size": str(record.get("lotSize", "") or record.get("lot_size", "")),
            "property_type": str(record.get("propertyType", "") or record.get("use_code", "")),
            "assessed_value": str(record.get("fullCashValue", "") or record.get("fcv", "")),
            "legal_description": str(record.get("legal", "") or record.get("legalDescription", "")),
            "subdivision": str(record.get("subdivision", "")),
            "phase": "",  # Will be enriched from permit data
        }

    def _parse_assessor_html(self, soup: BeautifulSoup) -> Optional[dict]:
        """Parse assessor HTML response for property data."""
        text = soup.get_text(" ", strip=True)
        if not text or "no results" in text.lower():
            return None

        data = {
            "parcel_number": "",
            "owner": "",
            "year_built": "",
            "sqft": "",
            "lot_size": "",
            "property_type": "",
            "assessed_value": "",
            "legal_description": "",
            "subdivision": "",
            "phase": "",
        }

        # Try to extract fields from HTML
        for label_el in soup.find_all(["th", "dt", "label", "span"]):
            label = label_el.get_text(strip=True).lower()
            value_el = label_el.find_next_sibling(["td", "dd", "span", "div"])
            if not value_el:
                continue
            value = value_el.get_text(strip=True)

            if "parcel" in label or "apn" in label:
                data["parcel_number"] = value
            elif "owner" in label:
                data["owner"] = value
            elif "year built" in label:
                data["year_built"] = value
            elif "living area" in label or "sqft" in label:
                data["sqft"] = value
            elif "lot size" in label or "lot area" in label:
                data["lot_size"] = value
            elif "subdivision" in label:
                data["subdivision"] = value

        return data if any(data.values()) else None

    def _search_permits(self) -> List[Property]:
        """Search for new residential construction permits in target areas."""
        properties = []

        # Try the Permit Viewer
        if self.permit_viewer_url:
            try:
                permit_props = self._search_permit_viewer()
                properties.extend(permit_props)
            except Exception as e:
                logger.warning(f"Permit Viewer search failed: {e}")

        # Try Accela portal
        if self.accela_url:
            try:
                accela_props = self._search_accela()
                properties.extend(accela_props)
            except Exception as e:
                logger.warning(f"Accela search failed: {e}")

        return properties

    def _search_permit_viewer(self) -> List[Property]:
        """Search the county Permit Viewer for new construction permits."""
        properties = []

        # The permit viewer may require Selenium
        soup = self.fetch_page_selenium(self.permit_viewer_url)
        if not soup:
            soup = self.fetch_page(self.permit_viewer_url)
        if not soup:
            logger.warning("Could not access Permit Viewer")
            return []

        # Look for search forms and results
        # The permit viewer typically has a search interface
        text = soup.get_text(" ", strip=True)
        logger.info(f"Permit Viewer page loaded ({len(text)} chars)")

        # Try to find permit listings or search results
        # Note: The actual interaction may require form submission via Selenium
        permit_entries = soup.select(
            "tr[class*='permit'], div[class*='permit'], "
            "div[class*='result'], tr.data-row"
        )

        for entry in permit_entries:
            prop = self._parse_permit_entry(entry)
            if prop:
                properties.append(prop)

        return properties

    def _search_accela(self) -> List[Property]:
        """Search the Accela Citizen Access portal for building permits."""
        properties = []

        soup = self.fetch_page_selenium(self.accela_url)
        if not soup:
            logger.warning("Could not access Accela portal")
            return []

        # Accela portals typically have permit search functionality
        # This would require form interaction with Selenium for full search
        text = soup.get_text(" ", strip=True)
        logger.info(f"Accela portal loaded ({len(text)} chars)")

        return properties

    def _parse_permit_entry(self, entry) -> Optional[Property]:
        """Parse a permit record entry into a Property."""
        text = entry.get_text(" ", strip=True)
        text_lower = text.lower()

        # Only interested in new residential construction
        if not any(kw in text_lower for kw in ["new", "residential", "single family", "custom home"]):
            return None

        # Check if it's in our target area
        if not any(z in text for z in self.target_zips):
            # Also check city names
            if not any(city in text_lower for city in self.target_cities):
                return None

        address = self._extract_address(text)
        if not address:
            return None

        # Determine phase from permit status
        phase = self._determine_phase_from_permit(text)

        prop = Property(
            address=address,
            city=self._extract_city(text),
            phase=phase,
            details=f"Permit record: {text[:1000]}",
            source=f"{self.county_name} Permits",
            date_scraped=datetime.now().strftime("%Y-%m-%d"),
        )

        return prop

    def _determine_phase_from_permit(self, text: str) -> str:
        """Determine construction phase from permit status text."""
        text_lower = text.lower()

        phase_map = [
            (["approved", "issued", "active"], "Construction"),
            (["submitted", "in review", "pending", "plan review"], "Pre-Construction"),
            (["inspection", "final inspection"], "Construction"),
            (["foundation", "footing"], "Foundation"),
            (["framing", "rough"], "Framing"),
            (["grading", "site work"], "Pre-Construction"),
            (["certificate of occupancy", "co issued", "finaled"], "Completed"),
        ]

        for keywords, phase in phase_map:
            if any(kw in text_lower for kw in keywords):
                return phase

        return "Under Review"

    def _extract_address(self, text: str) -> str:
        """Extract address from permit text."""
        match = re.search(
            r"(\d{1,6}\s+(?:[NSEW]\.?\s+)?[\w\s]+(?:St|Street|Ave|Avenue|Blvd|Boulevard|Dr|Drive|"
            r"Ct|Court|Ln|Lane|Rd|Road|Way|Trail|Trl|Path|Pl|Place))",
            text,
            re.IGNORECASE,
        )
        return match.group(1).strip() if match else ""

    def _extract_city(self, text: str) -> str:
        return lu.detect_city(text, self.config)

    def enrich_with_permits(self, properties: List[Property]) -> List[Property]:
        """Enrich properties with permit data by looking up each address."""
        for prop in properties:
            if not prop.address:
                continue

            try:
                permit_data = self._lookup_permit(prop.address)
                if permit_data:
                    if permit_data.get("phase") and not prop.phase:
                        prop.phase = permit_data["phase"]
                    if permit_data.get("permit_number"):
                        existing = prop.additional_info or ""
                        prop.additional_info = (
                            f"{existing}; Permit: {permit_data['permit_number']}"
                            if existing
                            else f"Permit: {permit_data['permit_number']}"
                        )
                    logger.info(f"Enriched {prop.address} with permit data")
            except Exception as e:
                logger.debug(f"Permit lookup failed for {prop.address}: {e}")

        return properties

    def _lookup_permit(self, address: str) -> Optional[dict]:
        """Look up building permit for an address."""
        if not self.accela_url:
            return None
        # Try the Accela API-style lookup
        try:
            search_url = f"{self.accela_url.rstrip('/')}/Cap/CapHome.aspx"
            self._delay()
            response = self.session.get(
                search_url,
                headers=self._get_headers(),
                params={"SearchType": "address", "SearchValue": address},
                timeout=self.timeout,
            )

            if response.status_code == 200:
                soup = BeautifulSoup(response.text, "lxml")
                return self._parse_permit_result(soup)
        except Exception as e:
            logger.debug(f"Permit lookup error: {e}")

        return None

    def _parse_permit_result(self, soup: BeautifulSoup) -> Optional[dict]:
        """Parse permit search results."""
        text = soup.get_text(" ", strip=True)
        if "no record" in text.lower() or "no result" in text.lower():
            return None

        result = {
            "permit_number": "",
            "phase": "",
            "status": "",
        }

        # Look for permit number
        permit_match = re.search(r"(RES[-\d]+|BLD[-\d]+|PLN[-\d]+)", text)
        if permit_match:
            result["permit_number"] = permit_match.group(1)

        result["phase"] = self._determine_phase_from_permit(text)

        return result if any(result.values()) else None
