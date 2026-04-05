"""Scraper for luxury realty agent websites with rich property data."""

import logging
import re
from typing import List, Optional
from datetime import datetime

from bs4 import BeautifulSoup

from models.property import Property
from scrapers.base import BaseScraper

logger = logging.getLogger(__name__)


class LuxuryRealtyScraper(BaseScraper):
    """Scrapes luxury real estate agent/brokerage websites.

    These sites often have the richest data including architect, designer,
    and builder information.
    """

    def __init__(self, config: dict):
        super().__init__(config)
        self.realty_sites = config.get("luxury_realty_sites", [])
        self.min_price = config.get("search", {}).get("min_price", 5_000_000)

    def scrape(self) -> List[Property]:
        """Scrape all configured luxury realty sites."""
        all_properties = []

        for site in self.realty_sites:
            try:
                name = site["name"]
                logger.info(f"Scraping luxury realty site: {name}")
                props = self._scrape_site(site)
                all_properties.extend(props)
                logger.info(f"Found {len(props)} properties from {name}")
            except Exception as e:
                logger.error(f"Error scraping {site['name']}: {e}")

        return all_properties

    def _scrape_site(self, site: dict) -> List[Property]:
        """Scrape a single luxury realty site."""
        properties = []
        base_url = site["url"].rstrip("/")

        for path in site.get("search_paths", []):
            url = f"{base_url}{path}"
            logger.info(f"Scraping: {url}")

            # Try requests first, fall back to Selenium
            soup = self.fetch_page(url)
            if not soup or not self._has_listing_content(soup):
                soup = self.fetch_page_selenium(url)

            if not soup:
                continue

            # Extract listings from the page
            props = self._extract_listings(soup, site, url)
            properties.extend(props)

            # Find and follow detail page links
            detail_links = self._find_listing_links(soup, base_url)
            for detail_url in detail_links[:15]:  # Limit per path
                try:
                    detail_soup = self.fetch_page(detail_url)
                    if not detail_soup:
                        detail_soup = self.fetch_page_selenium(detail_url)
                    if detail_soup:
                        detail_props = self._extract_detail_page(
                            detail_soup, site, detail_url
                        )
                        properties.extend(detail_props)
                except Exception as e:
                    logger.warning(f"Error on detail page {detail_url}: {e}")

        return properties

    def _has_listing_content(self, soup: BeautifulSoup) -> bool:
        """Check if the page has listing-type content."""
        text = soup.get_text(strip=True)
        listing_signals = ["bed", "bath", "sqft", "sq ft", "price", "$", "MLS"]
        signal_count = sum(1 for s in listing_signals if s.lower() in text.lower())
        return signal_count >= 2

    def _find_listing_links(self, soup: BeautifulSoup, base_url: str) -> List[str]:
        """Find links to individual property detail pages."""
        links = []
        listing_keywords = [
            "property", "listing", "home", "estate", "detail",
            "mls", "address", "for-sale",
        ]

        for a_tag in soup.find_all("a", href=True):
            href = a_tag["href"]
            href_lower = href.lower()

            is_listing_link = any(kw in href_lower for kw in listing_keywords)
            has_address_pattern = bool(re.search(r"\d{3,6}", href))

            if is_listing_link or has_address_pattern:
                full_url = href if href.startswith("http") else f"{base_url}{href}"
                # Only follow links on the same domain
                if base_url.split("//")[1].split("/")[0] in full_url and full_url not in links:
                    links.append(full_url)

        return links

    def _extract_listings(
        self, soup: BeautifulSoup, site: dict, url: str
    ) -> List[Property]:
        """Extract property listings from a search results page."""
        properties = []

        # Look for listing cards with various selectors
        card_selectors = [
            "div[class*='listing']", "div[class*='property']",
            "div[class*='home-card']", "div[class*='result']",
            "article[class*='listing']", "article[class*='property']",
            ".property-card", ".listing-card", ".home-card",
            "div[class*='HomeCard']", "div[class*='PropCard']",
            "li[class*='listing']", "li[class*='property']",
        ]

        cards = []
        for selector in card_selectors:
            found = soup.select(selector)
            if found:
                cards = found
                break

        if not cards:
            # Try a more generic approach - look for repeating structures
            cards = soup.select("div.card, article, .grid-item, .col")

        for card in cards:
            prop = self._parse_listing_card(card, site, url)
            if prop:
                properties.append(prop)

        return properties

    def _parse_listing_card(
        self, card, site: dict, url: str
    ) -> Optional[Property]:
        """Parse a listing card into a Property."""
        text = card.get_text(" ", strip=True)
        if len(text) < 20:
            return None

        # Extract address
        address = ""
        addr_selectors = [
            "address", "[class*='address']", "[class*='addr']",
            "h2", "h3", ".title", "[class*='title']",
        ]
        for sel in addr_selectors:
            addr_el = card.select_one(sel)
            if addr_el:
                addr_text = addr_el.get_text(strip=True)
                if re.search(r"\d{1,6}\s+\w", addr_text):
                    address = addr_text
                    break

        if not address:
            address = self._extract_address(text)

        # Extract price
        price = self._extract_price(text)
        if price:
            price_num = self._parse_price_number(price)
            if price_num and price_num < self.min_price:
                return None

        # Check for new construction indicators
        text_lower = text.lower()
        is_new_construction = any(
            kw in text_lower
            for kw in [
                "new construction", "under construction", "to be built",
                "pre-construction", "being built", "building",
                "proposed", "planned", "custom build",
            ]
        )

        # For luxury sites, we'll include listings even without explicit
        # "new construction" tags since we'll filter later with county data
        phase = self._extract_phase(text)

        # Get detail link
        detail_link = card.find("a", href=True)
        listing_url = detail_link["href"] if detail_link else url
        if listing_url and not listing_url.startswith("http"):
            listing_url = f"{site['url'].rstrip('/')}{listing_url}"

        # Extract builder/architect if mentioned
        builder = self._extract_builder(text)
        architect = self._extract_architect(text)

        prop = Property(
            address=address,
            city=self._extract_city(text),
            price=price,
            builder=builder,
            architect=architect,
            sqft=self._extract_sqft(text),
            bedrooms=self._extract_bedrooms(text),
            bathrooms=self._extract_bathrooms(text),
            lot_size=self._extract_lot_size(text),
            phase=phase if phase else ("New Construction" if is_new_construction else ""),
            community=self._extract_community(text),
            style=self._extract_style(text),
            details=text[:2000],
            listing_url=listing_url,
            source=f"Luxury Realty - {site['name']}",
            date_scraped=datetime.now().strftime("%Y-%m-%d"),
        )

        # Only return if we have at least an address or it's clearly new construction
        if address or is_new_construction:
            return prop
        return None

    def _extract_detail_page(
        self, soup: BeautifulSoup, site: dict, url: str
    ) -> List[Property]:
        """Extract rich property data from a detail page.

        Detail pages on luxury realty sites often have architect, designer,
        builder, and detailed descriptions.
        """
        text = soup.get_text(" ", strip=True)
        text_lower = text.lower()

        # Check if this is a Scottsdale/PV property
        if not any(
            loc in text_lower
            for loc in ["scottsdale", "paradise valley", "85255", "85253", "85259", "85262", "85266"]
        ):
            return []

        # Check for new construction signals
        construction_signals = [
            "new construction", "under construction", "to be built",
            "pre-construction", "being built", "custom build",
            "not yet built", "construction", "breaking ground",
        ]
        is_new_construction = any(sig in text_lower for sig in construction_signals)

        # Only include if there's a new construction signal
        # (detail pages without this are likely existing homes for sale)
        if not is_new_construction:
            return []

        # Extract all available data
        address = self._extract_address(text)
        price = self._extract_price(text)

        if price:
            price_num = self._parse_price_number(price)
            if price_num and price_num < self.min_price:
                return []

        # Luxury detail pages often have rich metadata
        prop = Property(
            address=address,
            city=self._extract_city(text),
            price=price,
            architect=self._extract_architect(text),
            designer=self._extract_designer(text),
            builder=self._extract_builder(text),
            community=self._extract_community(text),
            community_subdivision=self._extract_subdivision(text),
            phase=self._extract_phase(text) or "New Construction",
            sqft=self._extract_sqft(text),
            bedrooms=self._extract_bedrooms(text),
            bathrooms=self._extract_bathrooms(text),
            lot_size=self._extract_lot_size(text),
            stories=self._extract_stories(text),
            style=self._extract_style(text),
            details=text[:3000],  # Keep more detail from luxury sites
            listing_url=url,
            source=f"Luxury Realty - {site['name']}",
            date_scraped=datetime.now().strftime("%Y-%m-%d"),
        )

        return [prop]

    # --- Extraction helpers ---

    def _extract_address(self, text: str) -> str:
        patterns = [
            r"(\d{1,6}\s+(?:[NSEW]\.?\s+)?[\w\s]+(?:St|Street|Ave|Avenue|Blvd|Boulevard|Dr|Drive|"
            r"Ct|Court|Ln|Lane|Rd|Road|Way|Trail|Trl|Path|Pl|Place))\s*,?\s*"
            r"(?:Scottsdale|Paradise Valley)",
            r"(\d{1,6}\s+(?:[NSEW]\.?\s+)?[\w\s]+(?:St|Street|Ave|Avenue|Blvd|Boulevard|Dr|Drive|"
            r"Ct|Court|Ln|Lane|Rd|Road|Way|Trail|Trl|Path|Pl|Place))",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        return ""

    def _extract_price(self, text: str) -> str:
        patterns = [
            r"\$\s*([\d,]+(?:\.\d{2})?)\s*(?:million|M)\b",
            r"\$\s*([\d]{1,3}(?:,\d{3})*(?:,\d{3})?)(?:\+)?",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(0).strip()
        return ""

    def _parse_price_number(self, price_str: str) -> Optional[float]:
        if not price_str:
            return None
        million_match = re.search(r"([\d.]+)\s*(?:million|M)", price_str, re.IGNORECASE)
        if million_match:
            return float(million_match.group(1)) * 1_000_000
        cleaned = re.sub(r"[$,\s+]", "", price_str)
        try:
            return float(cleaned)
        except ValueError:
            return None

    def _extract_sqft(self, text: str) -> str:
        match = re.search(r"([\d,]+)\s*(?:sq\.?\s*ft|square\s*feet|SF)\b", text, re.IGNORECASE)
        return match.group(1) if match else ""

    def _extract_bedrooms(self, text: str) -> str:
        match = re.search(r"(\d+)\s*(?:bed(?:room)?s?|BR|bd)\b", text, re.IGNORECASE)
        return match.group(1) if match else ""

    def _extract_bathrooms(self, text: str) -> str:
        match = re.search(r"([\d.]+)\s*(?:bath(?:room)?s?|BA|ba)\b", text, re.IGNORECASE)
        return match.group(1) if match else ""

    def _extract_lot_size(self, text: str) -> str:
        match = re.search(r"([\d,.]+)\s*(?:acre|ac)\b", text, re.IGNORECASE)
        return f"{match.group(1)} acres" if match else ""

    def _extract_stories(self, text: str) -> str:
        match = re.search(r"(\d+)\s*(?:stor(?:y|ies)|level|floor)\b", text, re.IGNORECASE)
        return match.group(1) if match else ""

    def _extract_phase(self, text: str) -> str:
        text_lower = text.lower()
        phases = [
            ("design phase", "Design"),
            ("in design", "Design"),
            ("pre-construction", "Pre-Construction"),
            ("pre construction", "Pre-Construction"),
            ("under construction", "Construction"),
            ("being built", "Construction"),
            ("framing", "Framing"),
            ("foundation", "Foundation"),
            ("new construction", "New Construction"),
            ("to be built", "Pre-Construction"),
            ("coming soon", "Pre-Construction"),
        ]
        for keyword, phase in phases:
            if keyword in text_lower:
                return phase
        return ""

    def _extract_community(self, text: str) -> str:
        text_lower = text.lower()
        community_names = [
            "Silverleaf", "Desert Mountain", "DC Ranch", "Estancia",
            "Whisper Rock", "Pinnacle Peak Estates", "Pinnacle Peak",
            "McDowell Mountain Ranch", "Troon North",
            "Paradise Valley", "Arcadia", "Mummy Mountain", "Camelback Mountain",
        ]
        for name in community_names:
            if name.lower() in text_lower:
                return name
        return ""

    def _extract_subdivision(self, text: str) -> str:
        """Extract subdivision name."""
        text_lower = text.lower()
        subdivisions = [
            "Village", "Upper Canyon", "Horseshoe Canyon", "Lower Canyon",
            "The Summit", "ICON",
        ]
        for sub in subdivisions:
            if sub.lower() in text_lower:
                return sub
        return ""

    def _extract_city(self, text: str) -> str:
        text_lower = text.lower()
        if "paradise valley" in text_lower:
            return "Paradise Valley"
        if "scottsdale" in text_lower:
            return "Scottsdale"
        return ""

    def _extract_builder(self, text: str) -> str:
        """Extract builder name from text."""
        # Known builders
        known_builders = [
            "Cullum Homes", "Calvis Wyant", "Fratantoni", "Camelot Homes",
            "Desert Star Construction", "AFT Construction", "BedBrock",
            "Fenn Rogers", "Salcito", "Starwood Custom", "Peak One",
            "Integrity Luxury", "Sonora West", "Toll Brothers",
        ]
        text_lower = text.lower()
        for builder in known_builders:
            if builder.lower() in text_lower:
                return builder

        # Try pattern matching
        patterns = [
            r"(?:built|builder|constructed)\s+by\s+([A-Z][\w\s&]+?)(?:\s*[,.|])",
            r"builder[:\s]+([A-Z][\w\s&]+?)(?:\s*[,.|])",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        return ""

    def _extract_architect(self, text: str) -> str:
        patterns = [
            r"architect[:\s]+([A-Z][\w\s&]+?)(?:\s*[,.|]|\s+design|\s+build)",
            r"designed\s+by\s+([A-Z][\w\s&]+?)(?:\s*[,.|])",
            r"architecture\s+by\s+([A-Z][\w\s&]+?)(?:\s*[,.|])",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        return ""

    def _extract_designer(self, text: str) -> str:
        patterns = [
            r"(?:interior\s+)?designer[:\s]+([A-Z][\w\s&]+?)(?:\s*[,.|])",
            r"interior(?:s)?\s+by\s+([A-Z][\w\s&]+?)(?:\s*[,.|])",
            r"(?:interior\s+)?design\s+by\s+([A-Z][\w\s&]+?)(?:\s*[,.|])",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        return ""

    def _extract_style(self, text: str) -> str:
        styles = [
            "Contemporary", "Modern", "Desert Modern", "Desert Contemporary",
            "Mediterranean", "Spanish Colonial", "Tuscan", "Santa Fe",
            "Transitional", "Minimalist", "Mid-Century Modern",
            "Southwestern", "Adobe", "Hacienda",
        ]
        text_lower = text.lower()
        for style in styles:
            if style.lower() in text_lower:
                return style
        return ""
