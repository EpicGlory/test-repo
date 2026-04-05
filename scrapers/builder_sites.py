"""Scraper for luxury home builder websites (config-driven target market)."""

import logging
import re
from typing import List, Optional
from datetime import datetime

from bs4 import BeautifulSoup, Tag

from models.property import Property
from scrapers.base import BaseScraper
from scrapers import location_utils as lu

logger = logging.getLogger(__name__)


class BuilderSitesScraper(BaseScraper):
    """Scrapes luxury builder websites for current/upcoming projects."""

    def __init__(self, config: dict):
        super().__init__(config)
        self.builders = config.get("builders", [])
        self.communities = self._get_all_communities(config)
        self.min_price = config.get("search", {}).get("min_price", 5_000_000)
        self.target_keywords = lu.get_location_keywords(config)
        self.target_zips = lu.get_target_zips(config)
        self.target_cities = lu.get_target_cities(config)
        self.city_regex_alt = lu.build_city_regex_alt(config)
        self.styles = config.get("styles", [
            "Contemporary", "Modern", "Transitional", "Mediterranean",
            "Ranch", "Craftsman", "Mid-Century Modern",
        ])

    def _get_all_communities(self, config: dict) -> List[str]:
        """Extract all community names from config."""
        communities = []
        for region in config.get("communities", {}).values():
            for comm in region:
                communities.append(comm["name"].lower())
                for sub in comm.get("subdivisions", []):
                    communities.append(sub.lower())
        return communities

    def scrape(self) -> List[Property]:
        """Scrape all configured builder websites."""
        all_properties = []

        for builder in self.builders:
            try:
                logger.info(f"Scraping builder: {builder['name']} ({builder['url']})")
                properties = self._scrape_builder(builder)
                all_properties.extend(properties)
                logger.info(f"Found {len(properties)} properties from {builder['name']}")
            except Exception as e:
                logger.error(f"Error scraping {builder['name']}: {e}")

        return all_properties

    def _scrape_builder(self, builder: dict) -> List[Property]:
        """Scrape a single builder's website."""
        base_url = builder["url"].rstrip("/")
        projects_path = builder.get("projects_path", "")
        url = f"{base_url}{projects_path}"

        # Try requests first, fall back to Selenium
        soup = self.fetch_page(url)
        if not soup or not self._has_content(soup):
            logger.info(f"Trying Selenium for {url}")
            soup = self.fetch_page_selenium(url)

        if not soup:
            return []

        properties = self._extract_properties_generic(soup, builder, url)

        # Also try to find sub-pages (project detail pages)
        detail_urls = self._find_project_links(soup, base_url)
        for detail_url in detail_urls[:20]:  # Limit to 20 detail pages per builder
            try:
                detail_soup = self.fetch_page(detail_url)
                if not detail_soup:
                    detail_soup = self.fetch_page_selenium(detail_url)
                if detail_soup:
                    detail_props = self._extract_from_detail_page(
                        detail_soup, builder, detail_url
                    )
                    properties.extend(detail_props)
            except Exception as e:
                logger.warning(f"Error scraping detail page {detail_url}: {e}")

        return properties

    def _has_content(self, soup: BeautifulSoup) -> bool:
        """Check if the page has meaningful content."""
        text = soup.get_text(strip=True)
        return len(text) > 500

    def _find_project_links(self, soup: BeautifulSoup, base_url: str) -> List[str]:
        """Find links to individual project pages."""
        links = []
        project_keywords = [
            "project", "portfolio", "home", "estate", "residence",
            "construction", "building", "current", "community",
        ]

        for a_tag in soup.find_all("a", href=True):
            href = a_tag["href"]
            text = (a_tag.get_text(strip=True) or "").lower()
            href_lower = href.lower()

            # Check if link likely leads to a project page
            is_project_link = any(kw in href_lower or kw in text for kw in project_keywords)
            is_target_area = any(
                loc in href_lower or loc in text
                for loc in self.target_keywords
            )

            if is_project_link or is_target_area:
                full_url = href if href.startswith("http") else f"{base_url}{href}"
                if full_url not in links and base_url in full_url:
                    links.append(full_url)

        return links

    def _extract_properties_generic(
        self, soup: BeautifulSoup, builder: dict, url: str
    ) -> List[Property]:
        """Generic extraction - looks for common patterns across builder sites."""
        properties = []

        # Strategy 1: Look for property cards/items with addresses
        cards = self._find_property_cards(soup)
        for card in cards:
            prop = self._parse_card(card, builder, url)
            if prop and self._is_in_target_area(prop):
                properties.append(prop)

        # Strategy 2: If no cards found, try extracting from page text
        if not properties:
            props = self._extract_from_text(soup, builder, url)
            properties.extend(props)

        return properties

    def _find_property_cards(self, soup: BeautifulSoup) -> List[Tag]:
        """Find elements that look like property listing cards."""
        cards = []

        # Common card selectors
        selectors = [
            "div.property", "div.project", "div.listing", "div.home",
            "article.property", "article.project",
            "div.portfolio-item", "div.gallery-item",
            "div[class*='property']", "div[class*='project']",
            "div[class*='listing']", "div[class*='home-card']",
            "div[class*='estate']", "div[class*='residence']",
            ".card", ".grid-item", ".portfolio-item",
        ]

        for selector in selectors:
            found = soup.select(selector)
            if found:
                cards.extend(found)
                break

        return cards

    def _parse_card(self, card: Tag, builder: dict, url: str) -> Optional[Property]:
        """Parse a property card element into a Property object."""
        text = card.get_text(" ", strip=True)
        if not text:
            return None

        # Extract address (look for Arizona address patterns)
        address = self._extract_address(text)

        # Extract price
        price = self._extract_price(text)

        # Check if price meets minimum
        if price:
            price_num = self._parse_price_number(price)
            if price_num and price_num < self.min_price:
                return None

        # Extract other details
        sqft = self._extract_sqft(text)
        bedrooms = self._extract_bedrooms(text)
        bathrooms = self._extract_bathrooms(text)

        # Determine phase from text
        phase = self._extract_phase(text)

        # Extract community
        community = self._extract_community(text)

        # Get detail link if available
        detail_link = card.find("a", href=True)
        listing_url = detail_link["href"] if detail_link else url

        if not listing_url.startswith("http"):
            listing_url = f"{builder['url'].rstrip('/')}{listing_url}"

        prop = Property(
            address=address,
            city=self._extract_city(text),
            builder=builder["name"],
            price=price,
            sqft=sqft,
            bedrooms=bedrooms,
            bathrooms=bathrooms,
            phase=phase,
            community=community,
            details=text[:2000],  # Truncate long descriptions
            listing_url=listing_url,
            source=f"Builder - {builder['name']}",
            date_scraped=datetime.now().strftime("%Y-%m-%d"),
        )

        return prop

    def _extract_from_detail_page(
        self, soup: BeautifulSoup, builder: dict, url: str
    ) -> List[Property]:
        """Extract property info from a detail/project page."""
        text = soup.get_text(" ", strip=True)

        # Check if this page is about a Scottsdale/PV property
        if not self._text_mentions_target_area(text):
            return []

        address = self._extract_address(text)
        price = self._extract_price(text)
        phase = self._extract_phase(text)
        community = self._extract_community(text)

        # Look for architect/designer info on detail pages
        architect = self._extract_architect(text)
        designer = self._extract_designer(text)

        # Skip if price is known and below threshold
        if price:
            price_num = self._parse_price_number(price)
            if price_num and price_num < self.min_price:
                return []

        # Skip if this appears to be a completed home
        if self._appears_completed(text):
            return []

        prop = Property(
            address=address,
            city=self._extract_city(text),
            builder=builder["name"],
            price=price,
            architect=architect,
            designer=designer,
            sqft=self._extract_sqft(text),
            bedrooms=self._extract_bedrooms(text),
            bathrooms=self._extract_bathrooms(text),
            lot_size=self._extract_lot_size(text),
            phase=phase or "Under Construction",
            community=community,
            style=self._extract_style(text),
            details=text[:2000],
            listing_url=url,
            source=f"Builder - {builder['name']}",
            date_scraped=datetime.now().strftime("%Y-%m-%d"),
        )

        return [prop] if (address or community) else []

    def _extract_from_text(
        self, soup: BeautifulSoup, builder: dict, url: str
    ) -> List[Property]:
        """Try to extract properties from unstructured page text."""
        text = soup.get_text(" ", strip=True)
        properties = []

        if not self._text_mentions_target_area(text):
            return []

        # Look for address patterns in the text
        addresses = re.findall(
            r"\d{1,6}\s+(?:[NSEW]\.?\s+)?[\w\s]+(?:St|Street|Ave|Avenue|Blvd|Boulevard|Dr|Drive|"
            r"Ct|Court|Ln|Lane|Rd|Road|Way|Trail|Trl|Path|Pl|Place)\b",
            text,
            re.IGNORECASE,
        )

        for addr in addresses[:10]:  # Limit
            addr = addr.strip()
            if len(addr) > 10:
                prop = Property(
                    address=addr,
                    city=self._extract_city(text),
                    builder=builder["name"],
                    price=self._extract_price(text),
                    phase=self._extract_phase(text) or "Under Construction",
                    community=self._extract_community(text),
                    details=text[:2000],
                    listing_url=url,
                    source=f"Builder - {builder['name']}",
                    date_scraped=datetime.now().strftime("%Y-%m-%d"),
                )
                if self._is_in_target_area(prop):
                    properties.append(prop)

        return properties

    # --- Extraction helpers ---

    def _extract_address(self, text: str) -> str:
        """Extract a street address from text, preferring target-city matches."""
        patterns = [
            # Standard address anchored to a target city
            r"(\d{1,6}\s+(?:[NSEW]\.?\s+)?[\w\s]+(?:St|Street|Ave|Avenue|Blvd|Boulevard|Dr|Drive|"
            r"Ct|Court|Ln|Lane|Rd|Road|Way|Trail|Trl|Path|Pl|Place))\s*,?\s*"
            r"(?:" + self.city_regex_alt + r")",
            # Lot-style: Lot 5, <community>
            r"(Lot\s+\d+\s*,?\s*[\w\s]+)",
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()

        return ""

    def _extract_price(self, text: str) -> str:
        """Extract price from text."""
        patterns = [
            r"\$\s*([\d,]+(?:\.\d{2})?)\s*(?:million|M)\b",
            r"\$\s*([\d]{1,3}(?:,\d{3})*(?:,\d{3})?)(?:\+)?",
            r"([\d.]+)\s*(?:million|M)\s*(?:dollars?)?\b",
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                raw = match.group(0).strip()
                return raw

        return ""

    def _parse_price_number(self, price_str: str) -> Optional[float]:
        """Parse a price string to a number."""
        if not price_str:
            return None
        # Remove $ and commas
        cleaned = re.sub(r"[$,\s+]", "", price_str)
        # Handle "million" / "M"
        million_match = re.search(r"([\d.]+)\s*(?:million|M)", price_str, re.IGNORECASE)
        if million_match:
            return float(million_match.group(1)) * 1_000_000
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
        if match:
            return f"{match.group(1)} acres"
        match = re.search(r"([\d,]+)\s*(?:sq\.?\s*ft|SF)\s*(?:lot)", text, re.IGNORECASE)
        return f"{match.group(1)} sqft lot" if match else ""

    def _extract_phase(self, text: str) -> str:
        """Determine construction phase from text."""
        text_lower = text.lower()
        phases = [
            ("design phase", "Design"),
            ("in design", "Design"),
            ("pre-construction", "Pre-Construction"),
            ("pre construction", "Pre-Construction"),
            ("preconstruction", "Pre-Construction"),
            ("under construction", "Construction"),
            ("currently under construction", "Construction"),
            ("construction phase", "Construction"),
            ("being built", "Construction"),
            ("now building", "Construction"),
            ("framing", "Framing"),
            ("foundation", "Foundation"),
            ("breaking ground", "Pre-Construction"),
            ("ground breaking", "Pre-Construction"),
            ("planned", "Design"),
            ("proposed", "Design"),
            ("new construction", "Construction"),
            ("coming soon", "Pre-Construction"),
            ("estimated completion", "Construction"),
        ]

        for keyword, phase in phases:
            if keyword in text_lower:
                return phase

        return ""

    def _extract_community(self, text: str) -> str:
        """Extract community name from text (from config)."""
        return lu.detect_community(text, self.config)

    def _extract_city(self, text: str) -> str:
        """Extract target city from text (from config)."""
        return lu.detect_city(text, self.config)

    def _extract_architect(self, text: str) -> str:
        """Try to extract architect name from text."""
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
        """Try to extract interior designer name from text."""
        patterns = [
            r"(?:interior\s+)?designer[:\s]+([A-Z][\w\s&]+?)(?:\s*[,.|])",
            r"interior(?:s)?\s+by\s+([A-Z][\w\s&]+?)(?:\s*[,.|])",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        return ""

    def _extract_style(self, text: str) -> str:
        """Extract architectural style from text (from config)."""
        text_lower = text.lower()
        for style in self.styles:
            if style.lower() in text_lower:
                return style
        return ""

    def _is_in_target_area(self, prop: Property) -> bool:
        """Check if property is in the configured target area."""
        if prop.city:
            return prop.city.lower() in self.target_cities

        combined = f"{prop.address} {prop.details} {prop.community}".lower()
        return any(loc in combined for loc in self.target_keywords)

    def _text_mentions_target_area(self, text: str) -> bool:
        """Check if text mentions any configured target area."""
        return lu.text_mentions_target(text, self.config)

    def _appears_completed(self, text: str) -> bool:
        """Check if text suggests the home is already completed."""
        text_lower = text.lower()
        completed_phrases = [
            "completed in 20", "built in 20", "year built",
            "move-in ready", "just completed", "recently completed",
        ]
        return any(phrase in text_lower for phrase in completed_phrases)
