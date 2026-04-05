"""Scraper for major real estate listing platforms (Zillow, Redfin, Realtor.com)."""

import logging
import re
import json
from typing import List, Optional
from datetime import datetime

from bs4 import BeautifulSoup

from models.property import Property
from scrapers.base import BaseScraper
from scrapers import location_utils as lu

logger = logging.getLogger(__name__)


class ListingSitesScraper(BaseScraper):
    """Scrapes major listing platforms for new construction luxury homes."""

    def __init__(self, config: dict):
        super().__init__(config)
        self.listing_sites = config.get("listing_sites", [])
        self.min_price = config.get("search", {}).get("min_price", 5_000_000)
        self.city_regex_alt = lu.build_city_regex_alt(config)

    def scrape(self) -> List[Property]:
        """Scrape all configured listing sites."""
        all_properties = []

        for site in self.listing_sites:
            try:
                name = site["name"]
                logger.info(f"Scraping listing site: {name}")

                if "zillow" in name.lower():
                    props = self._scrape_zillow(site)
                elif "realtor" in name.lower():
                    props = self._scrape_realtor(site)
                elif "redfin" in name.lower():
                    props = self._scrape_redfin(site)
                else:
                    props = self._scrape_generic_listing(site)

                all_properties.extend(props)
                logger.info(f"Found {len(props)} properties from {name}")
            except Exception as e:
                logger.error(f"Error scraping {site['name']}: {e}")

        return all_properties

    def _scrape_zillow(self, site: dict) -> List[Property]:
        """Scrape Zillow new construction listings."""
        properties = []

        for search_url in site.get("search_urls", []):
            url = f"{site['base_url']}{search_url}"
            logger.info(f"Scraping Zillow URL: {url}")

            # Zillow requires Selenium due to heavy JS rendering
            soup = self.fetch_page_selenium(url)
            if not soup:
                continue

            # Try to extract data from Zillow's JSON-LD or script tags
            props = self._extract_zillow_json(soup, url)
            if props:
                properties.extend(props)
                continue

            # Fall back to HTML parsing
            cards = soup.select(
                "article.list-card, div[class*='ListItem'], "
                "li[class*='ListItem'], div[id*='grid-search-results'] li"
            )

            for card in cards:
                prop = self._parse_zillow_card(card, url)
                if prop:
                    properties.append(prop)

        return properties

    def _extract_zillow_json(self, soup: BeautifulSoup, source_url: str) -> List[Property]:
        """Try to extract property data from Zillow's embedded JSON."""
        properties = []

        # Look for Next.js data or JSON-LD
        for script in soup.find_all("script", type="application/json"):
            try:
                data = json.loads(script.string or "")
                props = self._parse_zillow_json_data(data, source_url)
                properties.extend(props)
            except (json.JSONDecodeError, TypeError):
                continue

        # Also check for __NEXT_DATA__
        for script in soup.find_all("script", id="__NEXT_DATA__"):
            try:
                data = json.loads(script.string or "")
                props = self._parse_zillow_next_data(data, source_url)
                properties.extend(props)
            except (json.JSONDecodeError, TypeError):
                continue

        return properties

    def _parse_zillow_json_data(self, data: dict, source_url: str) -> List[Property]:
        """Parse Zillow JSON data structure."""
        properties = []

        # Navigate through various Zillow JSON structures
        results = []
        if isinstance(data, dict):
            # Try common Zillow data paths
            for key in ["searchResults", "results", "listResults", "props"]:
                if key in data:
                    results = data[key] if isinstance(data[key], list) else [data[key]]
                    break

            # Try nested cat1/searchResults path
            cat1 = data.get("cat1", {})
            if isinstance(cat1, dict):
                search_results = cat1.get("searchResults", {})
                if isinstance(search_results, dict):
                    results = search_results.get("listResults", [])

        for result in results:
            if not isinstance(result, dict):
                continue

            address = result.get("address", "") or result.get("streetAddress", "")
            price = result.get("price", "") or result.get("unformattedPrice", "")
            if isinstance(price, (int, float)):
                if price < self.min_price:
                    continue
                price = f"${price:,.0f}"

            prop = Property(
                address=str(address),
                city=result.get("city", "") or result.get("addressCity", ""),
                zip_code=str(result.get("zipcode", "") or result.get("addressZipcode", "")),
                price=str(price),
                sqft=str(result.get("livingArea", "") or result.get("area", "")),
                bedrooms=str(result.get("beds", "") or result.get("bedrooms", "")),
                bathrooms=str(result.get("baths", "") or result.get("bathrooms", "")),
                lot_size=str(result.get("lotAreaValue", "")),
                builder=result.get("builderName", "") or result.get("attributionInfo", {}).get("builderName", ""),
                phase="New Construction",
                details=result.get("description", "") or "",
                listing_url=result.get("detailUrl", "") or result.get("url", source_url),
                source="Listing - Zillow",
                date_scraped=datetime.now().strftime("%Y-%m-%d"),
            )
            properties.append(prop)

        return properties

    def _parse_zillow_next_data(self, data: dict, source_url: str) -> List[Property]:
        """Parse Zillow __NEXT_DATA__ structure."""
        properties = []
        try:
            # Navigate Next.js data structure
            props_data = data.get("props", {}).get("pageProps", {})
            search_results = (
                props_data.get("searchPageState", {})
                .get("cat1", {})
                .get("searchResults", {})
                .get("listResults", [])
            )

            for result in search_results:
                if not isinstance(result, dict):
                    continue

                price_raw = result.get("unformattedPrice", 0) or result.get("price", 0)
                if isinstance(price_raw, str):
                    price_raw = int(re.sub(r"[^\d]", "", price_raw) or 0)

                if price_raw and price_raw < self.min_price:
                    continue

                addr_info = result.get("address", "") if isinstance(result.get("address"), str) else ""
                hdp_data = result.get("hdpData", {}).get("homeInfo", {})

                prop = Property(
                    address=addr_info or hdp_data.get("streetAddress", ""),
                    city=hdp_data.get("city", ""),
                    zip_code=str(hdp_data.get("zipcode", "")),
                    price=f"${price_raw:,.0f}" if price_raw else "",
                    sqft=str(hdp_data.get("livingArea", "")),
                    bedrooms=str(hdp_data.get("bedrooms", "")),
                    bathrooms=str(hdp_data.get("bathrooms", "")),
                    lot_size=str(hdp_data.get("lotAreaValue", "")),
                    phase="New Construction",
                    listing_url=result.get("detailUrl", source_url),
                    source="Listing - Zillow",
                    date_scraped=datetime.now().strftime("%Y-%m-%d"),
                )
                properties.append(prop)
        except (KeyError, TypeError, AttributeError) as e:
            logger.debug(f"Error parsing Zillow Next data: {e}")

        return properties

    def _parse_zillow_card(self, card, source_url: str) -> Optional[Property]:
        """Parse a Zillow listing card HTML element."""
        text = card.get_text(" ", strip=True)

        address = ""
        addr_el = card.select_one("address, [data-test='property-card-addr']")
        if addr_el:
            address = addr_el.get_text(strip=True)

        price = self._extract_price(text)
        if price:
            price_num = self._parse_price_number(price)
            if price_num and price_num < self.min_price:
                return None

        detail_link = card.find("a", href=True)
        listing_url = detail_link["href"] if detail_link else source_url
        if listing_url and not listing_url.startswith("http"):
            listing_url = f"https://www.zillow.com{listing_url}"

        return Property(
            address=address,
            price=price,
            sqft=self._extract_sqft(text),
            bedrooms=self._extract_bedrooms(text),
            bathrooms=self._extract_bathrooms(text),
            phase="New Construction",
            details=text[:2000],
            listing_url=listing_url,
            source="Listing - Zillow",
            date_scraped=datetime.now().strftime("%Y-%m-%d"),
        )

    def _scrape_realtor(self, site: dict) -> List[Property]:
        """Scrape Realtor.com new construction listings."""
        properties = []

        for search_url in site.get("search_urls", []):
            url = f"{site['base_url']}{search_url}"
            logger.info(f"Scraping Realtor.com URL: {url}")

            soup = self.fetch_page_selenium(url)
            if not soup:
                continue

            # Realtor.com uses data attributes and React
            cards = soup.select(
                "div[data-testid='property-card'], "
                "li[data-testid='result-card'], "
                "div.BasePropertyCard_propertyCardWrap"
            )

            for card in cards:
                text = card.get_text(" ", strip=True)
                address = ""
                addr_el = card.select_one(
                    "[data-testid='card-address'], .card-address, "
                    "div[class*='address']"
                )
                if addr_el:
                    address = addr_el.get_text(strip=True)

                price = self._extract_price(text)
                if price:
                    price_num = self._parse_price_number(price)
                    if price_num and price_num < self.min_price:
                        continue

                detail_link = card.find("a", href=True)
                listing_url = detail_link["href"] if detail_link else url
                if listing_url and not listing_url.startswith("http"):
                    listing_url = f"https://www.realtor.com{listing_url}"

                prop = Property(
                    address=address,
                    price=price,
                    sqft=self._extract_sqft(text),
                    bedrooms=self._extract_bedrooms(text),
                    bathrooms=self._extract_bathrooms(text),
                    phase="New Construction",
                    details=text[:2000],
                    listing_url=listing_url,
                    source="Listing - Realtor.com",
                    date_scraped=datetime.now().strftime("%Y-%m-%d"),
                )
                properties.append(prop)

        return properties

    def _scrape_redfin(self, site: dict) -> List[Property]:
        """Scrape Redfin new construction listings."""
        properties = []

        for search_url in site.get("search_urls", []):
            url = f"{site['base_url']}{search_url}"
            logger.info(f"Scraping Redfin URL: {url}")

            soup = self.fetch_page_selenium(url)
            if not soup:
                continue

            # Try to extract from Redfin's embedded data
            for script in soup.find_all("script"):
                if script.string and "reactServerAgent.OpenRedfinWindow" in (script.string or ""):
                    props = self._parse_redfin_script(script.string, url)
                    properties.extend(props)

            # Fall back to HTML card parsing
            cards = soup.select(
                "div.HomeCard, div[class*='HomeCard'], "
                "div.MapHomeCard, div[data-rf-test-id='MapHomeCard']"
            )

            for card in cards:
                text = card.get_text(" ", strip=True)

                address = ""
                addr_el = card.select_one(
                    ".homeAddressV2, .link-and-anchor, "
                    "div[class*='address'], a[class*='homecardV2']"
                )
                if addr_el:
                    address = addr_el.get_text(strip=True)

                price = self._extract_price(text)
                if price:
                    price_num = self._parse_price_number(price)
                    if price_num and price_num < self.min_price:
                        continue

                detail_link = card.find("a", href=True)
                listing_url = detail_link["href"] if detail_link else url
                if listing_url and not listing_url.startswith("http"):
                    listing_url = f"https://www.redfin.com{listing_url}"

                prop = Property(
                    address=address,
                    price=price,
                    sqft=self._extract_sqft(text),
                    bedrooms=self._extract_bedrooms(text),
                    bathrooms=self._extract_bathrooms(text),
                    phase="New Construction",
                    details=text[:2000],
                    listing_url=listing_url,
                    source="Listing - Redfin",
                    date_scraped=datetime.now().strftime("%Y-%m-%d"),
                )
                properties.append(prop)

        return properties

    def _parse_redfin_script(self, script_text: str, source_url: str) -> List[Property]:
        """Parse Redfin's embedded JavaScript data."""
        properties = []
        try:
            # Redfin embeds JSON in script tags
            json_match = re.search(r"window\.__reactServerState\s*=\s*({.*?});", script_text, re.DOTALL)
            if not json_match:
                return []

            data = json.loads(json_match.group(1))
            homes = data.get("searchData", {}).get("homes", [])

            for home in homes:
                price = home.get("price", {}).get("value", 0)
                if price and price < self.min_price:
                    continue

                prop = Property(
                    address=home.get("streetLine", {}).get("value", ""),
                    city=home.get("city", ""),
                    zip_code=str(home.get("zip", "")),
                    price=f"${price:,.0f}" if price else "",
                    sqft=str(home.get("sqFt", {}).get("value", "")),
                    bedrooms=str(home.get("beds", "")),
                    bathrooms=str(home.get("baths", "")),
                    lot_size=str(home.get("lotSize", {}).get("value", "")),
                    phase="New Construction",
                    listing_url=home.get("url", source_url),
                    source="Listing - Redfin",
                    date_scraped=datetime.now().strftime("%Y-%m-%d"),
                )
                properties.append(prop)
        except (json.JSONDecodeError, KeyError, TypeError) as e:
            logger.debug(f"Error parsing Redfin script data: {e}")

        return properties

    def _scrape_generic_listing(self, site: dict) -> List[Property]:
        """Generic listing site scraper."""
        properties = []

        for search_url in site.get("search_urls", []):
            url = f"{site['base_url']}{search_url}"
            soup = self.fetch_page(url) or self.fetch_page_selenium(url)
            if not soup:
                continue

            # Look for common listing card patterns
            cards = soup.select(
                "div[class*='listing'], div[class*='property'], "
                "article, .card, li[class*='result']"
            )

            for card in cards:
                text = card.get_text(" ", strip=True)
                address = self._extract_address_from_text(text)
                price = self._extract_price(text)

                if price:
                    price_num = self._parse_price_number(price)
                    if price_num and price_num < self.min_price:
                        continue

                if address:
                    prop = Property(
                        address=address,
                        price=price,
                        sqft=self._extract_sqft(text),
                        bedrooms=self._extract_bedrooms(text),
                        bathrooms=self._extract_bathrooms(text),
                        phase="New Construction",
                        details=text[:2000],
                        listing_url=url,
                        source=f"Listing - {site['name']}",
                        date_scraped=datetime.now().strftime("%Y-%m-%d"),
                    )
                    properties.append(prop)

        return properties

    # --- Helper methods ---

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
        cleaned = re.sub(r"[$,\s+]", "", price_str)
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

    def _extract_address_from_text(self, text: str) -> str:
        match = re.search(
            r"(\d{1,6}\s+(?:[NSEW]\.?\s+)?[\w\s]+(?:St|Street|Ave|Avenue|Blvd|Boulevard|Dr|Drive|"
            r"Ct|Court|Ln|Lane|Rd|Road|Way|Trail|Trl|Path|Pl|Place))\s*,?\s*"
            r"(?:" + self.city_regex_alt + r")?",
            text,
            re.IGNORECASE,
        )
        return match.group(1).strip() if match else ""
