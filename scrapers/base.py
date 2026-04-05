"""Base scraper class with shared functionality."""

import logging
import random
import time
from typing import Optional

import requests
from bs4 import BeautifulSoup
from fake_useragent import UserAgent

logger = logging.getLogger(__name__)


class BaseScraper:
    """Base class for all scrapers with common HTTP/Selenium functionality."""

    def __init__(self, config: dict):
        self.config = config
        scraping_cfg = config.get("scraping", {})
        self.delay_min = scraping_cfg.get("delay_min", 2)
        self.delay_max = scraping_cfg.get("delay_max", 5)
        self.max_retries = scraping_cfg.get("max_retries", 3)
        self.timeout = scraping_cfg.get("request_timeout", 30)
        self.headless = scraping_cfg.get("selenium_headless", True)

        self.session = requests.Session()
        self._ua = UserAgent(fallback="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
        self._driver = None

    def _get_headers(self) -> dict:
        """Generate request headers with a random user agent."""
        return {
            "User-Agent": self._ua.random,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
            "Accept-Encoding": "gzip, deflate, br",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
        }

    def _delay(self):
        """Random delay between requests to be respectful."""
        delay = random.uniform(self.delay_min, self.delay_max)
        time.sleep(delay)

    def fetch_page(self, url: str) -> Optional[BeautifulSoup]:
        """Fetch a page using requests and return parsed BeautifulSoup.

        Retries with exponential backoff on failure.
        """
        for attempt in range(self.max_retries):
            try:
                self._delay()
                response = self.session.get(
                    url,
                    headers=self._get_headers(),
                    timeout=self.timeout,
                )
                response.raise_for_status()
                return BeautifulSoup(response.text, "lxml")
            except requests.RequestException as e:
                wait = 2 ** (attempt + 1)
                logger.warning(f"Request failed for {url} (attempt {attempt + 1}): {e}. Retrying in {wait}s...")
                time.sleep(wait)

        logger.error(f"Failed to fetch {url} after {self.max_retries} attempts")
        return None

    def fetch_page_selenium(self, url: str) -> Optional[BeautifulSoup]:
        """Fetch a JS-rendered page using Selenium and return parsed BeautifulSoup."""
        driver = self._get_driver()
        if not driver:
            return None

        for attempt in range(self.max_retries):
            try:
                self._delay()
                driver.get(url)
                # Wait for page to load
                time.sleep(3)
                page_source = driver.page_source
                return BeautifulSoup(page_source, "lxml")
            except Exception as e:
                wait = 2 ** (attempt + 1)
                logger.warning(
                    f"Selenium request failed for {url} (attempt {attempt + 1}): {e}. Retrying in {wait}s..."
                )
                time.sleep(wait)

        logger.error(f"Selenium failed to fetch {url} after {self.max_retries} attempts")
        return None

    def _get_driver(self):
        """Initialize and return a Selenium WebDriver."""
        if self._driver:
            return self._driver

        try:
            from selenium import webdriver
            from selenium.webdriver.chrome.options import Options
            from selenium.webdriver.chrome.service import Service
            from webdriver_manager.chrome import ChromeDriverManager

            options = Options()
            if self.headless:
                options.add_argument("--headless=new")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-gpu")
            options.add_argument("--window-size=1920,1080")
            options.add_argument(f"--user-agent={self._ua.random}")

            # Stealth options to reduce detection
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option("useAutomationExtension", False)

            service = Service(ChromeDriverManager().install())
            self._driver = webdriver.Chrome(service=service, options=options)
            self._driver.execute_cdp_cmd(
                "Page.addScriptToEvaluateOnNewDocument",
                {"source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"},
            )
            return self._driver
        except Exception as e:
            logger.error(f"Failed to initialize Selenium WebDriver: {e}")
            return None

    def close(self):
        """Clean up resources."""
        if self._driver:
            try:
                self._driver.quit()
            except Exception:
                pass
            self._driver = None
        self.session.close()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()
