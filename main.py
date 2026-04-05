#!/usr/bin/env python3
"""
Luxury Homes Scraper - North Scottsdale & Paradise Valley, AZ

Scrapes multiple sources for UHNWI luxury homes ($5M+) that are not yet
completed (design, pre-construction, construction phases).

Usage:
    python main.py                          # Run all scrapers
    python main.py --source builders        # Only builder sites
    python main.py --source listings        # Only listing platforms
    python main.py --source luxury          # Only luxury realty sites
    python main.py --source county          # Only county records
    python main.py --update output/file.xlsx  # Update existing file
    python main.py --no-selenium            # Skip scrapers that need Selenium
"""

import argparse
import logging
import os
import sys
from datetime import datetime

import yaml

from models.property import Property
from scrapers.builder_sites import BuilderSitesScraper
from scrapers.listing_sites import ListingSitesScraper
from scrapers.luxury_realty import LuxuryRealtyScraper
from scrapers.county_records import CountyRecordsScraper
from utils.dedup import deduplicate_properties
from utils.enrichment import enrich_properties
from utils.excel_writer import write_to_excel


def setup_logging(verbose: bool = False):
    """Configure logging."""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(
                f"scraper_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
            ),
        ],
    )


def load_config(config_path: str = "config.yaml") -> dict:
    """Load configuration from YAML file."""
    if not os.path.exists(config_path):
        logging.error(f"Config file not found: {config_path}")
        sys.exit(1)

    with open(config_path, "r") as f:
        return yaml.safe_load(f)


def run_scrapers(config: dict, sources: list, use_selenium: bool = True) -> list:
    """Run selected scrapers and collect properties."""
    all_properties = []
    run_all = not sources or "all" in sources

    # 1. Builder sites
    if run_all or "builders" in sources:
        logging.info("=" * 60)
        logging.info("SCRAPING BUILDER WEBSITES")
        logging.info("=" * 60)
        try:
            with BuilderSitesScraper(config) as scraper:
                props = scraper.scrape()
                logging.info(f"Builder sites: found {len(props)} properties")
                all_properties.extend(props)
        except Exception as e:
            logging.error(f"Builder sites scraper failed: {e}")

    # 2. Listing platforms
    if run_all or "listings" in sources:
        logging.info("=" * 60)
        logging.info("SCRAPING LISTING PLATFORMS")
        logging.info("=" * 60)
        if not use_selenium:
            logging.warning("Skipping listing sites (requires Selenium)")
        else:
            try:
                with ListingSitesScraper(config) as scraper:
                    props = scraper.scrape()
                    logging.info(f"Listing sites: found {len(props)} properties")
                    all_properties.extend(props)
            except Exception as e:
                logging.error(f"Listing sites scraper failed: {e}")

    # 3. Luxury realty sites
    if run_all or "luxury" in sources:
        logging.info("=" * 60)
        logging.info("SCRAPING LUXURY REALTY SITES")
        logging.info("=" * 60)
        try:
            with LuxuryRealtyScraper(config) as scraper:
                props = scraper.scrape()
                logging.info(f"Luxury realty sites: found {len(props)} properties")
                all_properties.extend(props)
        except Exception as e:
            logging.error(f"Luxury realty scraper failed: {e}")

    # 4. County records
    if run_all or "county" in sources:
        logging.info("=" * 60)
        logging.info("SCRAPING COUNTY RECORDS")
        logging.info("=" * 60)
        try:
            with CountyRecordsScraper(config) as scraper:
                # Direct permit search
                props = scraper.scrape()
                logging.info(f"County records: found {len(props)} properties")
                all_properties.extend(props)

                # Enrich existing properties with county data
                if all_properties:
                    addresses = [p.address for p in all_properties if p.address]
                    if addresses:
                        logging.info(f"Looking up {len(addresses)} addresses in county records...")
                        county_data = scraper.lookup_properties(addresses)
                        logging.info(f"Found county data for {len(county_data)} addresses")

                        # Also enrich with permit data
                        scraper.enrich_with_permits(all_properties)
        except Exception as e:
            logging.error(f"County records scraper failed: {e}")

    return all_properties


def process_properties(properties: list, county_data: dict = None) -> list:
    """Process, deduplicate, enrich, and filter properties."""
    logging.info(f"Processing {len(properties)} raw properties...")

    # Deduplicate
    properties = deduplicate_properties(properties)
    logging.info(f"After deduplication: {len(properties)} properties")

    # Enrich (applies door_details, infers communities, filters completed homes)
    properties = enrich_properties(properties, county_data)
    logging.info(f"After enrichment and filtering: {len(properties)} properties")

    # Final filter: remove properties without address AND without community
    properties = [
        p for p in properties
        if p.address or p.community
    ]
    logging.info(f"Final property count: {len(properties)}")

    return properties


def print_summary(properties: list):
    """Print a summary of scraped results."""
    print("\n" + "=" * 60)
    print("SCRAPING RESULTS SUMMARY")
    print("=" * 60)
    print(f"Total properties found: {len(properties)}")

    if not properties:
        print("No properties found matching criteria.")
        return

    # Count by source
    sources = {}
    for p in properties:
        for src in (p.source or "Unknown").split("; "):
            src = src.strip()
            sources[src] = sources.get(src, 0) + 1
    print("\nBy source:")
    for src, count in sorted(sources.items(), key=lambda x: -x[1]):
        print(f"  {src}: {count}")

    # Count by phase
    phases = {}
    for p in properties:
        phase = p.phase or "Unknown"
        phases[phase] = phases.get(phase, 0) + 1
    print("\nBy phase:")
    for phase, count in sorted(phases.items(), key=lambda x: -x[1]):
        print(f"  {phase}: {count}")

    # Count by community
    communities = {}
    for p in properties:
        comm = p.community or "Unknown"
        communities[comm] = communities.get(comm, 0) + 1
    print("\nBy community:")
    for comm, count in sorted(communities.items(), key=lambda x: -x[1]):
        print(f"  {comm}: {count}")

    # Count by city
    cities = {}
    for p in properties:
        city = p.city or "Unknown"
        cities[city] = cities.get(city, 0) + 1
    print("\nBy city:")
    for city, count in sorted(cities.items(), key=lambda x: -x[1]):
        print(f"  {city}: {count}")

    # Properties with door_details
    door_details_count = sum(1 for p in properties if p.door_details)
    if door_details_count:
        print(f"\nProperties with door_details: {door_details_count}")

    # Properties with architect/designer info
    architect_count = sum(1 for p in properties if p.architect)
    designer_count = sum(1 for p in properties if p.designer)
    builder_count = sum(1 for p in properties if p.builder)
    print(f"\nData completeness:")
    print(f"  With architect: {architect_count}")
    print(f"  With designer: {designer_count}")
    print(f"  With builder: {builder_count}")

    print("=" * 60)


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Scrape luxury homes ($5M+) in North Scottsdale & Paradise Valley, AZ"
    )
    parser.add_argument(
        "--source",
        nargs="+",
        choices=["all", "builders", "listings", "luxury", "county"],
        default=["all"],
        help="Which scrapers to run (default: all)",
    )
    parser.add_argument(
        "--config",
        default="config.yaml",
        help="Path to config file (default: config.yaml)",
    )
    parser.add_argument(
        "--update",
        metavar="FILE",
        help="Update an existing Excel file instead of creating new",
    )
    parser.add_argument(
        "--no-selenium",
        action="store_true",
        help="Skip scrapers that require Selenium",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Enable verbose/debug logging",
    )

    args = parser.parse_args()

    # Setup
    setup_logging(args.verbose)
    logging.info("Luxury Homes Scraper starting...")
    logging.info(f"Sources: {args.source}")

    config = load_config(args.config)
    min_price = config.get("search", {}).get("min_price", 5_000_000)
    logging.info(f"Minimum price: ${min_price:,.0f}")

    # Run scrapers
    properties = run_scrapers(
        config,
        args.source,
        use_selenium=not args.no_selenium,
    )

    # Process results
    properties = process_properties(properties)

    # Write to Excel
    if properties:
        output_cfg = config.get("output", {})
        output_path = write_to_excel(
            properties,
            output_dir=output_cfg.get("directory", "output"),
            filename_prefix=output_cfg.get("filename_prefix", "luxury_homes_scottsdale_pv"),
            update_file=args.update,
        )
        logging.info(f"Results written to: {output_path}")
        print(f"\nExcel file saved: {output_path}")
    else:
        logging.warning("No properties found. No Excel file generated.")
        print("\nNo properties found matching the criteria.")

    # Print summary
    print_summary(properties)

    logging.info("Scraping complete.")


if __name__ == "__main__":
    main()
