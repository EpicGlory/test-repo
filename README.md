# Luxury Homes Scraper — Summit County, UT

Scrapes multiple sources (builder sites, luxury brokerages, major listing
platforms, and county assessor/permit records) for ultra-high-net-worth
($5M+) residences in Summit County, Utah — primarily Park City, Deer
Valley, The Colony at White Pine Canyon, Promontory, Glenwild, and the
Snyderville Basin — that are still under construction, pre-construction,
or in the design phase.

The target market (cities, zips, communities, builders, brokerages,
county records endpoints) is fully defined in `config.yaml`, mirroring
the process used for the Scottsdale / Paradise Valley scraper session so
the same pipeline applies to a new market.

## Usage

```bash
pip install -r requirements.txt
python main.py                          # Run all scrapers
python main.py --source builders        # Only builder sites
python main.py --source listings        # Only listing platforms
python main.py --source luxury          # Only luxury brokerage sites
python main.py --source county          # Only Summit County records
python main.py --no-selenium            # Skip Selenium-dependent scrapers
python main.py --update output/file.xlsx  # Update existing Excel output
```

## Target market (see `config.yaml`)

- **Cities / Zips:** Park City (84060, 84098), Kamas (84036),
  Coalville (84017), Oakley (84055), Peoa (84061)
- **Communities:** Deer Valley (Empire Pass, Silver Lake, Bald Eagle,
  Deer Crest), The Colony at White Pine Canyon, Promontory, Glenwild,
  Red Cloud, Stein Eriksen Residences, Park Meadows, Old Town,
  Thaynes Canyon, Aerie, Jeremy Ranch, Pinebrook, Silver Creek,
  Summit Park, Kimball Junction, Wohali
- **County records:** Summit County UT Assessor tax search, Summit
  County CitizenServe permit portal, Park City Municipal EnerGov
  permit portal (Park City is a separate jurisdiction inside Summit
  County)

## Output

Results are written to `output/luxury_homes_summit_county_ut_YYYY-MM-DD.xlsx`
with properties deduplicated across sources and phase/year-built filters
applied to exclude completed homes.
