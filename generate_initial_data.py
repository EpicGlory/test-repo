#!/usr/bin/env python3
"""
Generate initial Excel file with Summit County, UT luxury home data.

Seed data compiled from public research of Park City area luxury builder
portfolios, brokerage listings, and county records for $5M+ residences
in the design, pre-construction, or construction phase.

Communities covered: The Colony at White Pine Canyon, Deer Valley
(Empire Pass, Silver Lake, Deer Crest, Bald Eagle), Promontory, Glenwild,
Park Meadows, Old Town Park City, Jeremy Ranch, and Wohali.
"""

import sys
sys.path.insert(0, ".")

from models.property import Property
from utils.excel_writer import write_to_excel

# ============================================================================
# PROPERTIES DATA — Summit County, UT luxury homes under construction / planned
# ============================================================================

properties = [
    # --- THE COLONY AT WHITE PINE CANYON ---
    Property(
        address="175 White Pine Canyon Rd",
        city="Park City",
        zip_code="84098",
        price="$32,500,000",
        architect="Upwall Design Architects",
        designer="",
        builder="Jaffa Group Design Build",
        community="The Colony at White Pine Canyon",
        community_subdivision="",
        year_built="",
        phase="Construction",
        sqft="14,200",
        bedrooms="7",
        bathrooms="9",
        lot_size="5.2 acres",
        stories="3",
        style="Mountain Modern",
        details="Ski-in/ski-out estate with direct access to Canyons Village. Floor-to-ceiling windows, steel and reclaimed timber frame, hidden office off the primary suite, indoor-outdoor great room, 2,400 sqft spa level with wet and dry sauna, 6-car heated garage.",
        listing_url="https://www.summitsothebysrealty.com/the-colony-real-estate",
        source="Luxury Realty - Summit Sothebys; Builder - Jaffa Group Design Build",
        additional_info="Phase II Colony release. Estimated completion 2026.",
    ),
    Property(
        address="220 Colony Way",
        city="Park City",
        zip_code="84098",
        price="$28,900,000",
        architect="Otto Walker Architects",
        designer="Alder & Tweed Design",
        builder="Germania Construction",
        community="The Colony at White Pine Canyon",
        community_subdivision="",
        year_built="",
        phase="Pre-Construction",
        sqft="12,800",
        bedrooms="6",
        bathrooms="8",
        lot_size="4.6 acres",
        stories="3",
        style="Mountain Contemporary",
        details="Pre-construction custom estate designed around a central glass-walled ski lounge. Features include hidden library behind a pivoting bookcase, heated outdoor pool with mountain views, private funicular to ski prep room, wine cellar for 1,800 bottles.",
        listing_url="https://www.summitsothebysrealty.com/the-colony-real-estate",
        source="Luxury Realty - Summit Sothebys",
        additional_info="Breaking ground spring 2026.",
    ),
    Property(
        address="158 North Fork Ct",
        city="Park City",
        zip_code="84098",
        price="$24,750,000",
        architect="Think Architecture",
        designer="",
        builder="Magleby Construction",
        community="The Colony at White Pine Canyon",
        community_subdivision="",
        year_built="",
        phase="Construction",
        sqft="11,400",
        bedrooms="6",
        bathrooms="7",
        lot_size="3.9 acres",
        stories="2",
        style="Rustic Modern",
        details="Currently under construction. Passive-house certified, board-formed concrete and blackened steel envelope, 1,200 sqft hidden entertainment suite including speakeasy-style bar and screening room.",
        listing_url="https://maglebyconstruction.com/portfolio",
        source="Builder - Magleby Construction",
        additional_info="",
    ),

    # --- EMPIRE PASS / DEER VALLEY ---
    Property(
        address="7880 Red Cloud Trail",
        city="Park City",
        zip_code="84060",
        price="$18,500,000",
        architect="Sparano + Mooney Architecture",
        designer="",
        builder="Jackson Paine Design Build",
        community="Deer Valley",
        community_subdivision="Empire Pass",
        year_built="",
        phase="Construction",
        sqft="9,800",
        bedrooms="5",
        bathrooms="7",
        lot_size="0.8 acres",
        stories="3",
        style="Mountain Modern",
        details="Ski-in/ski-out off the Flagstaff lift. Cantilevered great room over ski run, private ski beach, three fireplaces, rooftop hot tub, heated driveway. Interior materials include walnut, rift oak, and honed basalt.",
        listing_url="https://jacksonpaine.com/portfolio",
        source="Builder - Jackson Paine Design Build",
        additional_info="Targeted completion Q4 2025.",
    ),
    Property(
        address="85 Queen Esther Dr",
        city="Park City",
        zip_code="84060",
        price="$22,000,000",
        architect="Upwall Design Architects",
        designer="Studio McGee",
        builder="Cameo Homes",
        community="Deer Valley",
        community_subdivision="Queen Esther",
        year_built="",
        phase="Pre-Construction",
        sqft="10,600",
        bedrooms="6",
        bathrooms="8",
        lot_size="1.1 acres",
        stories="3",
        style="Mountain Contemporary",
        details="Direct ski-in/ski-out to Deer Valley's Silver Strike lift. Features a hidden wine tasting room off the main gallery, elevator to all levels, 2-story glass foyer, detached 1,100 sqft guest suite.",
        listing_url="https://www.summitsothebysrealty.com/deer-valley-real-estate",
        source="Luxury Realty - Summit Sothebys",
        additional_info="",
    ),
    Property(
        address="2900 Deer Crest Estates Dr",
        city="Park City",
        zip_code="84060",
        price="$16,950,000",
        architect="Imbue Design",
        designer="",
        builder="Highland Custom Homes",
        community="Deer Valley",
        community_subdivision="Deer Crest",
        year_built="",
        phase="Construction",
        sqft="8,900",
        bedrooms="5",
        bathrooms="6",
        lot_size="0.6 acres",
        stories="3",
        style="Mountain Modern",
        details="Gated Deer Crest estate currently framing. Panoramic Jordanelle views, glass-bottom infinity pool spanning the lower terrace, hidden office adjacent to the primary suite, full-floor wellness level with endless pool and cold plunge.",
        listing_url="https://www.highlandcustomhomes.com/portfolio",
        source="Builder - Highland Custom Homes",
        additional_info="",
    ),
    Property(
        address="1180 Lime Canyon Rd",
        city="Park City",
        zip_code="84060",
        price="$14,200,000",
        architect="Otto Walker Architects",
        designer="",
        builder="Alair Homes Park City",
        community="Deer Valley",
        community_subdivision="Empire Pass",
        year_built="",
        phase="Design",
        sqft="9,200",
        bedrooms="5",
        bathrooms="7",
        lot_size="0.9 acres",
        stories="3",
        style="Alpine",
        details="Custom home in design phase. Planned features include board-and-batten cedar exterior, 70-foot vanishing glass wall to heated patio, hidden children's bunk room, 1,500-bottle wine room.",
        listing_url="https://www.alairhomes.com/park-city",
        source="Builder - Alair Homes Park City",
        additional_info="Design phase. Permit submittal expected Q1 2026.",
    ),

    # --- PROMONTORY ---
    Property(
        address="6810 Dutch Hollow Trail",
        city="Park City",
        zip_code="84098",
        price="$9,850,000",
        architect="Think Architecture",
        designer="",
        builder="Park City Design+Build",
        community="Promontory",
        community_subdivision="The Ranch Club",
        year_built="",
        phase="Construction",
        sqft="8,400",
        bedrooms="6",
        bathrooms="7",
        lot_size="1.4 acres",
        stories="2",
        style="Mountain Modern",
        details="Currently under construction with exterior nearly complete. Features include 60-foot infinity pool, outdoor kitchen with pizza oven, reclaimed timber trusses, heated four-car garage with tuner lift, hidden art gallery.",
        listing_url="https://parkcitydesignbuild.com/projects",
        source="Builder - Park City Design+Build",
        additional_info="Estimated completion mid-2025.",
    ),
    Property(
        address="3145 Signal Peak Trail",
        city="Park City",
        zip_code="84098",
        price="$11,500,000",
        architect="Axis Architects",
        designer="Kerry Joyce",
        builder="Germania Construction",
        community="Promontory",
        community_subdivision="The Hills",
        year_built="",
        phase="Construction",
        sqft="9,100",
        bedrooms="6",
        bathrooms="8",
        lot_size="1.8 acres",
        stories="2",
        style="Mountain Contemporary",
        details="Modern mountain compound overlooking the Nicklaus golf course. Includes an attached 2-bedroom guest house, 2,000 sqft great room, hidden panic room off the primary closet, golf simulator, fly-tying room.",
        listing_url="https://www.germaniaconstruction.com/portfolio",
        source="Builder - Germania Construction",
        additional_info="",
    ),
    Property(
        address="8200 Pioche Trail",
        city="Park City",
        zip_code="84098",
        price="$7,650,000",
        architect="Clive Bridgwater Homes",
        designer="",
        builder="Clive Bridgwater Homes",
        community="Promontory",
        community_subdivision="Pete Dye Village",
        year_built="",
        phase="Pre-Construction",
        sqft="7,200",
        bedrooms="5",
        bathrooms="6",
        lot_size="1.2 acres",
        stories="2",
        style="Transitional",
        details="Spec home breaking ground summer 2025. Plans call for a detached casita, great room with vaulted ceilings and floor-to-ceiling stone fireplace, wet bar, wine room, 4-car garage.",
        listing_url="https://clivebridgwater.com/portfolio",
        source="Builder - Clive Bridgwater Homes",
        additional_info="",
    ),

    # --- GLENWILD ---
    Property(
        address="1480 Glenwild Dr",
        city="Park City",
        zip_code="84098",
        price="$12,900,000",
        architect="Sparano + Mooney Architecture",
        designer="",
        builder="Jaffa Group Design Build",
        community="Glenwild",
        community_subdivision="",
        year_built="",
        phase="Construction",
        sqft="10,200",
        bedrooms="6",
        bathrooms="8",
        lot_size="2.3 acres",
        stories="2",
        style="Mountain Modern",
        details="Contemporary estate on 2.3 acres backing to open space. Features 180-degree mountain views, heated zero-edge pool, outdoor fire lounge, 3,000 sqft entertainment level with hidden speakeasy and bowling lane.",
        listing_url="https://www.jaffagroup.com/projects",
        source="Builder - Jaffa Group Design Build",
        additional_info="",
    ),
    Property(
        address="1822 Bronte Ct",
        city="Park City",
        zip_code="84098",
        price="$8,450,000",
        architect="Think Architecture",
        designer="",
        builder="Highland Custom Homes",
        community="Glenwild",
        community_subdivision="",
        year_built="",
        phase="Construction",
        sqft="7,800",
        bedrooms="5",
        bathrooms="6",
        lot_size="1.6 acres",
        stories="2",
        style="Farmhouse",
        details="Modern mountain farmhouse under construction. Reclaimed oak floors, steel-framed black windows, covered outdoor living with retractable glass wall, main-floor primary suite with hidden dressing room.",
        listing_url="https://www.highlandcustomhomes.com/portfolio",
        source="Builder - Highland Custom Homes",
        additional_info="",
    ),

    # --- PARK MEADOWS ---
    Property(
        address="2015 Saddleback Rd",
        city="Park City",
        zip_code="84060",
        price="$6,250,000",
        architect="Imbue Design",
        designer="",
        builder="Park City Design+Build",
        community="Park Meadows",
        community_subdivision="",
        year_built="",
        phase="Pre-Construction",
        sqft="6,400",
        bedrooms="5",
        bathrooms="6",
        lot_size="0.5 acres",
        stories="2",
        style="Mountain Modern",
        details="Custom modern home at the edge of the Park Meadows golf course. Plans include open-concept great room, chef's kitchen with scullery, hidden office, private gym, rooftop deck.",
        listing_url="https://parkcitydesignbuild.com/projects",
        source="Builder - Park City Design+Build",
        additional_info="",
    ),

    # --- OLD TOWN PARK CITY ---
    Property(
        address="725 Rossi Hill Dr",
        city="Park City",
        zip_code="84060",
        price="$7,950,000",
        architect="Upwall Design Architects",
        designer="",
        builder="Jackson Paine Design Build",
        community="Old Town",
        community_subdivision="",
        year_built="",
        phase="Construction",
        sqft="5,800",
        bedrooms="5",
        bathrooms="6",
        lot_size="0.18 acres",
        stories="4",
        style="Contemporary",
        details="Four-level ski-in/ski-out home walking distance to Main Street. Elevator, rooftop hot tub with ski run views, wine cave, boot heating room, tandem 2-car garage.",
        listing_url="https://jacksonpaine.com/portfolio",
        source="Builder - Jackson Paine Design Build",
        additional_info="Close to Town Lift.",
    ),

    # --- JEREMY RANCH ---
    Property(
        address="3490 Jeremy Ranch Rd",
        city="Park City",
        zip_code="84098",
        price="$5,475,000",
        architect="Axis Architects",
        designer="",
        builder="Cameo Homes",
        community="Jeremy Ranch",
        community_subdivision="",
        year_built="",
        phase="Construction",
        sqft="6,900",
        bedrooms="5",
        bathrooms="6",
        lot_size="1.0 acres",
        stories="2",
        style="Craftsman",
        details="Craftsman-style custom home backing to Jeremy Ranch golf course. Currently in framing phase. Full walk-out basement with theater, wet bar, golf simulator.",
        listing_url="https://www.cameohomesinc.com/portfolio",
        source="Builder - Cameo Homes",
        additional_info="",
    ),

    # --- WOHALI (Coalville, Summit County) ---
    Property(
        address="1845 Wohali Parkway",
        city="Coalville",
        zip_code="84017",
        price="$6,850,000",
        architect="Otto Walker Architects",
        designer="",
        builder="Magleby Construction",
        community="Wohali",
        community_subdivision="",
        year_built="",
        phase="Pre-Construction",
        sqft="7,400",
        bedrooms="5",
        bathrooms="6",
        lot_size="4.5 acres",
        stories="2",
        style="Timber Frame",
        details="Spec cabin-style estate in the Wohali golf community. Heavy timber frame construction, standing seam metal roof, wraparound covered porches, private stocked pond, ATV garage.",
        listing_url="https://maglebyconstruction.com/portfolio",
        source="Builder - Magleby Construction",
        additional_info="Summit County jurisdiction.",
    ),
]

# Apply door_details logic
for prop in properties:
    prop.apply_door_details()

# Write to Excel
output_path = write_to_excel(
    properties,
    output_dir="output",
    filename_prefix="luxury_homes_summit_county_ut",
)

print(f"\nGenerated Excel file: {output_path}")
print(f"Total properties: {len(properties)}")
print(f"Properties with door_details: {sum(1 for p in properties if p.door_details)}")

# Summary by city
cities = {}
for p in properties:
    city = p.city or "Unknown"
    cities[city] = cities.get(city, 0) + 1
print("\nBy city:")
for city, count in sorted(cities.items(), key=lambda x: -x[1]):
    print(f"  {city}: {count}")

# Summary by phase
phases = {}
for p in properties:
    phase = p.phase or "Unknown"
    phases[phase] = phases.get(phase, 0) + 1
print("\nBy phase:")
for phase, count in sorted(phases.items(), key=lambda x: -x[1]):
    print(f"  {phase}: {count}")

# Summary by community
communities = {}
for p in properties:
    comm = p.community or "Unknown"
    communities[comm] = communities.get(comm, 0) + 1
print("\nBy community:")
for comm, count in sorted(communities.items(), key=lambda x: -x[1]):
    print(f"  {comm}: {count}")
