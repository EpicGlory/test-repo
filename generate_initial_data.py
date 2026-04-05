#!/usr/bin/env python3
"""
Generate initial Excel file with luxury home data gathered from web research.

This script creates the first Excel output using data collected from
web searches of builder sites, listing platforms, and public records
for North Scottsdale and Paradise Valley, AZ luxury homes ($5M+)
that are not yet completed.
"""

import sys
sys.path.insert(0, ".")

from models.property import Property
from utils.excel_writer import write_to_excel

# ============================================================================
# PROPERTIES DATA - Gathered from web research of multiple sources
# ============================================================================

properties = [
    # --- SILVER SKY DEVELOPMENT, PARADISE VALLEY ---
    Property(
        address="6109 E Maverick Rd",
        city="Paradise Valley",
        zip_code="85253",
        price="$20,000,000",
        architect="",
        designer="Jeff Berghoff (Landscape Architect)",
        builder="Silver Sky Development",
        community="Silver Sky",
        community_subdivision="Mummy Mountain",
        year_built="",
        phase="Construction",
        sqft="11,832",
        bedrooms="6",
        bathrooms="10",
        lot_size="1.5 acres",
        stories="",
        style="European Modern",
        details="Nova showcase estate. Best in American Living Award winner. Floor-to-ceiling windows, captivating fusion of stone and stucco, lush gardens. 5-car garage, detached 985 sqft casita, regulation-size pickleball court. Mountain and desert views. Sold mid-construction for $20M - one of priciest off-plan transactions in metro Phoenix.",
        listing_url="https://silverskypv.com/properties/6109-e-maverick-road-paradise-valley-az-us-85253-20250122202724948427000000",
        source="Silver Sky Development; Hoodline; Redfin",
        additional_info="MLS# 6812640. Named 'Nova'. Part of 12-home enclave at base of Mummy Mountain.",
    ),
    Property(
        address="6123 E Ironwood Dr",
        city="Paradise Valley",
        zip_code="85253",
        price="",
        builder="Silver Sky Development",
        community="Silver Sky",
        community_subdivision="Mummy Mountain",
        year_built="",
        phase="Pre-Construction",
        lot_size="1.3 acres",
        details="Lot 11 - one of the last remaining homesites within Silver Sky. 12-home enclave at base of Mummy Mountain. Custom estate parcel.",
        listing_url="https://silverskypv.com/properties",
        source="Silver Sky Development",
        additional_info="One of last remaining lots in Silver Sky.",
    ),
    Property(
        address="6107 E Ironwood Dr",
        city="Paradise Valley",
        zip_code="85253",
        price="",
        builder="Silver Sky Development",
        community="Silver Sky",
        community_subdivision="Mummy Mountain",
        year_built="",
        phase="Pre-Construction",
        lot_size="1.3 acres",
        details="Lot 12 - nearly 1.3-acre estate parcel at the base of iconic Mummy Mountain. One of just twelve custom homesites in Silver Sky.",
        listing_url="https://silverskypv.com/properties",
        source="Silver Sky Development",
        additional_info="Lot 12. Available for custom build.",
    ),
    Property(
        address="Silver Sky - Aster Estate",
        city="Paradise Valley",
        zip_code="85253",
        price="",
        builder="Silver Sky Development",
        community="Silver Sky",
        community_subdivision="Mummy Mountain",
        year_built="",
        phase="Construction",
        details="Aster spec home - one of two new luxury spec homes hitting market in January 2026 inside Silver Sky. Includes detached casita, private pickleball court, and oversized garage designed for car collectors.",
        listing_url="https://silverskypv.com/",
        source="Williams Luxury Homes; Silver Sky Development",
        additional_info="Spec home launched January 2026.",
    ),
    Property(
        address="Silver Sky - Dionysus Estate",
        city="Paradise Valley",
        zip_code="85253",
        price="",
        builder="Silver Sky Development",
        community="Silver Sky",
        community_subdivision="Mummy Mountain",
        year_built="",
        phase="Design",
        details="Dionysus estate within Silver Sky development. Award-winning showcase design on one-acre parcel.",
        listing_url="https://silverskypv.com/developments/dionysus",
        source="Silver Sky Development",
    ),

    # --- MOCKINGBIRD LANE, PARADISE VALLEY ---
    Property(
        address="7301 N Mockingbird Ln",
        city="Paradise Valley",
        zip_code="85253",
        price="",
        architect="John Cates",
        builder="",
        community="Paradise Valley",
        year_built="",
        phase="Construction",
        sqft="9,052",
        bedrooms="5",
        bathrooms="6.5",
        lot_size="1+ acre",
        style="Modern",
        details="Custom estate blending modern scale with natural materials. Designed for grand-scale entertaining and comfortable daily living. Private executive office, elegant bar and lounge. Framed by floor-to-ceiling vistas of Camelback Mountain. Imported stone, custom millwork. Completion expected May 2026.",
        listing_url="https://www.redfin.com/AZ/Paradise-Valley/7301-N-Mockingbird-Ln-85253/home/26991035",
        source="Redfin; Coldwell Banker",
        additional_info="Completion: May 2026.",
    ),
    Property(
        address="7201 N Mockingbird Ln",
        city="Paradise Valley",
        zip_code="85253",
        price="$18,000,888",
        builder="",
        community="Paradise Valley",
        year_built="",
        phase="Construction",
        details="Ultra-luxury new construction estate on Mockingbird Lane in Paradise Valley.",
        listing_url="https://www.redfin.com/AZ/Paradise-Valley/7201-N-Mockingbird-Ln-85253/home/27438851",
        source="Redfin; Compass",
        additional_info="MLS# 6796965.",
    ),
    Property(
        address="7545 N Mockingbird Ln",
        city="Paradise Valley",
        zip_code="85253",
        price="$40,000,000",
        builder="",
        community="Paradise Valley",
        year_built="",
        phase="Construction",
        details="Exceptionally high-end ultra-luxury property on Mockingbird Lane. One of the most expensive new construction listings in Paradise Valley.",
        listing_url="https://www.homes.com/property/7545-n-mockingbird-ln-paradise-valley-az/43dh7z37f0wp7/",
        source="Homes.com",
        additional_info="$40M asking price.",
    ),

    # --- AZURE COMMUNITY, PARADISE VALLEY ---
    Property(
        address="6850 E Joshua Tree Ln",
        city="Paradise Valley",
        zip_code="85253",
        price="",
        builder="Shea Homes",
        community="Azure",
        year_built="",
        phase="Construction",
        details="New-construction home in the Azure gated community. Features walkout basement and guest casita. Coveted corner position within Azure. Views of Camelback and Mummy Mountain.",
        listing_url="https://www.sheahomes.com/new-homes/arizona/phoenix-area/paradise-valley/azure",
        source="KTAR; Shea Homes",
        additional_info="Azure is a luxurious gated community by Shea Homes in Paradise Valley.",
    ),
    Property(
        address="6802 E Joshua Tree Ln",
        city="Paradise Valley",
        zip_code="85253",
        price="$6,000,000",
        builder="Shea Homes",
        community="Azure",
        year_built="",
        phase="Construction",
        sqft="",
        bedrooms="3",
        bathrooms="7",
        details="New construction in Azure community. Views of Camelback and Mummy Mountain. Includes two private casitas. Approximately $6M asking price.",
        listing_url="https://www.sheahomes.com/new-homes/arizona/phoenix-area/paradise-valley/azure",
        source="KTAR; Shea Homes",
    ),

    # --- CALVIS WYANT, PARADISE VALLEY ---
    Property(
        address="6264 E Joshua Tree Ln",
        city="Paradise Valley",
        zip_code="85253",
        price="",
        architect="Calvis Wyant Design Group",
        designer="Calvis Wyant Design Group",
        builder="Calvis Wyant Luxury Homes",
        community="Paradise Valley",
        year_built="",
        phase="Construction",
        sqft="7,202",
        bedrooms="4+1",
        bathrooms="5+1",
        style="Transitional",
        details="Transitional style luxury home in the heart of Paradise Valley. Main house: 4 bed/5 bath, Casita: 1 bed/1 bath. Views of Camelback Mountain. Construction underway by Calvis Wyant Design Group.",
        listing_url="https://www.calviswyant.com/6264-e-joshua-tree-lane-paradise-valley/",
        source="Calvis Wyant",
        additional_info="Since 1986, Calvis Wyant has delivered award-winning residences.",
    ),
    Property(
        address="Calvis Wyant - Mummy Mountain Estate",
        city="Paradise Valley",
        zip_code="85253",
        price="",
        builder="Calvis Wyant Luxury Homes",
        community="Paradise Valley",
        community_subdivision="Mummy Mountain",
        year_built="",
        phase="Construction",
        details="On a secluded tree-lined street on the east slope of Mummy Mountain. Stunning home by Calvis Wyant Design Group underway. Coming soon.",
        listing_url="https://www.calviswyant.com/available-homes/",
        source="Calvis Wyant",
    ),
    Property(
        address="Calvis Wyant - Golf Course Estate",
        city="Scottsdale",
        zip_code="85255",
        price="",
        builder="Calvis Wyant Luxury Homes",
        community="",
        year_built="",
        phase="Construction",
        details="Construction underway on a fantastic lot on the 7th fairway with golf course, sunset and mountain views.",
        listing_url="https://www.calviswyant.com/available-homes/",
        source="Calvis Wyant",
    ),

    # --- E&S BUILDERS, PARADISE VALLEY ---
    Property(
        address="E&S Builders - Paradise Valley Estate",
        city="Paradise Valley",
        zip_code="85253",
        price="",
        architect="John Anthony Drafting & Design",
        designer="TheLifestyledCo",
        builder="E&S Builders",
        community="Paradise Valley",
        year_built="",
        phase="Construction",
        sqft="8,693",
        bedrooms="7",
        bathrooms="8",
        lot_size="1.18 acres",
        details="Brand-new estate on 1.18 acres. 7,482 sqft main residence + 1,211 sqft casita. Designed by John Anthony Drafting & Design, interiors curated by TheLifestyledCo. Blends contemporary luxury with warm organic tones and timeless sophistication. Anticipated completion Spring 2026.",
        listing_url="https://esbuildersaz.com/custom-new-builds/",
        source="E&S Builders; Loving Phoenix Realty",
        additional_info="Completion: Spring 2026.",
    ),

    # --- THOMAS JAMES HOMES, PARADISE VALLEY ---
    Property(
        address="Thomas James Homes - Country Club Acres Estate",
        city="Paradise Valley",
        zip_code="85253",
        price="",
        builder="Thomas James Homes",
        community="Country Club Acres",
        year_built="",
        phase="Construction",
        sqft="8,425",
        bedrooms="6",
        bathrooms="6.5",
        details="Custom new-build estate. 7,200 sqft main home + 1,225 sqft guest house. Single-level. Finishing by Q2 2026.",
        source="Loving Phoenix Realty",
        additional_info="Completion: Q2 2026.",
    ),

    # --- BEDBROCK DEVELOPERS, PARADISE VALLEY ---
    Property(
        address="BedBrock - Ruby Estate, Crown Canyon",
        city="Paradise Valley",
        zip_code="85253",
        price="",
        builder="BedBrock Developers",
        community="Crown Canyon",
        year_built="",
        phase="Construction",
        details="Ruby Estate in Crown Canyon - Arizona's most exclusive gated community. BedBrock broke ground on Ruby Estate. Crown Canyon is a new subdivision in Paradise Valley developed by BedBrock.",
        listing_url="https://bedbrock.com/homes",
        source="BedBrock Developers",
        additional_info="Crown Canyon - ultra-luxury gated community by BedBrock.",
    ),
    Property(
        address="BedBrock - Mummy View Estates",
        city="Paradise Valley",
        zip_code="85253",
        price="",
        builder="BedBrock Developers",
        community="Mummy View Estates",
        year_built="",
        phase="Pre-Construction",
        details="BedBrock's newest ultra-luxury community. Nestled on eight acres with 4 exclusive custom homesites in Paradise Valley.",
        listing_url="https://bedbrock.com/homes",
        source="BedBrock Developers",
        additional_info="4 custom homesites on 8 acres.",
    ),

    # --- SILVERLEAF, SCOTTSDALE ---
    Property(
        address="Silverleaf Upper Canyon - Car Collector Estate",
        city="Scottsdale",
        zip_code="85255",
        price="",
        builder="",
        community="Silverleaf",
        community_subdivision="Upper Canyon",
        year_built="",
        phase="Construction",
        sqft="14,000+",
        lot_size="4 acres",
        details="Car Collector's Dream Home. 14,000+ sqft garage capable of holding 27 cars on the ground or up to 54 with lifts. Panoramic views of McDowell Mountains and Golf Course. 4-acre private estate.",
        source="Redfin; Williams Luxury Homes",
        additional_info="One of the largest private garages in Silverleaf.",
    ),
    Property(
        address="Silverleaf Upper Canyon - Modern Mediterranean",
        city="Scottsdale",
        zip_code="85255",
        price="",
        architect="Scott Carson",
        designer="Kristin Hazen",
        builder="",
        community="Silverleaf",
        community_subdivision="Upper Canyon",
        year_built="",
        phase="Construction",
        sqft="11,619",
        bedrooms="5",
        bathrooms="3.5",
        lot_size="5 acres",
        style="Modern Mediterranean",
        details="Hillside estate in Silverleaf Upper Canyon. Designed by renowned architect Scott Carson, interiors by Kristin Hazen. Five en-suite bedrooms plus three half bathrooms on a five-acre lot.",
        source="Robb Report; Redfin",
        additional_info="Featured in Robb Report.",
    ),
    Property(
        address="Silverleaf Upper Canyon - Bing Hu Estate",
        city="Scottsdale",
        zip_code="85255",
        price="",
        architect="Bing Hu",
        builder="Sommer Custom Homes",
        community="Silverleaf",
        community_subdivision="Upper Canyon",
        year_built="",
        phase="Construction",
        details="Estate two-thirds through construction with mountain and city light views. Designed by architect Bing Hu, crafted by Sommer Custom Homes.",
        source="Scottsdale Market Guide",
    ),

    # --- CULLUM HOMES ---
    Property(
        address="Cullum Homes - Ascent at The Phoenician",
        city="Scottsdale",
        zip_code="85251",
        price="",
        builder="Cullum Homes",
        community="Ascent at The Phoenician",
        year_built="",
        phase="Construction",
        details="Collection of 51 single, two and three-level detached homes adjacent to the world-renowned Phoenician Resort overlooking their golf course. Limited inventory remaining.",
        listing_url="https://cullumhomes.com/available-homes/",
        source="Cullum Homes",
    ),
    Property(
        address="Cullum Homes - Village at Mountain Shadows",
        city="Paradise Valley",
        zip_code="85253",
        price="",
        builder="Cullum Homes",
        community="Village at Mountain Shadows",
        year_built="",
        phase="Construction",
        details="Award-winning development of 40 luxury lifestyle homes in Paradise Valley. Limited inventory remaining.",
        listing_url="https://cullumhomes.com/available-homes/",
        source="Cullum Homes",
    ),
    Property(
        address="Cullum Homes - Village at Seven Desert Mountain",
        city="Scottsdale",
        zip_code="85262",
        price="",
        builder="Cullum Homes",
        community="Desert Mountain",
        community_subdivision="Seven Desert Mountain",
        year_built="",
        phase="Construction",
        details="33 exquisite golf villas on the newest course in Desert Mountain resort community. Lock-and-leave residences backing up to the 13th hole of Renegade Golf Course or 10th/11th holes of Seven Desert Mountain Golf Course.",
        listing_url="https://cullumhomes.com/available-homes/",
        source="Cullum Homes",
    ),

    # --- CAMELOT HOMES ---
    Property(
        address="Camelot Homes - Legacy at DC Ranch",
        city="Scottsdale",
        zip_code="85255",
        price="",
        builder="Camelot Homes",
        community="DC Ranch",
        community_subdivision="Legacy at DC Ranch",
        year_built="",
        phase="Construction",
        sqft="2,700-3,400",
        bedrooms="2-3",
        details="Exclusive enclave of only 8 luxury patio homes. Last parcel of available residential land in DC Ranch. Light-inviting one and two-story floor plans, 2,700-3,400 sqft.",
        listing_url="https://camelothomes.com/community/legacy-at-dc-ranch/",
        source="Camelot Homes; Williams Luxury Homes",
    ),
    Property(
        address="Camelot Homes - Villas II at Seven Desert Mountain",
        city="Scottsdale",
        zip_code="85262",
        price="$3,200,000+",
        builder="Camelot Homes",
        community="Desert Mountain",
        community_subdivision="Villas II at Seven",
        year_built="",
        phase="Construction",
        sqft="",
        bedrooms="2-3",
        bathrooms="3.5",
        details="23 homes in final phase of construction on fairways of Seven golf course. South golf views or North golf views with sunsets and Continental Mountains. Prices start at $3.2M+ plus lot premiums and upgrades.",
        listing_url="https://camelothomes.com/community/the-retreat/",
        source="Camelot Homes; Top Scottsdale Homes",
    ),

    # --- DESERT MOUNTAIN ---
    Property(
        address="37200 N Cave Creek Rd 1018",
        city="Scottsdale",
        zip_code="85262",
        price="",
        builder="Sierra Custom Homes",
        community="Desert Mountain",
        community_subdivision="Eagle Feather",
        year_built="",
        phase="Pre-Construction",
        details="Brand-new home in Eagle Feather at Desert Mountain. Construction expected to start Sept 2025 with completion late 2026. Available Spring 2026.",
        source="Desert Mountain Homes Online",
        additional_info="Completion: Late 2026.",
    ),

    # --- SCOTTSDALE CACTUS CORRIDOR ---
    Property(
        address="Bolack Luxury - Cactus Corridor Estate",
        city="Scottsdale",
        zip_code="85260",
        price="",
        builder="Bolack Luxury",
        community="Cactus Corridor",
        year_built="",
        phase="Construction",
        sqft="6,376",
        bedrooms="6",
        bathrooms="7",
        lot_size="~1 acre",
        style="Spanish Modern",
        details="Spanish Modern new construction estate on nearly one acre within the Cactus Corridor. Scheduled for completion in Q2 2026.",
        source="Arizona Homes 411",
        additional_info="Completion: Q2 2026.",
    ),
    Property(
        address="Thomas James Homes - Cactus Corridor",
        city="Scottsdale",
        zip_code="85260",
        price="",
        builder="Thomas James Homes",
        community="Cactus Corridor",
        year_built="",
        phase="Construction",
        sqft="",
        bedrooms="6",
        bathrooms="7.5",
        style="Santa Barbara",
        details="Santa Barbara-style residence in Cactus Corridor. Scheduled for completion in Summer 2026.",
        source="Arizona Homes 411",
        additional_info="Completion: Summer 2026.",
    ),

    # --- TROON NORTH / NORTH SCOTTSDALE ---
    Property(
        address="Serene at Troon - Contemporary Enclave",
        city="Scottsdale",
        zip_code="85262",
        price="",
        architect="Swayback Architects",
        builder="Sonora West Development",
        community="Troon",
        community_subdivision="Serene",
        year_built="",
        phase="Construction",
        details="New gated enclave of 21 cutting-edge single-level contemporary homes. Open Great Room floor plans. Designed by Swayback Architects, built by Sonora West Development.",
        source="Williams Luxury Homes",
    ),
    Property(
        address="Pinnacle Peak - Worth Development Estate",
        city="Scottsdale",
        zip_code="85262",
        price="",
        architect="J. Cor Architecture",
        builder="Worth Development",
        community="Pinnacle Peak",
        year_built="",
        phase="Construction",
        sqft="",
        lot_size="4.22 acres",
        style="Desert Contemporary",
        details="Exquisite desert contemporary masterpiece by Worth Development and J. Cor Architecture on 4.22-acre parcel with unobstructed vistas of Pinnacle Peak, Troon Mountain, and Tom's Thumb/McDowell's.",
        source="Williams Luxury Homes",
    ),

    # --- STARWOOD / RITZ-CARLTON ---
    Property(
        address="Starwood - Ritz-Carlton Residences",
        city="Paradise Valley",
        zip_code="85253",
        price="",
        architect="Drewett Works",
        designer="Ownby Design",
        builder="Starwood Custom Homes",
        community="Ritz-Carlton Residences",
        year_built="",
        phase="Construction",
        details="Exclusive enclave of opulent custom homes at the Ritz-Carlton Residences in Paradise Valley. Partnership between Starwood Custom Homes, Five Star Development, Drewett Works (architecture), and Ownby Design (interiors).",
        listing_url="https://www.starwoodcustom.com/",
        source="Starwood Custom Homes",
        additional_info="Partnership with Five Star Development.",
    ),

    # --- STRATTON ANDREWS ESTATE ---
    Property(
        address="Stratton Andrews - Modern Estate",
        city="Paradise Valley",
        zip_code="85253",
        price="",
        architect="Stratton Andrews",
        designer="Holly Wright",
        builder="",
        community="Paradise Valley",
        year_built="",
        phase="Design",
        details="Newly constructed estate designed by renowned architect Stratton Andrews with interiors curated by acclaimed designer Holly Wright. Set to debut October 2027.",
        source="Web Search",
        additional_info="Completion: October 2027.",
    ),

    # --- ORANGE TREE GOLF COURSE ---
    Property(
        address="Orange Tree Golf Course - 17th Fairway",
        city="Scottsdale",
        zip_code="85260",
        price="",
        builder="",
        community="Orange Tree",
        year_built="",
        phase="Construction",
        details="Brand new 2026 construction on the 17th fairway of the Championship Orange Tree Golf Course.",
        source="Arizona Homes 411",
        additional_info="On 17th fairway.",
    ),
]

# Apply door_details logic to all properties
for prop in properties:
    prop.apply_door_details()

# Write to Excel
output_path = write_to_excel(
    properties,
    output_dir="output",
    filename_prefix="luxury_homes_scottsdale_pv",
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
