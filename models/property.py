"""Unified data model for luxury home properties."""

from dataclasses import dataclass, field, asdict
from datetime import datetime


@dataclass
class Property:
    """Represents a luxury home property under construction or in planning."""

    address: str = ""
    city: str = ""
    zip_code: str = ""
    price: str = ""
    architect: str = ""
    designer: str = ""
    builder: str = ""
    community: str = ""
    community_subdivision: str = ""
    year_built: str = ""  # Should be empty for incomplete homes
    phase: str = ""  # Design, Pre-Construction, Construction, Framing, etc.
    sqft: str = ""
    bedrooms: str = ""
    bathrooms: str = ""
    lot_size: str = ""
    stories: str = ""
    style: str = ""  # Architectural style (Contemporary, Desert Modern, etc.)
    details: str = ""  # Full description / features
    door_details: str = ""  # "hidden office", "hidden library", or "hidden"
    listing_url: str = ""
    source: str = ""  # Which scraper found it
    date_scraped: str = field(default_factory=lambda: datetime.now().strftime("%Y-%m-%d"))
    additional_info: str = ""

    def to_dict(self) -> dict:
        """Convert to dictionary for DataFrame/Excel export."""
        return asdict(self)

    def apply_door_details(self):
        """Scan details field for hidden room keywords and populate door_details."""
        if not self.details:
            return

        details_lower = self.details.lower()
        found = []

        if "hidden office" in details_lower:
            found.append("hidden office")
        if "hidden library" in details_lower:
            found.append("hidden library")

        # Only add generic "hidden" if neither specific phrase was found
        # but "hidden" appears in the text
        if not found and "hidden" in details_lower:
            found.append("hidden")

        self.door_details = "; ".join(found)

    def merge_with(self, other: "Property", priority: bool = False):
        """Merge another property's data into this one.

        Fills in empty fields from the other property.
        If priority is True, the other property's non-empty fields overwrite ours.
        """
        for fld in self.__dataclass_fields__:
            if fld in ("date_scraped", "source", "listing_url"):
                continue
            other_val = getattr(other, fld, "")
            self_val = getattr(self, fld, "")
            if priority and other_val:
                setattr(self, fld, other_val)
            elif not self_val and other_val:
                setattr(self, fld, other_val)

        # Append sources
        if other.source and other.source not in self.source:
            self.source = f"{self.source}; {other.source}" if self.source else other.source
        if other.listing_url and other.listing_url not in (self.listing_url or ""):
            self.listing_url = (
                f"{self.listing_url} | {other.listing_url}"
                if self.listing_url
                else other.listing_url
            )

    def is_complete_home(self) -> bool:
        """Check if this property appears to be a completed/built home."""
        if self.year_built and self.year_built.strip():
            return True
        phase_lower = (self.phase or "").lower()
        completed_keywords = ["completed", "sold", "closed", "move-in ready", "built"]
        return any(kw in phase_lower for kw in completed_keywords)
