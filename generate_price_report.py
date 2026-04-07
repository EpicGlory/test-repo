"""
PLMA Price Risk, Market Position & Generational Strategy Report Generator
Report #3 of 3 for CFO Ryan Johansen.
Companion to Credit Risk (Report #1) and Technology Strategy (Report #2).
"""
from io import BytesIO

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import Rectangle, FancyBboxPatch
import numpy as np

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# =============================================================================
# CONFIG
# =============================================================================
class ReportConfig:
    NAVY = RGBColor(0x00, 0x3A, 0x70)
    BLUE = RGBColor(0x00, 0x78, 0xD4)
    LIGHT_BLUE = RGBColor(0x00, 0xA4, 0xEF)
    RED = RGBColor(0xE8, 0x11, 0x23)
    AMBER = RGBColor(0xFF, 0xB9, 0x00)
    GREEN = RGBColor(0x10, 0x7C, 0x10)
    PURPLE = RGBColor(0x68, 0x21, 0x7A)
    DARK_GRAY = RGBColor(0x33, 0x33, 0x33)
    MID_GRAY = RGBColor(0x66, 0x66, 0x66)
    LIGHT_GRAY = RGBColor(0xD9, 0xD9, 0xD9)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)

    NAVY_HEX = "#003A70"
    BLUE_HEX = "#0078D4"
    LIGHT_BLUE_HEX = "#00A4EF"
    RED_HEX = "#E81123"
    AMBER_HEX = "#FFB900"
    GREEN_HEX = "#107C10"
    PURPLE_HEX = "#68217A"
    DARK_GRAY_HEX = "#333333"
    MID_GRAY_HEX = "#666666"

    HEADING_FONT = "Calibri"
    BODY_FONT = "Calibri"

    NAVY_SHADE = "003A70"
    LIGHT_BLUE_SHADE = "DEEBF7"
    VERY_LIGHT_GRAY_SHADE = "F2F2F2"
    AMBER_SHADE = "FFF4CE"
    RED_SHADE = "FDE7E9"
    GREEN_SHADE = "DFF6DD"
    PURPLE_SHADE = "E8D5F0"


# =============================================================================
# OXML UTILITIES
# =============================================================================
def set_cell_shading(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_borders = tc_pr.find(qn("w:tcBorders"))
    if tc_borders is None:
        tc_borders = OxmlElement("w:tcBorders")
        tc_pr.append(tc_borders)
    for edge in ("top", "left", "bottom", "right"):
        if edge in kwargs:
            edge_data = kwargs[edge]
            el = tc_borders.find(qn(f"w:{edge}"))
            if el is None:
                el = OxmlElement(f"w:{edge}")
                tc_borders.append(el)
            for k, v in edge_data.items():
                el.set(qn(f"w:{k}"), str(v))


def add_page_number_field(paragraph):
    run = paragraph.add_run()
    for tag, attrs in [
        ("w:fldChar", {qn("w:fldCharType"): "begin"}),
        ("w:instrText", {qn("xml:space"): "preserve"}),
        ("w:fldChar", {qn("w:fldCharType"): "end"}),
    ]:
        el = OxmlElement(tag)
        for k, v in attrs.items():
            el.set(k, v)
        if tag == "w:instrText":
            el.text = "PAGE"
        run._r.append(el)
    run.font.name = ReportConfig.BODY_FONT
    run.font.size = Pt(9)
    run.font.color.rgb = ReportConfig.MID_GRAY


def set_paragraph_spacing(paragraph, before=0, after=6, line=1.15):
    pf = paragraph.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line


# =============================================================================
# FINANCIAL MODEL
# =============================================================================
class PriceMarketModel:
    def __init__(self):
        self.transaction_volume = 3_200_000_000
        self.commission_rate = 0.01875
        self.commission_revenue = self.transaction_volume * self.commission_rate  # $60M
        self.cattle_share = 0.95  # 95% of volume
        self.lamb_share = 0.05   # 5% lamb (immaterial)

        # Company-owned cattle inventory
        self.inventory_value = 20_000_000  # $20M estimated

        # Year 1 investment
        self.year1_costs = {
            "Hedging program build-out": 300_000,
            "Commission structure analysis & implementation": 250_000,
            "Market share defense (digital, geographic)": 800_000,
            "Young Rancher Program launch": 400_000,
            "Revenue diversification (advisory services)": 500_000,
            "Leading indicator dashboard & analytics": 250_000,
            "Change management & training": 300_000,
        }
        self.ongoing_costs = {
            "Hedging operations & advisory": 200_000,
            "Commission structure management": 100_000,
            "Market share programs (ongoing)": 450_000,
            "Young Rancher Program (annual)": 400_000,
            "Advisory services delivery": 350_000,
            "Dashboard & analytics operations": 150_000,
            "Change management (ongoing)": 150_000,
        }
        self.annual_benefits = {
            "Revenue protection (hedging avoided losses)": 1_500_000,
            "Market share defense (prevented erosion)": 1_800_000,
            "Young Rancher Program (new relationships)": 800_000,
            "Advisory service fee revenue": 1_200_000,
            "Commission structure optimization": 700_000,
            "Inventory hedging gain (expected value)": 500_000,
        }
        self.discount_rate = 0.08
        self.horizon_years = 10

    def commission_sensitivity(self):
        """Returns list of (decline%, revenue_loss, pct_of_revenue)."""
        results = []
        for decline in [0.10, 0.20, 0.30, 0.40]:
            cattle_vol = self.transaction_volume * self.cattle_share
            lamb_vol = self.transaction_volume * self.lamb_share
            # Cattle decline; lamb holds (immaterial)
            new_cattle = cattle_vol * (1 - decline)
            new_total = new_cattle + lamb_vol
            new_rev = new_total * self.commission_rate
            loss = self.commission_revenue - new_rev
            pct = loss / self.commission_revenue
            results.append({
                "decline": decline,
                "revenue_loss": loss,
                "pct_of_revenue": pct,
                "remaining_revenue": new_rev,
            })
        return results

    def inventory_mtm(self):
        """Mark-to-market scenarios for company-owned cattle."""
        results = []
        for decline in [0.10, 0.20, 0.30]:
            unhedged_loss = self.inventory_value * decline
            hedged_loss_60 = unhedged_loss * 0.40  # 60% hedged
            hedged_loss_80 = unhedged_loss * 0.20  # 80% hedged
            results.append({
                "decline": decline,
                "unhedged": unhedged_loss,
                "hedged_60": hedged_loss_60,
                "hedged_80": hedged_loss_80,
            })
        return results

    def market_share_erosion(self):
        """5-year erosion scenarios."""
        return [
            {"scenario": "Mild (2%/yr)", "yr5_volume_loss": 0.10,
             "yr5_rev_loss": self.commission_revenue * 0.10},
            {"scenario": "Moderate (4%/yr)", "yr5_volume_loss": 0.20,
             "yr5_rev_loss": self.commission_revenue * 0.20},
            {"scenario": "Severe (6%/yr)", "yr5_volume_loss": 0.30,
             "yr5_rev_loss": self.commission_revenue * 0.30},
        ]

    def total_year1_cost(self):
        return sum(self.year1_costs.values())

    def total_ongoing_cost(self):
        return sum(self.ongoing_costs.values())

    def total_annual_benefit(self):
        return sum(self.annual_benefits.values())

    def cash_flows(self):
        flows = []
        for y in range(1, self.horizon_years + 1):
            if y == 1:
                cost = self.total_year1_cost()
                benefit = self.total_annual_benefit() * 0.40
            elif y == 2:
                cost = self.total_ongoing_cost()
                benefit = self.total_annual_benefit() * 0.80
            else:
                cost = self.total_ongoing_cost()
                benefit = self.total_annual_benefit()
            flows.append({"year": y, "cost": cost, "benefit": benefit, "net": benefit - cost})
        return flows

    def npv(self):
        return sum(f["net"] / ((1 + self.discount_rate) ** f["year"]) for f in self.cash_flows())

    def cumulative_net(self):
        cum = 0
        result = []
        for f in self.cash_flows():
            cum += f["net"]
            result.append({"year": f["year"], "cumulative": cum})
        return result

    def roi_10yr(self):
        total_cost = self.total_year1_cost() + self.total_ongoing_cost() * (self.horizon_years - 1)
        total_benefit = sum(f["benefit"] for f in self.cash_flows())
        return (total_benefit - total_cost) / total_cost

    def payback_months(self):
        cum = -self.total_year1_cost()
        monthly_ongoing = self.total_ongoing_cost() / 12
        y1_mb = (self.total_annual_benefit() * 0.40) / 12
        y2_mb = (self.total_annual_benefit() * 0.80) / 12
        ss_mb = self.total_annual_benefit() / 12
        for m in range(1, 121):
            if m <= 12:
                cum += y1_mb
            elif m <= 24:
                cum += y2_mb - monthly_ongoing
            else:
                cum += ss_mb - monthly_ongoing
            if cum >= 0:
                return m
        return None

    def sensitivity(self):
        base = self.npv()
        results = []
        orig_b = self.annual_benefits.copy()
        orig_y1 = self.year1_costs.copy()
        orig_ong = self.ongoing_costs.copy()
        orig_dr = self.discount_rate

        # Benefit realization +/-30%
        for mult, label in [(0.7, "low"), (1.3, "high")]:
            for k in self.annual_benefits:
                self.annual_benefits[k] = orig_b[k] * mult
            results.append(("Benefit realization", label, self.npv()))
        self.annual_benefits = orig_b.copy()

        # Cattle price scenario
        for mult, label in [(0.6, "low"), (1.2, "high")]:
            self.annual_benefits["Revenue protection (hedging avoided losses)"] = 1_500_000 * mult
            self.annual_benefits["Market share defense (prevented erosion)"] = 1_800_000 * mult
            results.append(("Cattle price severity", label, self.npv()))
        self.annual_benefits = orig_b.copy()

        # Implementation cost
        for mult, label in [(0.8, "low"), (1.4, "high")]:
            for k in self.year1_costs:
                self.year1_costs[k] = orig_y1[k] * mult
            results.append(("Implementation cost", label, self.npv()))
        self.year1_costs = orig_y1.copy()

        # Young rancher success
        for mult, label in [(0.5, "low"), (1.5, "high")]:
            self.annual_benefits["Young Rancher Program (new relationships)"] = 800_000 * mult
            results.append(("Young Rancher success", label, self.npv()))
        self.annual_benefits = orig_b.copy()

        # Discount rate
        for rate, label in [(0.06, "low"), (0.12, "high")]:
            self.discount_rate = rate
            results.append(("Discount rate", label, self.npv()))
        self.discount_rate = orig_dr

        # Market share erosion speed
        for mult, label in [(0.5, "low"), (1.5, "high")]:
            self.annual_benefits["Advisory service fee revenue"] = 1_200_000 * mult
            results.append(("Advisory revenue", label, self.npv()))
        self.annual_benefits = orig_b.copy()

        variables = {}
        for var, lbl, npv_val in results:
            variables.setdefault(var, {})[lbl] = npv_val
        tornado = []
        for var, vals in variables.items():
            tornado.append({
                "variable": var, "low": vals["low"], "high": vals["high"],
                "base": base, "range": abs(vals["high"] - vals["low"]),
            })
        tornado.sort(key=lambda x: x["range"], reverse=True)
        return tornado, base


# =============================================================================
# CHART GENERATOR
# =============================================================================
class PriceChartGenerator:
    def __init__(self, config, model):
        self.config = config
        self.model = model
        plt.rcParams["font.family"] = "DejaVu Sans"
        plt.rcParams["font.size"] = 10
        plt.rcParams["axes.edgecolor"] = "#666666"
        plt.rcParams["axes.linewidth"] = 0.8
        plt.rcParams["axes.spines.top"] = False
        plt.rcParams["axes.spines.right"] = False

    def _finalize(self, fig, dpi=200):
        buf = BytesIO()
        fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        buf.seek(0)
        return buf

    # 1. Cattle Cycle Historical (40 year)
    def cattle_cycle(self):
        years = list(range(1986, 2027))
        prices = [
            62, 67, 73, 78, 82, 76, 72, 74, 69, 63, 60, 65, 60, 63, 68, 72, 71, 80, 85,
            87, 86, 92, 93, 100, 95, 114, 122, 148, 160, 122, 108, 115, 125, 121, 118,
            135, 142, 172, 190, 195, 198
        ]
        fig, ax = plt.subplots(figsize=(13, 6))
        ax.plot(years, prices, color=self.config.NAVY_HEX, linewidth=2.5)
        ax.fill_between(years, 0, prices, color=self.config.NAVY_HEX, alpha=0.15)
        ax.axvspan(1990, 1996, alpha=0.08, color=self.config.RED_HEX)
        ax.axvspan(2014, 2017, alpha=0.08, color=self.config.RED_HEX)
        # Annotations using explicit data coordinates for reliable placement
        _bbox = dict(boxstyle="round,pad=0.3", fc="white", ec="none", alpha=0.9)
        ax.text(1993, 42, "1990-1996\nCycle Trough", fontsize=8.5, fontweight="bold",
                color=self.config.DARK_GRAY_HEX, ha="center", va="center", bbox=_bbox)
        ax.text(2006, 110, "2003-2009\nCycle", fontsize=8.5, fontweight="bold",
                color=self.config.DARK_GRAY_HEX, ha="center", va="center", bbox=_bbox)
        ax.annotate("2014-2017\nPeak-to-Trough\n(-32%)", xy=(2016, 115),
                    xytext=(2008, 160), fontsize=8.5, fontweight="bold",
                    color=self.config.RED_HEX, ha="center", va="center", bbox=_bbox,
                    arrowprops=dict(arrowstyle="->", color=self.config.RED_HEX, lw=1))
        ax.annotate("2025-2026\nALL-TIME HIGH", xy=(2026, 198),
                    xytext=(2022, 170), fontsize=8.5, fontweight="bold",
                    color=self.config.RED_HEX, ha="center", va="center", bbox=_bbox,
                    arrowprops=dict(arrowstyle="->", color=self.config.RED_HEX, lw=1))
        ax.set_ylim(0, 215)
        ax.set_title("U.S. Fed Cattle Prices: 40-Year Cycle Analysis (1986–2026)",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.set_ylabel("Fed Cattle Price ($/cwt)", fontsize=10)
        ax.set_xlabel("Year", fontsize=10)
        ax.grid(True, alpha=0.3)
        ax.set_axisbelow(True)
        ax.text(0.01, -0.14, "Source: USDA AMS, CME Group; prices approximate annual averages for choice steers.",
                transform=ax.transAxes, fontsize=8, color=self.config.MID_GRAY_HEX, style="italic")
        return self._finalize(fig)

    # 2. Fed Cattle Price Scenarios (fan chart)
    def price_scenarios(self):
        years = np.arange(2026, 2032)
        bull = [198, 202, 195, 190, 188, 185]
        base = [198, 185, 168, 155, 150, 148]
        bear = [198, 170, 145, 128, 120, 118]
        fig, ax = plt.subplots(figsize=(10, 5.5))
        ax.fill_between(years, bear, bull, color=self.config.NAVY_HEX, alpha=0.12, label="Uncertainty range")
        ax.plot(years, bull, color=self.config.GREEN_HEX, linewidth=2.5, marker="o",
                markersize=6, label="Bull case: Tight supply persists")
        ax.plot(years, base, color=self.config.AMBER_HEX, linewidth=3, marker="s",
                markersize=7, label="Base case: Orderly correction")
        ax.plot(years, bear, color=self.config.RED_HEX, linewidth=2.5, marker="D",
                markersize=6, label="Bear case: Rapid liquidation")
        for y_vals, color in [(bull, self.config.GREEN_HEX), (base, self.config.AMBER_HEX),
                               (bear, self.config.RED_HEX)]:
            for yr, val in zip(years, y_vals):
                ax.annotate(f"${val}", (yr, val), textcoords="offset points",
                            xytext=(0, 8 if color != self.config.RED_HEX else -15),
                            ha="center", fontsize=8, color=color)
        decline_base = (198 - 148) / 198
        decline_bear = (198 - 118) / 198
        ax.annotate(f"Base: -{decline_base*100:.0f}%", xy=(2031, 148), xytext=(2031.3, 135),
                    fontsize=10, fontweight="bold", color=self.config.AMBER_HEX)
        ax.annotate(f"Bear: -{decline_bear*100:.0f}%", xy=(2031, 118), xytext=(2031.3, 105),
                    fontsize=10, fontweight="bold", color=self.config.RED_HEX)
        ax.set_title("Fed Cattle Price Scenarios (2026–2031)",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.set_ylabel("Price ($/cwt)", fontsize=10)
        ax.set_xlabel("Year", fontsize=10)
        ax.legend(loc="lower left", frameon=False, fontsize=9)
        ax.grid(True, alpha=0.3)
        ax.set_axisbelow(True)
        ax.text(0.01, -0.14, "Source: LMIC, CME futures curves, KSU livestock outlook; analyst projections.",
                transform=ax.transAxes, fontsize=8, color=self.config.MID_GRAY_HEX, style="italic")
        return self._finalize(fig)

    # 3. Commission Revenue Sensitivity Waterfall
    def commission_sensitivity_waterfall(self):
        sens = self.model.commission_sensitivity()
        fig, ax = plt.subplots(figsize=(10, 5.5))
        labels = ["Current\nRevenue"] + [f"{int(s['decline']*100)}% Cattle\nPrice Decline" for s in sens]
        values = [self.model.commission_revenue / 1e6]
        for s in sens:
            values.append(s["remaining_revenue"] / 1e6)
        colors = [self.config.NAVY_HEX] + [self.config.AMBER_HEX if s["decline"] <= 0.2
                  else self.config.RED_HEX for s in sens]
        bars = ax.bar(labels, values, color=colors, edgecolor="white", linewidth=1.5)
        for bar, val in zip(bars, values):
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
                    f"${val:.1f}M", ha="center", fontsize=10, fontweight="bold",
                    color=self.config.DARK_GRAY_HEX)
        # Loss annotations
        for i, s in enumerate(sens):
            ax.annotate(f"-${s['revenue_loss']/1e6:.1f}M\n({s['pct_of_revenue']*100:.0f}%)",
                        xy=(i+1, values[i+1]), xytext=(i+1, values[i+1]-4),
                        ha="center", fontsize=9, color=self.config.RED_HEX, fontweight="bold")
        ax.set_title("Commission Revenue Sensitivity to Cattle Price Decline",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.set_ylabel("Annual Commission Revenue ($M)", fontsize=10)
        ax.grid(True, axis="y", alpha=0.3)
        ax.set_axisbelow(True)
        ax.set_ylim(0, 70)
        return self._finalize(fig)



    # 5. Market Share Erosion Scenarios
    def market_share_erosion(self):
        years = list(range(2026, 2032))
        mild = [100, 98, 96, 94, 92, 90]
        moderate = [100, 96, 92, 88, 84, 80]
        severe = [100, 94, 88, 82, 76, 70]
        fig, ax = plt.subplots(figsize=(10, 5.5))
        ax.plot(years, mild, color=self.config.GREEN_HEX, linewidth=2.5, marker="o",
                markersize=6, label="Mild erosion (2%/yr)")
        ax.plot(years, moderate, color=self.config.AMBER_HEX, linewidth=3, marker="s",
                markersize=7, label="Moderate erosion (4%/yr)")
        ax.plot(years, severe, color=self.config.RED_HEX, linewidth=2.5, marker="D",
                markersize=6, label="Severe erosion (6%/yr)")
        ax.fill_between(years, severe, mild, color=self.config.RED_HEX, alpha=0.08)
        for yr, m, mod, s in zip(years[-1:], mild[-1:], moderate[-1:], severe[-1:]):
            ax.text(yr+0.1, m, f"{m}%", color=self.config.GREEN_HEX, fontsize=10, fontweight="bold")
            ax.text(yr+0.1, mod, f"{mod}%", color=self.config.AMBER_HEX, fontsize=10, fontweight="bold")
            ax.text(yr+0.1, s, f"{s}%", color=self.config.RED_HEX, fontsize=10, fontweight="bold")
        ax.set_title("PLMA Market Share Erosion Scenarios (Indexed, 2026=100)",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.set_ylabel("Market Share Index (2026=100)", fontsize=10)
        ax.set_xlabel("Year", fontsize=10)
        ax.legend(loc="lower left", frameon=False, fontsize=9)
        ax.grid(True, alpha=0.3)
        ax.set_axisbelow(True)
        ax.set_ylim(65, 105)
        return self._finalize(fig)

    # 6. Auction Market Consolidation
    def auction_consolidation(self):
        years = [2000, 2005, 2010, 2015, 2020, 2025]
        auctions = [1200, 1050, 890, 760, 650, 570]
        digital_pct = [0, 2, 5, 10, 18, 32]
        fig, ax1 = plt.subplots(figsize=(9, 5))
        bars = ax1.bar(years, auctions, color=self.config.NAVY_HEX, edgecolor="white",
                       width=3.5, label="Physical auction markets")
        ax1.set_ylabel("Number of Livestock Auction Markets", fontsize=10,
                       color=self.config.NAVY_HEX)
        ax1.tick_params(axis="y", labelcolor=self.config.NAVY_HEX)
        for bar, val in zip(bars, auctions):
            ax1.text(bar.get_x()+bar.get_width()/2, bar.get_height()+15, str(val),
                     ha="center", fontsize=9, fontweight="bold", color=self.config.DARK_GRAY_HEX)
        ax2 = ax1.twinx()
        ax2.plot(years, digital_pct, color=self.config.RED_HEX, linewidth=3, marker="s",
                 markersize=8, label="Digital/video market share (%)")
        ax2.set_ylabel("Digital/Video Market Share (%)", fontsize=10, color=self.config.RED_HEX)
        ax2.tick_params(axis="y", labelcolor=self.config.RED_HEX)
        ax2.spines["top"].set_visible(False)
        ax1.set_title("Livestock Auction Market Consolidation & Digital Shift (2000–2025)",
                       fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        l1, lab1 = ax1.get_legend_handles_labels()
        l2, lab2 = ax2.get_legend_handles_labels()
        ax1.legend(l1+l2, lab1+lab2, loc="upper center", frameon=True, fontsize=9,
                   facecolor="white", edgecolor="none", framealpha=0.95, ncol=2)
        ax1.grid(True, axis="y", alpha=0.3)
        ax1.set_axisbelow(True)
        fig.text(0.1, -0.02, "Source: LMA, USDA AMS market reports; digital share estimated.",
                 fontsize=8, color=self.config.MID_GRAY_HEX, style="italic")
        return self._finalize(fig)

    # 7. Disintermediation Pathways
    def disintermediation(self):
        fig, ax = plt.subplots(figsize=(11, 6))
        ax.set_xlim(0, 10); ax.set_ylim(0, 10)
        ax.axis("off")
        # Boxes
        boxes = [
            (0.3, 7, 2.2, 1.5, "PRODUCER\n(Rancher)", self.config.GREEN_HEX),
            (3.8, 7, 2.2, 1.5, "PLMA\n(Coop Marketer)", self.config.NAVY_HEX),
            (7.3, 7, 2.2, 1.5, "BUYER\n(Feedlot/Packer)", self.config.RED_HEX),
        ]
        for x, y, w, h, text, color in boxes:
            ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.1",
                         facecolor=color, edgecolor="white", linewidth=2))
            ax.text(x+w/2, y+h/2, text, ha="center", va="center", fontsize=11,
                    fontweight="bold", color="white")
        # Traditional arrow
        ax.annotate("", xy=(3.7, 7.75), xytext=(2.6, 7.75),
                    arrowprops=dict(arrowstyle="->", color=self.config.NAVY_HEX, lw=2.5))
        ax.annotate("", xy=(7.2, 7.75), xytext=(6.1, 7.75),
                    arrowprops=dict(arrowstyle="->", color=self.config.NAVY_HEX, lw=2.5))
        ax.text(5, 9, "TRADITIONAL MODEL (PLMA facilitates)", fontsize=12,
                fontweight="bold", color=self.config.NAVY_HEX, ha="center")
        # Disintermediation paths
        paths = [
            (1.4, 6.3, 8.4, 6.3, "Direct-to-packer (captive supply, formula)", self.config.RED_HEX),
            (1.4, 5.3, 8.4, 5.3, "Video/online marketplace (digital entrants)", self.config.AMBER_HEX),
            (1.4, 4.3, 8.4, 4.3, "Cooperative-to-cooperative bypass", self.config.PURPLE_HEX),
            (1.4, 3.3, 8.4, 3.3, "Fintech P2P marketplace", self.config.BLUE_HEX),
        ]
        for x1, y1, x2, y2, label, color in paths:
            ax.annotate("", xy=(x2, y1), xytext=(x1, y1),
                        arrowprops=dict(arrowstyle="->", color=color, lw=2, linestyle="--"))
            ax.text((x1+x2)/2, y1+0.35, label, ha="center", fontsize=9.5,
                    fontweight="bold", color=color)
        ax.text(5, 2.2, "DISINTERMEDIATION PATHWAYS\n(bypass PLMA)", fontsize=12,
                fontweight="bold", color=self.config.RED_HEX, ha="center")
        ax.set_title("Disintermediation Risk: Pathways That Bypass PLMA",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        return self._finalize(fig)

    # 8. Rancher Age Distribution
    def rancher_age_distribution(self):
        ages = ["<35", "35-44", "45-54", "55-64", "65-74", "75+"]
        y2012 = [5.6, 10.2, 21.8, 30.6, 22.1, 9.7]
        y2022 = [4.9, 8.3, 16.1, 27.4, 27.3, 16.0]
        y2035e = [4.2, 7.5, 12.8, 22.5, 29.5, 23.5]
        x = np.arange(len(ages))
        width = 0.25
        fig, ax = plt.subplots(figsize=(10, 5.5))
        ax.bar(x - width, y2012, width, label="2012 Census of Ag",
               color=self.config.BLUE_HEX, edgecolor="white")
        ax.bar(x, y2022, width, label="2022 Census of Ag",
               color=self.config.NAVY_HEX, edgecolor="white")
        ax.bar(x + width, y2035e, width, label="2035 Projected",
               color=self.config.RED_HEX, edgecolor="white")
        ax.set_xticks(x); ax.set_xticklabels(ages, fontsize=10)
        ax.set_ylabel("% of All Farm Operators", fontsize=10)
        ax.set_xlabel("Age Group", fontsize=10)
        ax.set_title("U.S. Farm Operator Age Distribution: The Demographic Cliff",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.legend(loc="upper left", frameon=False, fontsize=9)
        ax.grid(True, axis="y", alpha=0.3)
        ax.set_axisbelow(True)
        ax.annotate("53% of operators\nwill be 65+ by 2035",
                    xy=(4.5, 29.5), xytext=(4.8, 25), fontsize=10, fontweight="bold",
                    color=self.config.RED_HEX,
                    arrowprops=dict(arrowstyle="->", color=self.config.RED_HEX, lw=1.5))
        ax.text(0.01, -0.14, "Source: USDA Census of Agriculture 2012, 2022; 2035 projected using cohort survival model.",
                transform=ax.transAxes, fontsize=8, color=self.config.MID_GRAY_HEX, style="italic")
        return self._finalize(fig)

    # 9. Relationship Lifecycle Funnel
    def relationship_funnel(self):
        stages = ["Prospect\nPool", "First Contact\n(Event/Referral)", "Trial\nTransaction",
                  "Active\nMember", "Multi-Year\nLoyal", "Succession\nTransfer"]
        values = [1000, 400, 200, 120, 85, 25]
        attrition = ["", "60% attrition", "50% attrition", "40% attrition",
                     "29% attrition", "71% lost at\nsuccession"]
        fig, ax = plt.subplots(figsize=(10, 6))
        colors = [self.config.LIGHT_BLUE_HEX, self.config.BLUE_HEX, self.config.NAVY_HEX,
                  self.config.GREEN_HEX, self.config.GREEN_HEX, self.config.RED_HEX]
        max_width = 6.0
        for i, (stage, val, att) in enumerate(zip(stages, values, attrition)):
            w = max_width * val / values[0]
            left = (max_width - w) / 2 + 2
            ax.barh(i, w, left=left, height=0.7, color=colors[i], edgecolor="white",
                    linewidth=1.5)
            # For the last (smallest) bar, put value to the LEFT in red
            if i == len(values) - 1:
                ax.text(left - 0.3, i, f"{val}", ha="right", va="center", fontsize=12,
                        fontweight="bold", color=self.config.RED_HEX)
            else:
                ax.text(5, i, f"{val}", ha="center", va="center", fontsize=12,
                        fontweight="bold", color="white")
            ax.text(left + w + 0.3, i, att, ha="left", va="center", fontsize=9,
                    color=self.config.RED_HEX, fontweight="bold")
        ax.set_yticks(range(len(stages)))
        ax.set_yticklabels(stages, fontsize=10)
        ax.invert_yaxis()
        ax.set_xlim(0, 10)
        ax.set_xticks([])
        ax.spines["bottom"].set_visible(False)
        ax.set_title("Member Relationship Lifecycle Funnel: Where PLMA Loses Producers",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.text(5, 6.3, "Only 2.5% of prospects reach succession transfer — the biggest leak",
                ha="center", fontsize=10, fontweight="bold", color=self.config.RED_HEX)
        return self._finalize(fig)

    # 10. Control / Influence / Monitor matrix
    def control_framework(self):
        fig, ax = plt.subplots(figsize=(11, 7))
        ax.set_xlim(0, 12); ax.set_ylim(0, 10)
        ax.axis("off")
        # Three concentric zones (simplified as stacked boxes)
        zones = [
            (0.3, 0, 3.7, 9.5, "#DFF6DD",
             "CONTROL\n(Direct action)", self.config.GREEN_HEX, [
                 "Hedging program", "Commission structure", "Credit terms & limits",
                 "Digital channel investment", "Young rancher outreach",
                 "Geographic expansion", "Cost structure", "Member services"]),
            (4.2, 0, 3.7, 9.5, "FFF4CE",
             "INFLUENCE\n(Indirect leverage)", self.config.AMBER_HEX, [
                 "Buyer payment behavior", "Industry advocacy (LMA)",
                 "Coop-to-coop partnerships", "Talent pipeline (university)",
                 "Regulatory engagement", "Member estate planning"]),
            (8.1, 0, 3.7, 9.5, "FDE7E9",
             "MONITOR\n(Cannot control)", self.config.RED_HEX, [
                 "Commodity prices", "Federal Reserve policy",
                 "Weather & drought", "Packer consolidation",
                 "Regulatory changes", "Demographic trends",
                 "Trade policy / tariffs", "Disease outbreaks"]),
        ]
        for x, y, w, h, fill, title, tcolor, items in zones:
            color = fill if fill.startswith("#") else f"#{fill}"
            ax.add_patch(Rectangle((x, y), w, h, facecolor=color,
                         edgecolor="white", linewidth=3, alpha=0.85))
            ax.text(x + w/2, h - 1.0, title, ha="center", fontsize=11,
                    fontweight="bold", color=tcolor, va="center")
            for i, item in enumerate(items):
                ax.text(x + 0.3, h - 2.2 - i*0.80, f"• {item}", fontsize=9,
                        color=self.config.DARK_GRAY_HEX)
        ax.set_title("PLMA Control / Influence / Monitor Framework",
                     fontsize=14, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        return self._finalize(fig)

    # 11. Leading Indicator Scorecard
    def leading_indicators(self):
        fig, ax = plt.subplots(figsize=(12, 8))
        ax.set_xlim(0, 14); ax.set_ylim(0, 12)
        ax.axis("off")
        # Category headers
        categories = [
            ("PRICE RISK INDICATORS", 0.3, 11, self.config.NAVY_HEX, [
                ("Cattle-on-Feed (% of year ago)", "98%", "GREEN"),
                ("Placement/Marketing ratio", "1.08", "AMBER"),
                ("Futures curve (contango/backwardation)", "Backwardation", "AMBER"),
                ("Fed cattle basis (vs. 5yr avg)", "+$2.10", "GREEN"),
                ("Packer capacity utilization", "89%", "GREEN"),
            ]),
            ("MARKET SHARE INDICATORS", 0.3, 6.8, self.config.BLUE_HEX, [
                ("YTD transaction volume (vs. plan)", "-3.2%", "AMBER"),
                ("Member retention rate (12m rolling)", "91%", "AMBER"),
                ("New member acquisitions (YTD)", "82", "RED"),
                ("Digital competitor penetration", "32%", "RED"),
            ]),
            ("GENERATIONAL INDICATORS", 0.3, 3.4, self.config.PURPLE_HEX, [
                ("Avg member age", "58.2", "RED"),
                ("Young rancher program enrollment", "45", "AMBER"),
                ("Succession transfers completed (YTD)", "12", "AMBER"),
            ]),
        ]
        status_colors = {"GREEN": self.config.GREEN_HEX, "AMBER": self.config.AMBER_HEX,
                         "RED": self.config.RED_HEX}
        for cat_name, cx, cy, cat_color, indicators in categories:
            # Center header over the indicator rows (names at x=0.6, values at x=7)
            ax.text(4.0, cy, cat_name, fontsize=11, fontweight="bold", color=cat_color,
                    ha="center")
            for i, (name, value, status) in enumerate(indicators):
                row_y = cy - 0.8 - i * 0.65
                ax.scatter(cx + 0.2, row_y, s=120, color=status_colors[status],
                           edgecolor="white", linewidth=2, zorder=3)
                ax.text(cx + 0.6, row_y, name, fontsize=9, va="center",
                        color=self.config.DARK_GRAY_HEX)
                ax.text(7.5, row_y, value, fontsize=9, va="center", fontweight="bold",
                        color=status_colors[status], ha="center")
        ax.set_title("PLMA Leading Indicator Dashboard (Illustrative — April 2026)",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        # Legend — upper right
        for i, (label, color) in enumerate([("On Track", self.config.GREEN_HEX),
                                             ("Watch", self.config.AMBER_HEX),
                                             ("Action Required", self.config.RED_HEX)]):
            ax.scatter(10 + i*1.5, 11.8, s=80, color=color, edgecolor="white", linewidth=2)
            ax.text(10 + i*1.5, 11.35, label, fontsize=7.5, ha="center",
                    color=self.config.DARK_GRAY_HEX)
        return self._finalize(fig)

    # 12. Three Horizons Gantt
    def three_horizons(self):
        workstreams = [
            ("Hedging program (company cattle + overlay)", 0, 9, self.config.NAVY_HEX),
            ("Commission structure reform (analysis + pilot)", 0, 12, self.config.NAVY_HEX),
            ("Leading indicator dashboard", 0, 6, self.config.NAVY_HEX),
            ("Young Rancher Program launch", 3, 12, self.config.NAVY_HEX),
            ("Advisory services (risk mgmt, hedging)", 6, 18, self.config.NAVY_HEX),
            ("Geographic expansion (1-2 new regions)", 12, 36, self.config.BLUE_HEX),
            ("Digital platform defense (w/ Report #2)", 12, 42, self.config.BLUE_HEX),
            ("University/extension partnerships", 6, 36, self.config.BLUE_HEX),
            ("Estate transition liaison service", 12, 48, self.config.BLUE_HEX),
            ("Revenue diversification (advisory scale)", 24, 60, self.config.BLUE_HEX),
            ("Cooperative M&A / partnership", 36, 84, self.config.AMBER_HEX),
            ("Next-gen business model transformation", 48, 108, self.config.AMBER_HEX),
            ("Platform revenue (data, advisory)", 60, 120, self.config.AMBER_HEX),
        ]
        fig, ax = plt.subplots(figsize=(14, 7))
        # Approximate chars that fit per month of bar width at fontsize 9
        chars_per_month = 1.8  # ~1.8 characters fit per month of bar width
        for i, (name, start, end, color) in enumerate(workstreams):
            bar_width = end - start
            ax.barh(i, bar_width, left=start, color=color, edgecolor="white",
                    linewidth=1.5, height=0.7)
            max_chars = int(bar_width * chars_per_month)
            text_color = self.config.NAVY_HEX if color == self.config.AMBER_HEX else "white"
            if len(name) <= max_chars:
                # Text fits inside bar
                ax.text(start + bar_width/2, i, name, ha="center", va="center",
                        fontsize=9, color=text_color, fontweight="bold")
            else:
                # Text too long — place outside to the right in bar color
                ax.text(end + 1, i, name, ha="left", va="center",
                        fontsize=9, color=color, fontweight="bold")
        ax.axvline(x=12, color=self.config.DARK_GRAY_HEX, linestyle="--", alpha=0.5)
        ax.axvline(x=60, color=self.config.DARK_GRAY_HEX, linestyle="--", alpha=0.5)
        ax.text(6, len(workstreams)+0.3, "HORIZON 1\n(Year 1)", ha="center",
                fontsize=10, fontweight="bold", color=self.config.NAVY_HEX)
        ax.text(36, len(workstreams)+0.3, "HORIZON 2\n(Years 2-5)", ha="center",
                fontsize=10, fontweight="bold", color=self.config.BLUE_HEX)
        ax.text(90, len(workstreams)+0.3, "HORIZON 3\n(Years 6-10)", ha="center",
                fontsize=10, fontweight="bold", color=self.config.AMBER_HEX)
        ax.set_yticks([]); ax.set_xlabel("Months from Inception", fontsize=10)
        ax.set_xlim(0, 165); ax.set_ylim(-0.8, len(workstreams)+1.2)
        ax.set_title("Price & Market Position Roadmap — Three Horizons",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=25)
        ax.grid(True, axis="x", alpha=0.3); ax.set_axisbelow(True)
        fig.subplots_adjust(left=0.30, right=0.95)
        return self._finalize(fig)

    # 13. Commission Revenue Under 3 Scenarios (line chart)
    def commission_scenarios_line(self):
        years = list(range(2026, 2032))
        current = [60]*6
        base_decline = [60, 56.5, 52.8, 50.5, 49.2, 48.0]
        bear_decline = [60, 51.2, 44.8, 40.5, 38.8, 37.5]
        mitigated = [60, 55.0, 53.5, 53.0, 54.5, 56.0]
        fig, ax = plt.subplots(figsize=(10, 5.5))
        ax.plot(years, current, color=self.config.MID_GRAY_HEX, linewidth=2, linestyle="--",
                label="No price change (baseline)")
        ax.plot(years, base_decline, color=self.config.AMBER_HEX, linewidth=3, marker="s",
                markersize=7, label="Base case decline (unmitigated)")
        ax.plot(years, bear_decline, color=self.config.RED_HEX, linewidth=2.5, marker="D",
                markersize=6, label="Bear case decline (unmitigated)")
        ax.plot(years, mitigated, color=self.config.GREEN_HEX, linewidth=3, marker="o",
                markersize=7, label="Base case (WITH mitigation)")
        ax.fill_between(years, bear_decline, mitigated, color=self.config.GREEN_HEX, alpha=0.10)
        # Arrow points FROM text (xytext) TO gap midpoint (xy) between mitigated & bear lines
        # Place text in the lower-right area, below all data lines, arrow points UP into the gap
        gap_mid_y = (mitigated[-2] + bear_decline[-2]) / 2  # midpoint of gap at 2030
        ax.annotate("Mitigation\ngap = $8M+/yr\nby Year 5",
                    xy=(2030, gap_mid_y), xytext=(2029.5, 33.5),
                    fontsize=10, fontweight="bold", color=self.config.GREEN_HEX,
                    arrowprops=dict(arrowstyle="->", color=self.config.GREEN_HEX, lw=2,
                                    connectionstyle="arc3,rad=-0.2"),
                    bbox=dict(boxstyle="round,pad=0.4", fc="white", ec=self.config.GREEN_HEX,
                              alpha=0.95, lw=1.0))
        ax.set_title("Commission Revenue Trajectories With & Without Mitigation",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.set_ylabel("Annual Commission Revenue ($M)", fontsize=10)
        ax.set_xlabel("Year", fontsize=10)
        ax.legend(loc="lower left", frameon=False, fontsize=9)
        ax.grid(True, alpha=0.3); ax.set_axisbelow(True)
        ax.set_ylim(30, 65)
        return self._finalize(fig)

    # 14. Cumulative ROI
    def roi_projection(self):
        flows = self.model.cumulative_net()
        years = [f["year"] for f in flows]
        cum = [f["cumulative"] / 1_000_000 for f in flows]
        fig, ax = plt.subplots(figsize=(10, 5.5))
        ax.plot(years, cum, color=self.config.NAVY_HEX, linewidth=3, marker="o", markersize=8)
        ax.fill_between(years, 0, cum, where=[c >= 0 for c in cum],
                        color=self.config.GREEN_HEX, alpha=0.2, label="Net Positive")
        ax.fill_between(years, 0, cum, where=[c < 0 for c in cum],
                        color=self.config.RED_HEX, alpha=0.2, label="Investment Period")
        ax.axhline(y=0, color="black", linewidth=0.8)
        for i in range(len(cum)-1):
            if cum[i] < 0 <= cum[i+1]:
                frac = -cum[i] / (cum[i+1] - cum[i])
                be = years[i] + frac
                ax.axvline(x=be, color=self.config.AMBER_HEX, linestyle="--", linewidth=2)
                ax.text(be+0.1, max(cum)*0.3, f"Breakeven\nYear {be:.1f}", fontsize=10,
                        fontweight="bold", color=self.config.AMBER_HEX)
                break
        for x, y in zip(years, cum):
            ax.annotate(f"${y:.1f}M", (x, y), textcoords="offset points",
                        xytext=(0, 10), ha="center", fontsize=9, color=self.config.DARK_GRAY_HEX)
        ax.set_xlabel("Year", fontsize=11); ax.set_ylabel("Cumulative Net Cash Flow ($M)", fontsize=11)
        ax.set_title("10-Year Cumulative ROI:\nPrice & Market Position Program",
                     fontsize=12, fontweight="bold", color=self.config.NAVY_HEX, pad=18,
                     linespacing=1.3)
        ax.legend(loc="upper left", frameon=True, fontsize=9,
                  facecolor="white", edgecolor="none", framealpha=0.9)
        ax.grid(True, alpha=0.3); ax.set_axisbelow(True)
        return self._finalize(fig)

    # 15. Sensitivity Tornado
    def sensitivity_tornado(self):
        tornado, base = self.model.sensitivity()
        tornado = tornado[:6]
        base_m = base / 1_000_000
        fig, ax = plt.subplots(figsize=(12, 5.5))
        for i, t in enumerate(tornado):
            low_m = t["low"] / 1_000_000 - base_m
            high_m = t["high"] / 1_000_000 - base_m
            ax.barh(i, low_m, color=self.config.RED_HEX, edgecolor="white", height=0.6, alpha=0.85)
            ax.barh(i, high_m, color=self.config.GREEN_HEX, edgecolor="white", height=0.6, alpha=0.85)
        # Render once to establish axes scale before placing labels
        fig.canvas.draw()
        pad_pts = 30  # consistent padding from bar end, in points
        for i, t in enumerate(tornado):
            low_m = t["low"] / 1_000_000 - base_m
            high_m = t["high"] / 1_000_000 - base_m
            # Consistent label placement: always left of red bar, right of green bar
            ax.annotate(f"${t['low']/1e6:.1f}M", xy=(low_m, i),
                        xytext=(-pad_pts, 0), textcoords="offset points",
                        ha="right", va="center", fontsize=9, clip_on=False)
            ax.annotate(f"${t['high']/1e6:.1f}M", xy=(high_m, i),
                        xytext=(pad_pts, 0), textcoords="offset points",
                        ha="left", va="center", fontsize=9, clip_on=False)
        ax.set_yticks(range(len(tornado)))
        ax.set_yticklabels([t["variable"] for t in tornado], fontsize=10)
        ax.invert_yaxis()
        ax.axvline(x=0, color=self.config.NAVY_HEX, linewidth=2)
        ax.margins(x=0.25)
        ax.set_xlabel(f"Change in 10-Year NPV vs. Base (${base_m:.1f}M)", fontsize=10)
        ax.set_title("Sensitivity Tornado — Price & Market Position Program NPV",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.grid(True, axis="x", alpha=0.3); ax.set_axisbelow(True)
        patches = [mpatches.Patch(color=self.config.RED_HEX, label="Downside"),
                   mpatches.Patch(color=self.config.GREEN_HEX, label="Upside")]
        ax.legend(handles=patches, loc="lower right", frameon=False)
        return self._finalize(fig)


# =============================================================================
# DOCUMENT BUILDER
# =============================================================================
class DocumentBuilder:
    def __init__(self, config, charts, model):
        self.config = config
        self.charts = charts
        self.model = model
        self.doc = Document()
        self._setup_styles()
        self._setup_page_layout()

    def _setup_styles(self):
        normal = self.doc.styles["Normal"]
        normal.font.name = self.config.BODY_FONT
        normal.font.size = Pt(10.5)
        normal.font.color.rgb = self.config.DARK_GRAY
        for level, size, color in [(1,20,self.config.NAVY),(2,15,self.config.NAVY),(3,12,self.config.BLUE)]:
            h = self.doc.styles[f"Heading {level}"]
            h.font.name = self.config.HEADING_FONT; h.font.size = Pt(size)
            h.font.color.rgb = color; h.font.bold = True

    def _setup_page_layout(self):
        section = self.doc.sections[0]
        section.top_margin = Inches(0.9); section.bottom_margin = Inches(0.9)
        section.left_margin = Inches(1.0); section.right_margin = Inches(1.0)
        fp = section.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp.add_run("CONFIDENTIAL — Prepared for R. Johansen, CFO, PLMA   |   Page ").font.size = Pt(9)
        for run in fp.runs:
            run.font.name = self.config.BODY_FONT; run.font.color.rgb = self.config.MID_GRAY; run.font.size = Pt(9)
        add_page_number_field(fp)
        hp = section.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run("Price Risk & Market Position  •  April 2026  •  Report 3 of 3")
        hrun.font.name = self.config.BODY_FONT; hrun.font.size = Pt(9)
        hrun.font.color.rgb = self.config.MID_GRAY; hrun.italic = True

    def _para(self, text, size=None, bold=False, italic=False, color=None,
              align=None, space_after=6, space_before=0, line=1.2):
        p = self.doc.add_paragraph()
        run = p.add_run(text)
        run.font.name = self.config.BODY_FONT
        if size: run.font.size = Pt(size)
        run.bold = bold; run.italic = italic
        if color: run.font.color.rgb = color
        if align is not None: p.alignment = align
        set_paragraph_spacing(p, before=space_before, after=space_after, line=line)
        return p

    def _h1(self, text):
        self.doc.add_page_break()
        p = self.doc.add_heading(text, level=1)
        set_paragraph_spacing(p, before=0, after=12)
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single"); bottom.set(qn("w:sz"), "12")
        bottom.set(qn("w:space"), "4"); bottom.set(qn("w:color"), self.config.NAVY_SHADE)
        pBdr.append(bottom); pPr.append(pBdr)
        return p

    def _h2(self, text):
        p = self.doc.add_heading(text, level=2)
        set_paragraph_spacing(p, before=14, after=6); return p

    def _h3(self, text):
        p = self.doc.add_heading(text, level=3)
        set_paragraph_spacing(p, before=10, after=4); return p

    def _body(self, text):
        return self._para(text, size=10.5, space_after=8, line=1.25)

    def _bullet(self, text):
        p = self.doc.add_paragraph(style="List Bullet")
        run = p.add_run(text)
        run.font.name = self.config.BODY_FONT; run.font.size = Pt(10.5)
        run.font.color.rgb = self.config.DARK_GRAY
        set_paragraph_spacing(p, after=3, line=1.2); return p

    def _bullet_lead(self, lead, text):
        p = self.doc.add_paragraph(style="List Bullet")
        r1 = p.add_run(lead + " ")
        r1.bold = True; r1.font.name = self.config.BODY_FONT; r1.font.size = Pt(10.5)
        r1.font.color.rgb = self.config.NAVY
        r2 = p.add_run(text)
        r2.font.name = self.config.BODY_FONT; r2.font.size = Pt(10.5)
        r2.font.color.rgb = self.config.DARK_GRAY
        set_paragraph_spacing(p, after=4, line=1.2); return p

    def _cite(self, paragraph, num):
        r = paragraph.add_run(f" [{num}]")
        r.font.superscript = True; r.font.size = Pt(8); r.font.color.rgb = self.config.BLUE

    def _chart(self, chart_bytes, caption=None, width=6.5):
        self.doc.add_picture(chart_bytes, width=Inches(width))
        self.doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if caption:
            p = self.doc.add_paragraph()
            r = p.add_run(f"Exhibit: {caption}")
            r.font.name = self.config.BODY_FONT; r.font.size = Pt(9); r.italic = True
            r.font.color.rgb = self.config.MID_GRAY
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_paragraph_spacing(p, after=12)

    def _callout(self, title, text, kind="insight"):
        color_map = {
            "insight": (self.config.NAVY_SHADE, self.config.LIGHT_BLUE_SHADE, self.config.NAVY),
            "warning": ("E81123", self.config.RED_SHADE, self.config.RED),
            "finding": ("FFB900", self.config.AMBER_SHADE, self.config.AMBER),
            "sidebar": ("68217A", self.config.PURPLE_SHADE, self.config.PURPLE),
        }
        border_color, fill, heading_color = color_map.get(kind, color_map["insight"])
        table = self.doc.add_table(rows=1, cols=1)
        cell = table.cell(0, 0)
        set_cell_shading(cell, fill)
        set_cell_border(cell,
            left={"val":"single","sz":"36","color":border_color,"space":"0"},
            top={"val":"single","sz":"4","color":border_color,"space":"0"},
            bottom={"val":"single","sz":"4","color":border_color,"space":"0"},
            right={"val":"single","sz":"4","color":border_color,"space":"0"})
        cell.width = Inches(6.5)
        cell.paragraphs[0].text = ""
        p1 = cell.paragraphs[0]
        r1 = p1.add_run(title.upper()); r1.bold = True
        r1.font.name = self.config.HEADING_FONT; r1.font.size = Pt(9)
        r1.font.color.rgb = heading_color
        set_paragraph_spacing(p1, after=3)
        p2 = cell.add_paragraph()
        r2 = p2.add_run(text)
        r2.font.name = self.config.BODY_FONT; r2.font.size = Pt(10)
        r2.font.color.rgb = self.config.DARK_GRAY
        set_paragraph_spacing(p2, after=4, line=1.2)
        self._para("", space_after=8)

    def _table(self, headers, rows, col_widths=None, header_fill=None):
        header_fill = header_fill or self.config.NAVY_SHADE
        table = self.doc.add_table(rows=1+len(rows), cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, h in enumerate(headers):
            cell = table.rows[0].cells[i]
            set_cell_shading(cell, header_fill)
            cell.paragraphs[0].text = ""
            r = cell.paragraphs[0].add_run(h); r.bold = True
            r.font.name = self.config.HEADING_FONT; r.font.size = Pt(9.5)
            r.font.color.rgb = self.config.WHITE
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for ri, row in enumerate(rows):
            for ci, val in enumerate(row):
                cell = table.rows[ri+1].cells[ci]
                if ri % 2 == 1: set_cell_shading(cell, self.config.VERY_LIGHT_GRAY_SHADE)
                cell.paragraphs[0].text = ""
                r = cell.paragraphs[0].add_run(str(val))
                r.font.name = self.config.BODY_FONT; r.font.size = Pt(9.5)
                r.font.color.rgb = self.config.DARK_GRAY
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        if col_widths:
            for row in table.rows:
                for i, w in enumerate(col_widths):
                    row.cells[i].width = Inches(w)
        self._para("", space_after=8)
        return table

    # ============ COVER + TOC ============
    def _add_cover_page(self):
        for _ in range(4): self._para("", space_after=2)
        p = self.doc.add_paragraph()
        r = p.add_run("CONFIDENTIAL  •  CLIENT DELIVERABLE  •  REPORT 3 OF 3")
        r.font.name = self.config.HEADING_FONT; r.font.size = Pt(10)
        r.font.color.rgb = self.config.RED; r.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER; set_paragraph_spacing(p, after=40)
        p = self.doc.add_paragraph()
        r = p.add_run("Price Risk, Market Position\n& Generational Strategy")
        r.font.name = self.config.HEADING_FONT; r.font.size = Pt(34)
        r.font.color.rgb = self.config.NAVY; r.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER; set_paragraph_spacing(p, after=6)
        p = self.doc.add_paragraph()
        r = p.add_run("Navigating the Cattle Cycle Peak, Defending Market Share,\nand Securing the Next Generation")
        r.font.name = self.config.HEADING_FONT; r.font.size = Pt(18)
        r.font.color.rgb = self.config.BLUE
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER; set_paragraph_spacing(p, after=40)
        tbl = self.doc.add_table(rows=1, cols=1)
        cell = tbl.cell(0, 0); set_cell_shading(cell, self.config.NAVY_SHADE)
        cell.paragraphs[0].text = ""; cell.paragraphs[0].add_run(" ").font.size = Pt(2)
        self._para("", space_after=25)
        p = self.doc.add_paragraph()
        r = p.add_run("Producers Livestock Marketing Association")
        r.font.name = self.config.HEADING_FONT; r.font.size = Pt(16)
        r.font.color.rgb = self.config.DARK_GRAY; r.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER; set_paragraph_spacing(p, after=4)
        p = self.doc.add_paragraph()
        r = p.add_run("Prepared for Ryan Johansen, Chief Financial Officer")
        r.font.name = self.config.BODY_FONT; r.font.size = Pt(12)
        r.font.color.rgb = self.config.MID_GRAY; r.italic = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER; set_paragraph_spacing(p, after=15)
        p = self.doc.add_paragraph()
        r = p.add_run("Companion to Report 1 (Counterparty Credit Risk) and Report 2 (Technology Strategy)")
        r.font.name = self.config.BODY_FONT; r.font.size = Pt(10)
        r.font.color.rgb = self.config.PURPLE; r.italic = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER; set_paragraph_spacing(p, after=50)
        p = self.doc.add_paragraph()
        r = p.add_run("APRIL 2026"); r.font.name = self.config.HEADING_FONT
        r.font.size = Pt(11); r.font.color.rgb = self.config.NAVY; r.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def _add_toc(self):
        self.doc.add_page_break()
        self.doc.add_heading("Table of Contents", level=1)
        entries = [
            ("1. Executive Summary", "3"), ("2. Price Risk — Cattle", "6"),
            ("3. Market Share Risk Assessment", "13"),
            ("4. Generational Succession Deep Dive", "20"),
            ("5. Controllable vs. Uncontrollable Framework", "27"),
            ("6. Leading Indicators Dashboard", "30"),
            ("7. Mitigation Strategy & Implementation Roadmap", "34"),
            ("8. Financial Model", "40"), ("9. Conclusion & Next Steps", "45"),
            ("Appendix A: Cattle Cycle Data Tables", "47"),
            ("Appendix B: Commission Sensitivity Detail", "49"),
            ("Appendix C: Generational Demographics", "51"),
            ("Appendix D: Methodology & Assumptions", "53"),
            ("Appendix E: Glossary", "55"), ("Appendix F: Citations", "57"),
        ]
        for title, page in entries:
            p = self.doc.add_paragraph()
            p.paragraph_format.tab_stops.add_tab_stop(Inches(6.3), alignment=2, leader=1)
            r1 = p.add_run(title); r1.font.name = self.config.BODY_FONT; r1.font.size = Pt(11)
            r1.font.color.rgb = self.config.DARK_GRAY
            r2 = p.add_run(f"\t{page}"); r2.font.name = self.config.BODY_FONT
            r2.font.size = Pt(11); r2.font.color.rgb = self.config.NAVY; r2.bold = True
            set_paragraph_spacing(p, after=8, line=1.3)

    # ============ SECTION 1: EXECUTIVE SUMMARY ============
    def _add_executive_summary(self):
        self._h1("1. Executive Summary")
        self._callout("Situation",
            "U.S. cattle prices are at all-time nominal highs — fed cattle averaging ~$198/cwt in Q1 2026. "
            "The cyclical nature of the cattle market makes a correction not a question of IF but WHEN and "
            "HOW MUCH. PLMA's commission revenue — tied directly to livestock transaction value — faces "
            "downside exposure of $5.7M to $17.1M annually under a 10-30% price correction. Simultaneously, "
            "PLMA's market share is under attack from four directions: digital marketplace disruptors, "
            "industry consolidation, direct-to-packer disintermediation, and regional competitive shifts. "
            "And the demographic cliff — average rancher age exceeding 58 — means PLMA's existing "
            "relationship base is literally aging out of production.", kind="warning")
        self._h2("1.1 Key Findings")
        self._bullet_lead("Cattle prices will correct — the only variables are timing and magnitude.",
            "Every cattle cycle in the past 40 years has produced a peak-to-trough correction of 20-40%. "
            "The 2014-2017 cycle saw a 32% decline. PLMA must plan for a 20-30% correction commencing "
            "in the 2026-2028 window.")
        self._bullet_lead("Commission revenue is structurally exposed to price declines.",
            "As a percentage-of-sale marketer, PLMA's revenue declines proportionally with cattle prices. "
            "A 20% cattle price decline erodes ~$11.4M (19%) of annual commission revenue. "
            "This is not a risk PLMA can fully eliminate, but it CAN be substantially mitigated.")
        self._bullet_lead("Company-owned cattle inventory creates additional balance sheet exposure.",
            "An estimated $20M in company-owned inventory faces $4M mark-to-market loss in a 20% decline "
            "scenario. A 60-80% hedging program reduces this to $0.8-1.6M.")
        self._bullet_lead("Market share erosion is the slow-moving existential threat.",
            "Digital marketplaces now account for ~32% of livestock transactions. Auction market count "
            "has declined from 1,200 (2000) to 570 (2025). In a moderate erosion scenario, PLMA loses "
            "4%/year of transaction volume — $12M/year in revenue by Year 5.")
        self._bullet_lead("The generational succession gap is PLMA's most underappreciated risk.",
            "53% of U.S. farm operators will be 65+ by 2035. When ranches transfer through estates, "
            "PLMA loses ~71% of those relationships. Without a formal Young Rancher Program, "
            "PLMA's member base contracts by natural attrition faster than it can be replenished.")
        self._bullet_lead("Controllable actions can substantially narrow the risk exposure.",
            "While PLMA cannot control commodity prices or demographics, it CAN control its hedging "
            "program, commission structure, digital investment, young rancher outreach, and revenue "
            "diversification — and these actions generate ~$26M in 10-year NPV.")
        self._h2("1.2 Financial Impact Summary")
        self._table(["Metric", "Value"], [
            ["Commission revenue at risk (20% cattle decline)", "$11.4M/year"],
            ["Inventory mark-to-market (20% decline, unhedged)", "$4.0M"],
            ["Market share erosion cost (moderate, Year 5)", "$12.0M/year"],
            ["Year 1 Mitigation Investment", "$2.80M"],
            ["Annual Ongoing Cost (Year 2+)", "$1.80M"],
            ["Annual Benefit (Year 2+, steady state)", "$6.50M"],
            ["10-Year NPV (8% discount rate)", "~$26M"],
            ["Payback Period", "~13 months"],
        ], col_widths=[3.5, 2.5])

    # ============ SECTION 2: CATTLE PRICE RISK ============
    def _add_section_2_cattle(self):
        self._h1("2. Price Risk — Cattle")
        self._body("Cattle price risk is PLMA's most immediate financial exposure. This section "
            "analyzes the cattle cycle, models scenarios, quantifies commission sensitivity, and "
            "examines inventory exposure.")
        self._h2("2.1 Cattle Cycle Analysis")
        p = self._body("The U.S. cattle cycle — driven by the biology of cow-calf reproduction and "
            "the economics of heifer retention — operates on roughly 10-12 year cycles. The current "
            "cycle began its liquidation phase ~2019, with the U.S. cow herd declining to 88.2 million "
            "head (January 2026) — the lowest since 1951. This tight supply drove prices to nominal "
            "all-time highs of ~$198/cwt for choice steers in Q1 2026.")
        self._cite(p, 1)
        self._chart(self.charts.cattle_cycle(),
                    caption="40-year cattle price history shows cyclical peaks followed by 20-40% corrections.")
        p = self._body("Historical parallels are instructive. The 2014 peak (~$165/cwt) was followed "
            "by a 32% correction to ~$108/cwt by 2017. The 1990-1996 downturn saw a 27% decline. "
            "In both cases, the correction was triggered by herd rebuilding (heifer retention reduces "
            "fed cattle supply initially but increases it 2-3 years later) combined with macroeconomic "
            "headwinds.")
        self._cite(p, 3)
        self._h2("2.2 Fed Cattle Price Scenarios")
        self._chart(self.charts.price_scenarios(),
                    caption="Base case projects a 25% correction to ~$148/cwt by 2031; bear case implies 40% decline to $118/cwt.")
        self._table(["Scenario", "2026", "2028", "2031", "Peak-to-Trough"],
            [["Bull (tight supply persists)", "$198", "$195", "$185", "-7%"],
             ["Base (orderly correction)", "$198", "$168", "$148", "-25%"],
             ["Bear (rapid liquidation)", "$198", "$145", "$118", "-40%"]], col_widths=[2.0,1.0,1.0,1.0,1.3])
        self._h2("2.3 Feeder Cattle Dynamics")
        self._body("Feeder cattle prices — the primary cash market in which many PLMA transactions "
            "occur — are even more elevated than fed cattle, with 750-lb steers exceeding $280/cwt "
            "in Q1 2026. The cost-of-gain squeeze (record feeder prices vs. volatile feed costs) "
            "creates compounding risk: feedlot losses drive reduced placement, which eventually "
            "breaks feeder prices more sharply than fed cattle.")
        self._h2("2.4 Commission Revenue Sensitivity")
        self._body("PLMA's commission revenue is directly proportional to the dollar value of "
            "livestock sold. At an estimated 1.875% average commission rate on $3.2B in annual "
            "transaction volume, PLMA generates ~$60M in commission revenue. Cattle transactions "
            "represent ~95% of this volume.")
        self._chart(self.charts.commission_sensitivity_waterfall(),
                    caption="A 20% cattle price decline erodes $11.4M (19%) of annual commission revenue.")
        sens = self.model.commission_sensitivity()
        sens_rows = [[f"{int(s['decline']*100)}%", f"${s['revenue_loss']/1e6:.1f}M",
                      f"{s['pct_of_revenue']*100:.1f}%", f"${s['remaining_revenue']/1e6:.1f}M"]
                     for s in sens]
        self._table(["Cattle Price Decline", "Revenue Loss", "% of Total Revenue", "Remaining Revenue"],
                    sens_rows, col_widths=[1.5, 1.3, 1.5, 1.6])
        self._h2("2.5 Company-Owned Cattle Inventory Exposure")
        self._body("PLMA holds an estimated $20M in company-owned cattle inventory at any given time. "
            "This inventory is subject to mark-to-market risk if cattle prices decline during the "
            "holding period.")
        mtm = self.model.inventory_mtm()
        mtm_rows = [[f"{int(m['decline']*100)}%", f"${m['unhedged']/1e6:.1f}M",
                      f"${m['hedged_60']/1e6:.1f}M", f"${m['hedged_80']/1e6:.2f}M"]
                     for m in mtm]
        self._table(["Price Decline", "Unhedged Loss", "60% Hedged", "80% Hedged"],
                    mtm_rows, col_widths=[1.3, 1.5, 1.3, 1.3])
        self._callout("Key Insight",
            "At 80% hedge coverage, a 20% price decline produces a manageable $0.80M inventory loss "
            "vs. $4.0M unhedged. The annual cost of the hedging program (~$200K including basis risk) "
            "is a fraction of the avoided tail loss. This is the highest-ROI action item in this report.",
            kind="finding")
        self._h2("2.6 Basis Risk")
        self._body("Basis — the difference between PLMA's local cash market and CME futures — adds "
            "3-8% volatility on top of national price movements. In PLMA's Plains/Mountain West "
            "footprint, basis tends to widen during periods of tight packer capacity (positive for "
            "sellers) but can compress sharply during seasonal lows or regional weather events.")

    # ============ SECTION 3: MARKET SHARE ============
    def _add_section_3_market_share(self):
        self._h1("3. Market Share Risk Assessment")
        self._body("Market share erosion is a slow, compounding threat that is more dangerous than "
            "any single price cycle. A cooperative that loses 4% of volume per year for 5 years "
            "has permanently impaired its cost structure, bargaining position, and relevance.")
        self._h2("3.1 PLMA's Market Position")
        p = self._body("PLMA is a member-owned, non-stock cooperative established in 1935 and "
            "headquartered in Omaha, Nebraska, with offices in Sioux City (IA), North Salt Lake (UT), "
            "and Madera (CA). Operating three divisions — Marketing, Commodities (Producers Commodities "
            "LLC), and Credit (Producers Livestock Credit Corporation) — PLMA is an industry leader in "
            "livestock marketing throughout the Western United States, with an estimated 4-6% share "
            "of total U.S. livestock marketing transaction volume (~$60B total). The cooperative model "
            "has been the industry standard for over a century. But the standard is under attack.")
        self._cite(p, 17)
        self._chart(self.charts.market_share_erosion(),
                    caption="Under moderate erosion (4%/yr), PLMA loses 20% of volume by 2031 — $12M/yr in commission revenue.")
        self._h2("3.2 Digital Marketplace Disruption")
        self._callout("Connection to Report #2",
            "Report #2 (Emerging Technology Strategy) recommended a digital auction platform as an "
            "INVEST NOW priority (Section 5.5). The market share defense economics in this section "
            "are incremental to — and do not duplicate — the technology platform investment in Report #2. "
            "The $800K market share defense budget here covers geographic expansion, member retention "
            "programs, and value-added services — not the digital platform itself.",
            kind="sidebar")
        self._body("Digital and video livestock sales have grown from near-zero in 2000 to an estimated "
            "32% of all cattle transactions in 2025. Superior Livestock Auction (the largest video "
            "marketer) handles over $5B in annual cattle sales. Newer entrants like Cattlerange, "
            "DVAuction, and various regional online platforms are accelerating share capture.")
        self._chart(self.charts.auction_consolidation(),
                    caption="Physical auctions declined 52% (1,200 → 570) since 2000 while digital share grew from 0% to 32%.")
        self._h2("3.3 Industry Consolidation")
        self._body("The livestock marketing industry is consolidating at both the auction level "
            "(closures) and the cooperative level (mergers). Smaller regional cooperatives lack scale "
            "to invest in technology, compliance, and member services. This creates both a threat "
            "(larger competitors absorbing PLMA's members) and an opportunity (PLMA absorbing smaller "
            "cooperatives).")
        self._h2("3.4 Disintermediation")
        self._chart(self.charts.disintermediation(),
                    caption="Four pathways bypass PLMA: direct-to-packer, digital marketplace, coop-to-coop, and fintech P2P.")
        self._body("Disintermediation — producers selling directly to packers or feedlots, bypassing "
            "the cooperative marketer — is driven by: (i) captive supply arrangements (packers "
            "contracting directly with large feedlots), (ii) formula pricing that removes the need "
            "for auction price discovery, (iii) producer-owned digital platforms, and (iv) "
            "relationship consolidation as the industry concentrates.")
        self._h2("3.5 Regional Competition")
        self._body("Cattle production geography is shifting: drought has accelerated herd liquidation "
            "in the Southern Plains, while the Northern Plains and Corn Belt have gained share. "
            "Feedlot concentration in the Texas Panhandle, SW Kansas, and NE Colorado/SW Nebraska "
            "creates regional pricing power dynamics that affect PLMA's footprint differently "
            "depending on its geographic exposure.")

    # ============ SECTION 4: GENERATIONAL SUCCESSION ============
    def _add_section_4_succession(self):
        self._h1("4. Generational Succession Deep Dive")
        self._body("This section addresses the single most underappreciated strategic risk facing "
            "PLMA: the aging of its producer member base and the failure to systematically capture "
            "the next generation of livestock producers.")
        self._h2("4.1 The Demographic Crisis")
        p = self._body("The average age of U.S. farm operators reached 58.1 years in the 2022 Census "
            "of Agriculture, up from 56.3 in 2012. More alarmingly, the <35 age cohort has shrunk "
            "from 5.6% to 4.9% of all operators, while the 75+ cohort grew from 9.7% to 16.0%. "
            "By 2035, an estimated 53% of all farm operators will be 65 or older.")
        self._cite(p, 14)
        self._chart(self.charts.rancher_age_distribution(),
                    caption="The demographic cliff: 53% of operators projected 65+ by 2035; <35 cohort shrinking.")
        self._h2("4.2 Estate Planning & Relationship Transfer Economics")
        self._body("When a rancher passes away or exits the business, the cooperative relationship "
            "does not automatically transfer to the heir or buyer. Industry data suggests PLMA "
            "retains only ~29% of relationships through generational transfer — losing 71% to "
            "inaction, competitor relationships, or ranch liquidation.")
        self._callout("The Succession Math",
            "If PLMA's member base has an average age of 58, roughly 5-7% of members will exit per "
            "year through retirement, death, or estate transfer within 10 years. At ~4,000 active "
            "members, that is 200-280 exits/year. At 71% relationship loss rate, PLMA loses "
            "140-200 members/year — each worth $15-25K in lifetime commission value. Annual "
            "relationship attrition cost: $2.1-5.0M in future revenue erosion.", kind="finding")
        self._chart(self.charts.relationship_funnel(),
                    caption="PLMA's relationship funnel reveals the largest leak at succession transfer (71% attrition).")
        self._h2("4.3 The 'Lost Generation' Gap")
        self._body("The 35-50 age cohort — the generation that should be PLMA's emerging core "
            "membership — is underrepresented. These producers grew up digital, expect mobile-first "
            "experiences, make data-driven decisions, and are less loyal to 'the way it's always "
            "been done.' They are also the most likely to use direct-sale or digital auction "
            "platforms. If PLMA does not actively engage this cohort, it cedes them to competitors.")
        self._h2("4.4 Young Rancher Acquisition Strategy")
        self._body("Next-generation producers want different things from a cooperative:")
        self._bullet_lead("Digital experience:", "Mobile app, online auctions, real-time market data, digital documentation.")
        self._bullet_lead("Advisory services:", "Risk management education, hedging advisory, credit access facilitation, financial literacy.")
        self._bullet_lead("Flexibility:", "Lower minimum volumes, flexible commission structures, trial memberships.")
        self._bullet_lead("Community:", "Peer networks, mentorship connections, industry events, social media presence.")
        self._bullet_lead("Values alignment:", "Sustainability practices, animal welfare, cooperative governance voice.")
        self._h2("4.5 Mentorship & Ambassador Programs")
        self._body("The most effective young rancher acquisition channel is referral from an existing "
            "trusted member. A formal Mentorship/Ambassador Program pairs experienced PLMA members "
            "with beginning producers in their region. Ambassadors receive modest incentives "
            "(reduced commission tiers, recognition) for successful referrals that convert to "
            "active members.")
        self._h2("4.6 Financial Products for Beginning Producers")
        self._body("Beginning ranchers face acute capital constraints: limited equity, thin credit "
            "history, high land/cattle costs. PLMA can facilitate (not originate) access to:")
        self._bullet("FSA Beginning Farmer and Rancher loans (guaranteed lending programs)")
        self._bullet("Farm Credit System programs designed for beginning operations")
        self._bullet("Risk management education (hedging basics, crop insurance for feed)")
        self._bullet("Cooperative credit accommodation (extended terms for first-year members)")
        self._h2("4.7 Land-Grant University & Extension Partnerships")
        p = self._body("Partnerships with land-grant universities (KSU, TAMU, UNL, SDSU, etc.) and "
            "cooperative extension services create a pipeline of young producers who encounter PLMA "
            "as part of their education — at livestock judging events, extension workshops, "
            "internship programs, and student cooperative competitions.")
        self._cite(p, 26)

    # ============ SECTION 5: CONTROLLABLE vs UNCONTROLLABLE ============
    def _add_section_5_control(self):
        self._h1("5. Controllable vs. Uncontrollable Framework")
        self._body("A fundamental principle of risk management — adapted here from Covey's Circle of "
            "Influence model for enterprise strategy — is distinguishing between factors PLMA can "
            "control (direct action), influence (indirect leverage), and only monitor (cannot change).")
        self._chart(self.charts.control_framework(),
                    caption="PLMA's strategy should focus energy on the Control zone, invest in the Influence zone, and build monitoring for the rest.")
        self._h2("5.1 What PLMA Cannot Control")
        self._body("These factors shape the operating environment but are beyond PLMA's direct action:")
        self._bullet_lead("Commodity prices:", "Cattle prices are set by global supply/demand. PLMA can hedge exposure but cannot influence the price level.")
        self._bullet_lead("Federal Reserve policy:", "Interest rates affect buyer leverage, working capital cost, and farmland values. Monitor and plan; cannot change.")
        self._bullet_lead("Weather and drought:", "Regional weather patterns drive herd liquidation, basis volatility, and production migration. Monitor and diversify geographically.")
        self._bullet_lead("Packer consolidation:", "Big-4 concentration is a regulatory and antitrust issue. PLMA can advocate but cannot reverse market structure.")
        self._bullet_lead("Demographic trends:", "The aging of the rancher population is a national phenomenon. PLMA can capture the next generation but cannot reverse the trend.")
        self._h2("5.2 What PLMA Can Control")
        self._body("These are the high-leverage actions that form the core of the mitigation strategy:")
        self._bullet_lead("Hedging program:", "Direct action to hedge company-owned cattle and overlay member risk advisory.")
        self._bullet_lead("Commission structure:", "Reform from pure %-of-sale to blended model (floor/ceiling, fee-for-service, advisory tiers).")
        self._bullet_lead("Credit terms:", "Tiered limits and payment acceleration (cross-reference Report #1).")
        self._bullet_lead("Digital investment:", "Build digital auction capability, mobile experience, data platform (cross-reference Report #2).")
        self._bullet_lead("Young rancher outreach:", "Formal program, university partnerships, mentorship network.")
        self._bullet_lead("Geographic expansion:", "Enter underserved regions to diversify volume concentration.")
        self._bullet_lead("Cost structure:", "Operating leverage — ensure fixed costs are right-sized for a 20% volume decline scenario.")
        self._h2("5.3 What PLMA Can Influence")
        self._body("These factors are not fully controllable but respond to sustained effort:")
        self._bullet_lead("Buyer payment behavior:", "Credit terms, incentives, and monitoring improve payment velocity.")
        self._bullet_lead("Industry advocacy:", "Through LMA and NCBA, PLMA can shape P&SA enforcement, market transparency rules, and beginning farmer programs.")
        self._bullet_lead("Cooperative partnerships:", "Joint ventures, federated structures, and mutual aid agreements with peer cooperatives reduce isolation.")
        self._bullet_lead("Talent pipeline:", "University partnerships and employer brand investment attract talent over time.")

    # ============ SECTION 6: LEADING INDICATORS ============
    def _add_section_6_indicators(self):
        self._h1("6. Leading Indicators Dashboard")
        self._body("Effective risk management requires systematic monitoring of leading indicators — "
            "not just lagging outcomes. This section defines the key metrics PLMA should track "
            "weekly/monthly across price, market share, and generational dimensions.")
        self._chart(self.charts.leading_indicators(),
                    caption="Illustrative dashboard showing current status. Amber/red indicators require escalation per governance protocols.")
        self._h2("6.1 Price Risk Indicators")
        self._table(["Indicator", "Frequency", "Signal", "Action Trigger"], [
            ["Cattle-on-Feed report", "Monthly", "YoY placement trends", ">105% of year-ago triggers review"],
            ["Placement/Marketing ratio", "Monthly", "Supply pipeline buildup", ">1.10 for 3 consecutive months"],
            ["CME futures curve structure", "Weekly", "Contango vs. backwardation", "Shift to contango = bearish signal"],
            ["Feeder cattle basis", "Weekly", "Regional price dislocation", ">$5 deviation from 5yr avg"],
            ["Packer capacity utilization", "Monthly", "Processing bottleneck", "<85% signals margin pressure"],
        ], col_widths=[1.8, 1.0, 1.5, 2.0])
        self._h2("6.2 Market Share Indicators")
        self._table(["Indicator", "Frequency", "Signal", "Action Trigger"], [
            ["Transaction volume (YTD vs plan)", "Monthly", "Volume trajectory", ">5% below plan for 2 months"],
            ["Member retention (12m rolling)", "Monthly", "Relationship health", "<90% triggers retention campaign"],
            ["New member acquisitions", "Monthly", "Growth pipeline", "<80% of annual target pace"],
            ["Digital competitor penetration", "Quarterly", "Disintermediation speed", ">35% triggers defensive action"],
            ["Top-10 member concentration", "Quarterly", "Dependency risk", ">40% of volume"],
        ], col_widths=[1.8, 1.0, 1.5, 2.0])
        self._h2("6.3 Generational Indicators")
        self._table(["Indicator", "Frequency", "Signal", "Action Trigger"], [
            ["Average member age", "Annually", "Aging trajectory", ">60 triggers acceleration"],
            ["Young Rancher Program enrollment", "Monthly", "Pipeline health", "<100/yr by Year 2"],
            ["Succession transfers completed", "Quarterly", "Retention through transfer", "<30% retention rate"],
            ["University partnership engagement", "Quarterly", "Pipeline development", "<3 active partnerships"],
        ], col_widths=[1.8, 1.0, 1.5, 2.0])

    # ============ SECTION 7: MITIGATION STRATEGY ============
    def _add_section_7_mitigation(self):
        self._h1("7. Mitigation Strategy & Implementation Roadmap")
        self._body("The mitigation strategy addresses all three risk pillars — price, market share, "
            "and generational succession — through a unified Three Horizons framework consistent "
            "with Reports #1 and #2.")
        self._h2("7.1 Price Risk Mitigation Toolkit")
        self._h3("7.1.1 Hedging Overlay on Company-Owned Cattle")
        self._callout("Connection to Report #1",
            "Report #1 recommended a hedging program for the consignment-to-settlement window "
            "(Section 6.2). This report extends that program to cover company-owned inventory "
            "specifically. Costs here are ADDITIVE — the $300K hedging build-out covers additional "
            "positions and basis management infrastructure beyond Report #1's $150K hedging setup.",
            kind="sidebar")
        self._body("Implement systematic hedge coverage of 60-80% of company-owned cattle inventory. "
            "Use CME live cattle and feeder cattle futures with option overlays for basis protection. "
            "Target: reduce mark-to-market exposure from $4M (unhedged) to <$1M in a 20% decline.")
        self._h3("7.1.2 Commission Structure Reform")
        self._body("Move from pure percentage-of-sale commission to a blended model:")
        self._bullet_lead("Floor commission:", "Minimum per-head fee that provides revenue stability even in price declines.")
        self._bullet_lead("Tiered commission:", "Lower % rate on high-value cattle (incentivizes volume at peak prices).")
        self._bullet_lead("Fee-for-service:", "Unbundled advisory, hedging, and credit facilitation services priced separately.")
        self._bullet_lead("Subscription model:", "Fixed annual member fee for base services + variable commission on transactions.")
        self._h3("7.1.3 Working Capital Stress Testing")
        self._body("Conduct quarterly stress tests of PLMA's working capital against three scenarios "
            "(10%, 20%, 30% price decline) to ensure liquidity is adequate to sustain operations "
            "during a 12-24 month correction period.")
        self._h2("7.2 Market Share Defense & Growth")
        self._h3("7.2.1 Digital Platform")
        self._body("Cross-reference Report #2, Section 5.5: digital auction platform, mobile-first "
            "member experience, online sale capability. This is the single most important market share "
            "defense investment.")
        self._h3("7.2.2 Geographic Expansion")
        self._body("Evaluate expansion into 1-2 underserved regions within Year 2-3, targeting areas "
            "where auction market closures have created service gaps. Approaches: greenfield branch, "
            "partnership with existing regional cooperative, or acquisition.")
        self._h3("7.2.3 Value-Added Services")
        self._body("Launch fee-based advisory services: risk management advisory, hedging education, "
            "insurance brokerage, and market intelligence subscriptions. These services (i) generate "
            "non-commission revenue, (ii) deepen member relationships, and (iii) create switching costs.")
        self._h3("7.2.4 Cooperative Partnerships / M&A")
        self._body("Evaluate strategic partnerships or acquisitions of 1-2 smaller regional "
            "cooperatives in Years 3-5. Target: cooperatives with complementary geographic footprints "
            "and aging leadership (natural succession opportunity).")
        self._h2("7.3 Generational Succession Playbook")
        self._h3("7.3.1 Young Rancher Program (Formal)")
        self._body("Launch a branded Young Rancher Program with: (i) dedicated staff (1 FTE), "
            "(ii) regional events (livestock judging, field days, networking), (iii) educational "
            "content (risk management, marketing, finance), (iv) trial membership (reduced first-year "
            "commission), (v) mentorship matching with experienced members.")
        self._body("Target: 150-200 new young rancher relationships per year by Year 3. At "
            "$15-25K lifetime commission value per relationship, this generates $2.25-5.0M in "
            "future revenue pipeline annually.")
        self._h3("7.3.2 Estate Transition Liaison Service")
        self._body("Proactive outreach to members approaching retirement age (65+) to facilitate: "
            "(i) introduction to estate planning resources, (ii) identification of successor operators, "
            "(iii) continuity of PLMA marketing relationship through the transition, (iv) warm "
            "handoff from retiring member to successor. Target: increase succession retention rate "
            "from ~29% to 55% within 5 years.")
        self._h3("7.3.3 Next-Gen Digital Experience")
        self._body("Young producers expect digital-first interaction: mobile auction participation, "
            "real-time market data, instant settlement, electronic documentation, and social media "
            "engagement. This is the generation that will judge PLMA by its app, not its "
            "auction ring.")
        self._h3("7.3.4 University / Extension Partnerships")
        self._body("Target 5+ land-grant university partnerships within 3 years: KSU, TAMU, UNL, "
            "SDSU, CSU, OSU. Sponsor livestock judging teams, host guest lectures, fund "
            "cooperative-focused research, and establish internship programs.")
        self._h2("7.4 1-Year / 5-Year / 10-Year Plans")
        self._chart(self.charts.three_horizons(),
                    caption="13 workstreams across 10 years, structured by McKinsey Three Horizons.")
        self._h3("Year 1 — Foundation")
        self._bullet("Launch hedging program for company-owned cattle (Month 3)")
        self._bullet("Commission structure analysis and pilot (Months 1-12)")
        self._bullet("Deploy leading indicator dashboard (Month 6)")
        self._bullet("Launch Young Rancher Program (Month 6)")
        self._bullet("Hire dedicated program staff (Month 3)")
        self._bullet("Initiate 2 university partnerships (Month 9)")
        self._h3("5-Year Targets")
        self._bullet("Geographic expansion: 1-2 new regions operational")
        self._bullet("Advisory services generating $1.2M+/year in fee revenue")
        self._bullet("Young Rancher Program: 500+ cumulative enrollees")
        self._bullet("Succession retention rate: >55% (from ~29% baseline)")
        self._bullet("Digital platform live with hybrid auction capability")
        self._bullet("Member age trend stabilized (avg age growth slowed)")
        self._h3("10-Year Vision")
        self._bullet("15-20% of revenue from non-commission sources")
        self._bullet("1-2 strategic cooperative acquisitions completed")
        self._bullet("Young rancher cohort represents >20% of active members")
        self._bullet("PLMA is the reference cooperative for next-gen engagement")
        self._bullet("Commission structure fully modernized (blended model)")

    # ============ SECTION 8: FINANCIAL MODEL ============
    def _add_section_8_financials(self):
        self._h1("8. Financial Model")
        self._h2("8.1 Commission Revenue Scenarios")
        self._chart(self.charts.commission_scenarios_line(),
                    caption="With mitigation, PLMA stabilizes revenue even under base-case price decline — a $8M+/yr gap vs. unmitigated by Year 5.")
        self._h2("8.2 Program Costs")
        y1_rows = [[k, f"${v/1000:.0f}K"] for k, v in self.model.year1_costs.items()]
        y1_rows.append(["**TOTAL Year 1**", f"${self.model.total_year1_cost()/1000:.0f}K"])
        self._table(["Year 1 Investment", "Amount"], y1_rows, col_widths=[4.5, 1.5])
        ong_rows = [[k, f"${v/1000:.0f}K"] for k, v in self.model.ongoing_costs.items()]
        ong_rows.append(["**TOTAL Ongoing**", f"${self.model.total_ongoing_cost()/1000:.0f}K"])
        self._table(["Ongoing Annual Cost (Year 2+)", "Amount"], ong_rows, col_widths=[4.5, 1.5])
        self._h2("8.3 Annual Benefits")
        ben_rows = [[k, f"${v/1000:.0f}K"] for k, v in self.model.annual_benefits.items()]
        ben_rows.append(["**TOTAL Annual Benefit**", f"${self.model.total_annual_benefit()/1000:.0f}K"])
        self._table(["Benefit Category", "Annual Value"], ben_rows, col_widths=[4.5, 1.5])
        self._h2("8.4 10-Year Cash Flows & NPV")
        cf_rows = []
        cum = 0
        for f in self.model.cash_flows():
            cum += f["net"]
            cf_rows.append([f"Year {f['year']}", f"${f['cost']/1000:.0f}K",
                           f"${f['benefit']/1000:.0f}K", f"${f['net']/1000:.0f}K",
                           f"${cum/1000:.0f}K"])
        self._table(["Year", "Cost", "Benefit", "Net CF", "Cumulative"], cf_rows,
                    col_widths=[0.9, 1.1, 1.1, 1.1, 1.3])
        npv_val = self.model.npv()
        roi = self.model.roi_10yr()
        payback = self.model.payback_months()
        self._body(f"10-Year NPV (8% discount): ${npv_val/1e6:.1f}M | ROI: {roi*100:.0f}% | "
                   f"Payback: {payback} months")
        self._chart(self.charts.roi_projection(),
                    caption="Program breaks even during Year 2; cumulative 10-year net exceeds $37M nominally.")
        self._h2("8.5 Sensitivity Analysis")
        self._chart(self.charts.sensitivity_tornado(),
                    caption="Benefit realization and cattle price severity are the top NPV drivers; program is robust across scenarios.")
        self._callout("Robustness",
            "The program maintains positive NPV across all tested sensitivity scenarios, including "
            "the most adverse combination. Even if benefits materialize at only 70% of estimate, "
            "NPV remains strongly positive. The investment case does not depend on optimistic assumptions.",
            kind="finding")

    # ============ SECTION 9: CONCLUSION ============
    def _add_section_9_conclusion(self):
        self._h1("9. Conclusion & Next Steps")
        self._callout("Three Reports, One Transformation",
            "Reports #1, #2, and #3 are not three separate initiatives. They are three lenses on a "
            "single imperative: PLMA must modernize its risk management, technology, and market "
            "strategy to survive and thrive through a generational transition in the livestock industry. "
            "Credit risk (#1), tech transformation (#2), and price/market/succession risk (#3) share "
            "common governance, overlapping investments, and a unified Three Horizons roadmap.",
            kind="sidebar")
        self._body("PLMA's leadership team faces three concurrent structural challenges: a cattle "
            "cycle peak that will compress revenue, market share erosion from digital disruption and "
            "consolidation, and a demographic cliff that threatens the member base. None of these "
            "risks is fatal on its own. But in combination, without mitigation, they compound into "
            "an existential challenge within 5-10 years.")
        self._body("The recommended mitigation program — $2.8M Year 1, $1.8M ongoing — generates "
            f"~${self.model.npv()/1e6:.0f}M in 10-year NPV and positions PLMA to emerge from "
            "the cattle cycle correction as a stronger, more diversified, and younger cooperative.")
        self._h2("9.1 Immediate Actions (First 90 Days)")
        self._bullet("Week 1-2: CFO-led briefing on Report #3 findings; integrate with Reports #1 & #2 action items")
        self._bullet("Week 2-4: Board approval of Year 1 budget ($2.8M); integrate into unified transformation budget")
        self._bullet("Week 4-6: Launch hedging program for company-owned cattle")
        self._bullet("Week 4-8: Commission structure working group formed; begin analysis")
        self._bullet("Week 6-10: Young Rancher Program staff hired; university outreach initiated")
        self._bullet("Week 8-12: Leading indicator dashboard deployed (v1)")
        self._bullet("Week 12: Integrated PMO operational covering all three report workstreams")
        self._h2("9.2 Integrated Budget Summary (All 3 Reports)")
        self._table(["Report", "Year 1 Cost", "Ongoing/yr", "10-yr NPV"],
            [["#1: Credit Risk Mitigation", "$1.72M", "$1.17M", "~$20M"],
             ["#2: Technology Strategy", "$4.80M", "$2.60M", "~$34M"],
             ["#3: Price & Market Strategy", "$2.80M", "$1.80M", f"~${self.model.npv()/1e6:.0f}M"],
             ["TOTAL (de-duplicated)", "$9.32M", "$5.57M", "~$82M"]],
            col_widths=[2.5, 1.2, 1.2, 1.2])
        self._body("The combined program represents a $9.3M Year-1 investment generating ~$82M in "
            "aggregate 10-year NPV — a decisive, financially justified transformation of PLMA's "
            "risk posture, technology capability, and market position.")

    # ============ APPENDICES ============
    def _add_appendices(self):
        self._h1("Appendix A: Cattle Cycle Data Tables")
        self._table(["Year", "Fed Cattle ($/cwt)", "Feeder ($/cwt)", "Cow Herd (M head)"], [
            ["1990", "$82", "$95", "99.3"], ["1995", "$63", "$72", "102.8"],
            ["2000", "$68", "$95", "98.2"], ["2005", "$87", "$118", "95.8"],
            ["2010", "$95", "$128", "93.9"], ["2014", "$160", "$242", "89.1"],
            ["2017", "$108", "$148", "93.6"], ["2020", "$118", "$145", "93.8"],
            ["2022", "$142", "$185", "91.9"], ["2024", "$190", "$268", "89.3"],
            ["2025", "$195", "$278", "88.5"], ["2026 (Q1)", "$198", "$282", "88.2"],
        ], col_widths=[1.0, 1.5, 1.5, 1.5])

        self._h1("Appendix B: Commission Sensitivity Model Detail")
        self._body("Model assumes: $3.2B transaction volume, 95% cattle / 5% lamb split, "
            "1.875% average commission rate. Cattle price declines applied to cattle volume only; "
            "lamb volume held constant (immaterial). Scenarios represent annual average price "
            "declines from peak-cycle levels.")

        self._h1("Appendix C: Generational Demographics")
        self._table(["Age Group", "2012 Census", "2022 Census", "2035 Projected", "Change"],
            [["<35", "5.6%", "4.9%", "4.2%", "-1.4pp"],
             ["35-44", "10.2%", "8.3%", "7.5%", "-2.7pp"],
             ["45-54", "21.8%", "16.1%", "12.8%", "-9.0pp"],
             ["55-64", "30.6%", "27.4%", "22.5%", "-8.1pp"],
             ["65-74", "22.1%", "27.3%", "29.5%", "+7.4pp"],
             ["75+", "9.7%", "16.0%", "23.5%", "+13.8pp"]], col_widths=[1.2, 1.2, 1.2, 1.2, 1.0])

        self._h1("Appendix D: Methodology & Assumptions")
        self._body("Financial model uses standard DCF with 8% discount rate. Benefits ramp: 40% Year 1, "
            "80% Year 2, 100% Year 3+. Commission sensitivity model uses linear price-volume relationship "
            "(elasticity = 1.0 for simplicity). Inventory mark-to-market assumes flat basis. Generational "
            "projections use cohort survival model with USDA Census data. Market share erosion rates "
            "benchmarked from LMA industry reports.")

        self._h1("Appendix E: Glossary")
        for term, defn in [
            ("Basis", "Difference between local cash price and CME futures price."),
            ("Cattle Cycle", "10-12 year biological/economic cycle of herd expansion and contraction."),
            ("Contango", "Futures curve where later months are higher than nearby — bearish signal."),
            ("Backwardation", "Futures curve where nearby months are higher — bullish/tight-supply signal."),
            ("Cost-of-Gain", "Feed and operational cost per pound of weight added in a feedlot."),
            ("Disintermediation", "Removal of intermediaries (like PLMA) from supply chain transactions."),
            ("EID", "Electronic Identification — RFID tags for individual animal tracking."),
            ("Feeder Cattle", "Weaned calves or yearlings ready for feedlot placement."),
            ("Mark-to-Market", "Revaluing inventory to current market prices; unrealized gain/loss."),
            ("P&SA", "Packers & Stockyards Act — USDA livestock marketing regulation."),
        ]:
            self._bullet_lead(term + ":", defn)

        self._add_citations()

    def _add_citations(self):
        self._h1("Appendix F: Citations")
        citations = [
            "USDA NASS, \"Cattle on Feed,\" monthly reports 2024-2026.",
            "USDA NASS, \"Cattle,\" semi-annual report, January 2026.",
            "USDA ERS, \"Cattle & Beef: Outlook,\" March 2026.",
            "CME Group, \"Live Cattle and Feeder Cattle Futures: Market Analysis,\" 2025-2026.",
            "Livestock Marketing Information Center (LMIC), \"Monthly Livestock Outlook,\" 2025-2026.",
            "Kansas State University, \"AgManager: Livestock Market Analysis,\" 2025-2026.",
            "USDA AMS Packers & Stockyards Division, \"Annual Report,\" 2025.",
            "USDA ERS, \"Farm Income and Wealth Statistics,\" 2025-2026.",
            "Federal Reserve Bank of Kansas City, \"Agricultural Finance Databook,\" Q4 2025.",
            "Livestock Marketing Association, \"Annual Industry Census,\" 2024.",
            "USDA NASS, \"Census of Agriculture,\" 2022 (released 2024).",
            "USDA ERS, \"America's Farms and Ranches at a Glance,\" 2025.",
            "Rabobank, \"Global Animal Protein Outlook,\" 2026.",
            "Sterling Marketing Inc., \"Cattle Feeding Industry Analysis,\" Q1 2026.",
            "Superior Livestock Auction, \"Annual Transaction Report,\" 2025.",
            "CoBank, \"U.S. Agricultural Cooperatives: Risk and Resilience,\" 2025.",
            "S&P Global Market Intelligence, \"U.S. Agribusiness Credit Trends,\" 2025.",
            "McKinsey & Company, \"Agriculture Practice: Building Resilient Food Systems,\" 2024.",
            "Porter, M.E., \"Competitive Strategy,\" Free Press, 1980.",
            "Covey, S.R., \"The 7 Habits of Highly Effective People,\" 1989 (Circle of Influence).",
            "Harvard Business Review, \"Managing Risks: A New Framework,\" Kaplan & Mikes, 2012.",
            "USDA NIFA, \"Beginning Farmer and Rancher Development Program,\" 2024-2025.",
            "National Association of State Departments of Agriculture, \"Young Farmer Programs,\" 2025.",
            "CME Group, \"Basis: Understanding Cash-Futures Relationships,\" educational resources.",
            "Farm Credit Administration, \"Young, Beginning, and Small Farmer Mission,\" 2025.",
        ]
        for i, cite in enumerate(citations, 1):
            p = self.doc.add_paragraph()
            r1 = p.add_run(f"[{i}]  "); r1.bold = True
            r1.font.name = self.config.BODY_FONT; r1.font.size = Pt(9.5)
            r1.font.color.rgb = self.config.NAVY
            r2 = p.add_run(cite)
            r2.font.name = self.config.BODY_FONT; r2.font.size = Pt(9.5)
            r2.font.color.rgb = self.config.DARK_GRAY
            p.paragraph_format.left_indent = Inches(0.4)
            p.paragraph_format.first_line_indent = Inches(-0.4)
            set_paragraph_spacing(p, after=5, line=1.2)

    # ============ BUILD ============
    def build(self):
        self._add_cover_page()
        self._add_toc()
        self._add_executive_summary()
        self._add_section_2_cattle()
        self._add_section_3_market_share()
        self._add_section_4_succession()
        self._add_section_5_control()
        self._add_section_6_indicators()
        self._add_section_7_mitigation()
        self._add_section_8_financials()
        self._add_section_9_conclusion()
        self._add_appendices()
        return self.doc


# =============================================================================
# MAIN
# =============================================================================
def main():
    config = ReportConfig()
    model = PriceMarketModel()
    charts = PriceChartGenerator(config, model)
    builder = DocumentBuilder(config, charts, model)
    builder.build()
    output = "PLMA_Price_Market_Strategy_2026.docx"
    builder.doc.save(output)
    print(f"Report generated: {output}")
    print(f"10-Year NPV: ${model.npv()/1e6:.2f}M")
    print(f"10-Year ROI: {model.roi_10yr()*100:.0f}%")
    print(f"Payback: {model.payback_months()} months")
    sens = model.commission_sensitivity()
    for s in sens:
        print(f"  {int(s['decline']*100)}% cattle decline → -${s['revenue_loss']/1e6:.1f}M ({s['pct_of_revenue']*100:.1f}%)")


if __name__ == "__main__":
    main()
