"""
PLMA Counterparty Credit Risk Assessment Report Generator
Generates a McKinsey-caliber Word document for CFO Ryan Johansen.
"""
import io
from io import BytesIO

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, Rectangle, Polygon
import numpy as np

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_BREAK
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION, WD_ORIENTATION
from docx.oxml.ns import qn, nsmap
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
    DARK_GRAY = RGBColor(0x33, 0x33, 0x33)
    MID_GRAY = RGBColor(0x66, 0x66, 0x66)
    LIGHT_GRAY = RGBColor(0xD9, 0xD9, 0xD9)
    VERY_LIGHT_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)

    # Hex for matplotlib
    NAVY_HEX = "#003A70"
    BLUE_HEX = "#0078D4"
    LIGHT_BLUE_HEX = "#00A4EF"
    RED_HEX = "#E81123"
    AMBER_HEX = "#FFB900"
    GREEN_HEX = "#107C10"
    DARK_GRAY_HEX = "#333333"
    MID_GRAY_HEX = "#666666"
    LIGHT_GRAY_HEX = "#D9D9D9"

    HEADING_FONT = "Calibri"
    BODY_FONT = "Calibri"
    CHART_FONT = "DejaVu Sans"

    # Shading hex (no #)
    NAVY_SHADE = "003A70"
    LIGHT_BLUE_SHADE = "DEEBF7"
    VERY_LIGHT_GRAY_SHADE = "F2F2F2"
    AMBER_SHADE = "FFF4CE"
    RED_SHADE = "FDE7E9"
    GREEN_SHADE = "DFF6DD"


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


def remove_cell_borders(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "nil")
        tc_borders.append(el)
    tc_pr.append(tc_borders)


def add_page_number_field(paragraph):
    run = paragraph.add_run()
    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = "PAGE"
    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "end")
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
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
class FinancialModel:
    def __init__(self):
        # Transaction volume and revenue estimates (industry-based)
        self.transaction_volume = 3_200_000_000  # $3.2B
        self.revenue = 60_000_000  # ~1.9% of volume
        self.working_capital_exposure = 200_000_000  # $200M

        # Expected Loss components by segment
        self.segments = [
            {"name": "Independent Feedlots", "ead": 80_000_000, "pd": 0.080, "lgd": 0.55},
            {"name": "Small/Mid Processors", "ead": 50_000_000, "pd": 0.060, "lgd": 0.45},
            {"name": "Large Packers", "ead": 40_000_000, "pd": 0.015, "lgd": 0.30},
            {"name": "Backgrounders/Stocker Ops", "ead": 30_000_000, "pd": 0.120, "lgd": 0.60},
        ]

        # Mitigation costs
        self.year1_costs = {
            "Credit scoring system (build/buy)": 350_000,
            "Trade credit insurance premiums": 600_000,
            "Risk analytics platform": 275_000,
            "Credit analyst FTEs (2)": 240_000,
            "Hedging program setup": 150_000,
            "Training & change management": 100_000,
        }
        self.ongoing_costs = {
            "Credit scoring maintenance": 85_000,
            "Trade credit insurance premiums": 600_000,
            "Risk analytics platform": 120_000,
            "Credit analyst FTEs (2)": 250_000,
            "Hedging advisory": 75_000,
            "Training & change management": 40_000,
        }

        # Benefits
        self.annual_benefits = {
            "Credit scoring + tiered limits (25% UL reduction)": 1_800_000,
            "Trade credit insurance (net recovery)": 1_300_000,
            "Early warning system (15% add'l avoidance)": 650_000,
            "Hedging program market risk reduction": 800_000,
        }

        self.discount_rate = 0.08
        self.horizon_years = 10

    def expected_loss_by_segment(self):
        results = []
        for s in self.segments:
            el = s["ead"] * s["pd"] * s["lgd"]
            results.append({**s, "el": el})
        return results

    def total_expected_loss(self):
        return sum(r["el"] for r in self.expected_loss_by_segment())

    def stress_expected_loss(self, pd_multiplier=2.0):
        return sum(s["ead"] * s["pd"] * pd_multiplier * s["lgd"] for s in self.segments)

    def total_year1_cost(self):
        return sum(self.year1_costs.values())

    def total_ongoing_cost(self):
        return sum(self.ongoing_costs.values())

    def total_annual_benefit(self):
        return sum(self.annual_benefits.values())

    def cash_flows(self):
        """Returns list of (year, cost, benefit, net) for years 1-10."""
        flows = []
        for y in range(1, self.horizon_years + 1):
            if y == 1:
                cost = self.total_year1_cost()
                benefit = self.total_annual_benefit() * 0.5  # ramp
            else:
                cost = self.total_ongoing_cost()
                benefit = self.total_annual_benefit()
            flows.append({"year": y, "cost": cost, "benefit": benefit, "net": benefit - cost})
        return flows

    def npv(self):
        total = 0
        for f in self.cash_flows():
            total += f["net"] / ((1 + self.discount_rate) ** f["year"])
        return total

    def cumulative_net(self):
        cum = 0
        result = []
        for f in self.cash_flows():
            cum += f["net"]
            result.append({"year": f["year"], "cumulative": cum})
        return result

    def roi_10yr(self):
        total_cost = self.total_year1_cost() + self.total_ongoing_cost() * (self.horizon_years - 1)
        total_benefit = self.total_annual_benefit() * 0.5 + self.total_annual_benefit() * (self.horizon_years - 1)
        return (total_benefit - total_cost) / total_cost

    def payback_months(self):
        cum = 0
        monthly_benefit = self.total_annual_benefit() / 12
        year1_monthly_benefit = (self.total_annual_benefit() * 0.5) / 12
        year1_cost = self.total_year1_cost()
        for m in range(1, 121):
            if m <= 12:
                cum += year1_monthly_benefit
                if m == 1:
                    cum -= year1_cost  # upfront cost
            else:
                cum += monthly_benefit - (self.total_ongoing_cost() / 12)
            if cum >= 0:
                return m
        return None

    def sensitivity(self):
        """Returns list of (variable, low_npv, high_npv) for tornado chart."""
        base_npv = self.npv()
        results = []

        # Default rate +/- 50%
        orig_segs = [s.copy() for s in self.segments]
        for mult, label in [(0.5, "low"), (1.5, "high")]:
            for s in self.segments:
                s["pd"] *= mult
            # benefits scale with EL reduction
            el_mult = mult
            orig_benefits = self.annual_benefits.copy()
            for k in self.annual_benefits:
                self.annual_benefits[k] *= el_mult if "loss" in k.lower() or "warning" in k.lower() or "credit" in k.lower() else 1.0
            npv_val = self.npv()
            results.append(("Default rate", label, npv_val))
            self.annual_benefits = orig_benefits.copy()
            for i, s in enumerate(self.segments):
                s["pd"] = orig_segs[i]["pd"]

        # Recovery rate (LGD) +/- 20%
        for mult, label in [(0.8, "low"), (1.2, "high")]:
            for s in self.segments:
                s["lgd"] *= mult
            npv_val = self.npv()
            results.append(("Recovery rate (LGD)", label, npv_val))
            for i, s in enumerate(self.segments):
                s["lgd"] = orig_segs[i]["lgd"]

        # Insurance cost +/- 30%
        orig_ongoing = self.ongoing_costs.copy()
        orig_y1 = self.year1_costs.copy()
        for mult, label in [(0.7, "low"), (1.3, "high")]:
            self.ongoing_costs["Trade credit insurance premiums"] = 600_000 * mult
            self.year1_costs["Trade credit insurance premiums"] = 600_000 * mult
            npv_val = self.npv()
            results.append(("Insurance premium", label, npv_val))
            self.ongoing_costs = orig_ongoing.copy()
            self.year1_costs = orig_y1.copy()

        # Implementation cost overrun
        for mult, label in [(1.0, "low"), (1.5, "high")]:
            for k in self.year1_costs:
                self.year1_costs[k] = orig_y1[k] * mult
            npv_val = self.npv()
            results.append(("Implementation cost", label, npv_val))
            self.year1_costs = orig_y1.copy()

        # Discount rate 6-12%
        orig_dr = self.discount_rate
        for rate, label in [(0.06, "low"), (0.12, "high")]:
            self.discount_rate = rate
            npv_val = self.npv()
            results.append(("Discount rate", label, npv_val))
        self.discount_rate = orig_dr

        # Benefit realization +/- 25%
        orig_benefits = self.annual_benefits.copy()
        for mult, label in [(0.75, "low"), (1.25, "high")]:
            for k in self.annual_benefits:
                self.annual_benefits[k] = orig_benefits[k] * mult
            npv_val = self.npv()
            results.append(("Benefit realization", label, npv_val))
        self.annual_benefits = orig_benefits.copy()

        # Aggregate into tornado format
        variables = {}
        for var, lbl, npv_val in results:
            if var not in variables:
                variables[var] = {}
            variables[var][lbl] = npv_val

        tornado = []
        for var, vals in variables.items():
            tornado.append({
                "variable": var,
                "low": vals["low"],
                "high": vals["high"],
                "base": base_npv,
                "range": abs(vals["high"] - vals["low"]),
            })
        tornado.sort(key=lambda x: x["range"], reverse=True)
        return tornado, base_npv


# =============================================================================
# CHART GENERATOR
# =============================================================================
class ChartGenerator:
    def __init__(self, config, model):
        self.config = config
        self.model = model
        plt.rcParams["font.family"] = "DejaVu Sans"
        plt.rcParams["font.size"] = 10
        plt.rcParams["axes.edgecolor"] = "#666666"
        plt.rcParams["axes.linewidth"] = 0.8
        plt.rcParams["axes.labelcolor"] = "#333333"
        plt.rcParams["xtick.color"] = "#333333"
        plt.rcParams["ytick.color"] = "#333333"
        plt.rcParams["axes.spines.top"] = False
        plt.rcParams["axes.spines.right"] = False

    def _finalize(self, fig, dpi=200):
        buf = BytesIO()
        fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        buf.seek(0)
        return buf

    def farm_bankruptcy_trend(self):
        years = [2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026]
        values = [498, 595, 552, 371, 275, 287, 258, 315, 380]
        is_forecast = [False]*8 + [True]
        fig, ax = plt.subplots(figsize=(9, 5))
        colors = [self.config.NAVY_HEX if not f else self.config.AMBER_HEX for f in is_forecast]
        bars = ax.bar(years, values, color=colors, edgecolor="white", linewidth=1, zorder=3)
        # Trendline
        z = np.polyfit(range(len(years)), values, 2)
        p = np.poly1d(z)
        ax.plot(years, p(range(len(years))), color=self.config.RED_HEX, linewidth=2.5,
                linestyle="--", label="Trend", zorder=4)
        ax.set_title("Chapter 12 Farm Bankruptcies (2018–2026E)", fontsize=13,
                     fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.set_ylabel("Number of Filings", fontsize=10)
        ax.set_xlabel("Year", fontsize=10)
        ax.grid(True, axis="y", alpha=0.3, zorder=0)
        ax.set_axisbelow(True)
        for bar, val in zip(bars, values):
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 8,
                    f"{val}", ha="center", fontsize=9, color=self.config.DARK_GRAY_HEX)
        ax.annotate("2025–2026\nAcceleration", xy=(2025.5, 350), xytext=(2023, 500),
                    fontsize=10, color=self.config.RED_HEX, fontweight="bold",
                    arrowprops=dict(arrowstyle="->", color=self.config.RED_HEX, lw=1.5))
        ax.legend(loc="upper right", frameon=False)
        ax.text(0.01, -0.15, "Source: American Bankruptcy Institute; 2026 projection.",
                transform=ax.transAxes, fontsize=8, color=self.config.MID_GRAY_HEX, style="italic")
        return self._finalize(fig)

    def farm_debt_growth(self):
        years = list(range(2015, 2027))
        re_debt = [213, 223, 240, 253, 263, 279, 300, 325, 353, 382, 405, 430]
        non_re_debt = [172, 180, 190, 198, 205, 215, 210, 205, 210, 215, 220, 230]
        fig, ax = plt.subplots(figsize=(9, 5))
        ax.fill_between(years, 0, re_debt, color=self.config.NAVY_HEX, alpha=0.85, label="Real Estate Debt")
        ax.fill_between(years, re_debt, [r+n for r,n in zip(re_debt,non_re_debt)],
                        color=self.config.BLUE_HEX, alpha=0.85, label="Non-Real-Estate Debt")
        total = [r+n for r,n in zip(re_debt,non_re_debt)]
        ax.plot(years, total, color=self.config.RED_HEX, linewidth=2.5, marker="o", markersize=5, label="Total")
        ax.annotate(f"${total[-2]}B\n(2025)", xy=(2025, total[-2]), xytext=(2022.5, 680),
                    fontsize=10, fontweight="bold", color=self.config.RED_HEX,
                    arrowprops=dict(arrowstyle="->", color=self.config.RED_HEX))
        ax.set_title("Total U.S. Farm Sector Debt (2015–2026E)", fontsize=13,
                     fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.set_ylabel("Debt ($ Billions)", fontsize=10)
        ax.set_xlabel("Year", fontsize=10)
        ax.legend(loc="upper left", frameon=False)
        ax.grid(True, axis="y", alpha=0.3)
        ax.set_axisbelow(True)
        ax.text(0.01, -0.15, "Source: USDA Economic Research Service, Farm Sector Balance Sheet (Feb 2026).",
                transform=ax.transAxes, fontsize=8, color=self.config.MID_GRAY_HEX, style="italic")
        return self._finalize(fig)

    def risk_heat_map(self):
        fig, ax = plt.subplots(figsize=(8, 7))
        # Create 5x5 grid
        for i in range(5):
            for j in range(5):
                severity = (i+1) * (j+1)
                if severity <= 4:
                    color = "#DFF6DD"
                elif severity <= 9:
                    color = "#FFF4CE"
                elif severity <= 15:
                    color = "#FED9B7"
                else:
                    color = "#FDE7E9"
                ax.add_patch(Rectangle((j, i), 1, 1, facecolor=color, edgecolor="white", linewidth=2))
        risks = [
            ("Buyer bankruptcy (feedlot)", 4, 4),
            ("Producer default", 4, 3),
            ("Packer consolidation", 2, 4),
            ("Cattle price collapse", 3, 4),
            ("ASF outbreak", 2, 5),
            ("Interest rate spike", 3, 3),
            ("Regulatory (P&SA)", 2, 2),
            ("IT/cyber breach", 2, 3),
            ("Key personnel loss", 3, 2),
            ("Feed cost spike", 3, 3),
            ("Basis risk", 4, 2),
            ("Concentration risk", 3, 4),
        ]
        for name, likelihood, impact in risks:
            ax.scatter(likelihood-0.5, impact-0.5, s=150, c=self.config.NAVY_HEX,
                       edgecolor="white", linewidth=2, zorder=5)
            ax.annotate(name, (likelihood-0.5, impact-0.5), fontsize=8,
                        xytext=(5, 5), textcoords="offset points", color=self.config.DARK_GRAY_HEX)
        ax.set_xlim(0, 5)
        ax.set_ylim(0, 5)
        ax.set_xticks([0.5, 1.5, 2.5, 3.5, 4.5])
        ax.set_xticklabels(["Rare", "Unlikely", "Possible", "Likely", "Certain"])
        ax.set_yticks([0.5, 1.5, 2.5, 3.5, 4.5])
        ax.set_yticklabels(["Negligible", "Minor", "Moderate", "Major", "Severe"])
        ax.set_xlabel("Likelihood", fontsize=11, fontweight="bold")
        ax.set_ylabel("Impact", fontsize=11, fontweight="bold")
        ax.set_title("PLMA Counterparty Risk Heat Map", fontsize=13,
                     fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.set_aspect("equal")
        return self._finalize(fig)

    def buyer_default_probability(self):
        segments = ["Backgrounders/\nStocker Ops", "Independent\nFeedlots",
                    "Small/Mid\nProcessors", "Large\nPackers"]
        pds = [12.0, 8.0, 6.0, 1.5]
        colors = [self.config.RED_HEX, self.config.AMBER_HEX,
                  self.config.BLUE_HEX, self.config.GREEN_HEX]
        fig, ax = plt.subplots(figsize=(9, 5))
        bars = ax.barh(segments, pds, color=colors, edgecolor="white", linewidth=1)
        for bar, val in zip(bars, pds):
            ax.text(bar.get_width() + 0.2, bar.get_y() + bar.get_height()/2,
                    f"{val:.1f}%", va="center", fontsize=10, fontweight="bold",
                    color=self.config.DARK_GRAY_HEX)
        ax.set_title("Estimated 12-Month Default Probability by Buyer Segment",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.set_xlabel("Default Probability (%)", fontsize=10)
        ax.grid(True, axis="x", alpha=0.3)
        ax.set_axisbelow(True)
        ax.set_xlim(0, 15)
        ax.text(0.01, -0.18, "Source: S&P Global Market Intelligence; Rabobank Animal Protein Outlook 2026; analyst estimates.",
                transform=ax.transAxes, fontsize=8, color=self.config.MID_GRAY_HEX, style="italic")
        return self._finalize(fig)

    def commodity_correlation_heatmap(self):
        labels = ["Live Cattle", "Feeder Cattle", "Lean Hogs", "Corn", "Soybean Meal", "Fed Funds"]
        corr = np.array([
            [1.00, 0.78, 0.42, -0.25, -0.18, -0.35],
            [0.78, 1.00, 0.38, -0.52, -0.38, -0.28],
            [0.42, 0.38, 1.00, -0.15, -0.20, -0.22],
            [-0.25, -0.52, -0.15, 1.00, 0.72, 0.18],
            [-0.18, -0.38, -0.20, 0.72, 1.00, 0.12],
            [-0.35, -0.28, -0.22, 0.18, 0.12, 1.00],
        ])
        fig, ax = plt.subplots(figsize=(8, 7))
        im = ax.imshow(corr, cmap="RdBu_r", vmin=-1, vmax=1, aspect="equal")
        ax.set_xticks(range(len(labels)))
        ax.set_yticks(range(len(labels)))
        ax.set_xticklabels(labels, rotation=45, ha="right")
        ax.set_yticklabels(labels)
        for i in range(len(labels)):
            for j in range(len(labels)):
                color = "white" if abs(corr[i][j]) > 0.5 else "black"
                ax.text(j, i, f"{corr[i][j]:.2f}", ha="center", va="center",
                        fontsize=9, color=color, fontweight="bold")
        ax.set_title("Commodity Price Correlation Matrix (2020–2026)",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        cbar = plt.colorbar(im, ax=ax, shrink=0.7)
        cbar.set_label("Correlation", fontsize=9)
        ax.text(0.01, -0.25, "Source: CME Group, FRED; monthly returns.",
                transform=ax.transAxes, fontsize=8, color=self.config.MID_GRAY_HEX, style="italic")
        return self._finalize(fig)

    def interest_rate_sensitivity(self):
        rates = np.arange(3.0, 7.5, 0.25)
        base_rate = 4.50
        cost_impact = (rates - base_rate) * self.model.working_capital_exposure / 100 * 0.75
        diff = np.maximum(rates - base_rate, 0)
        buyer_stress = diff ** 1.8 * 2.2
        fig, ax1 = plt.subplots(figsize=(9, 5))
        ax1.plot(rates, cost_impact / 1_000_000, color=self.config.NAVY_HEX,
                 linewidth=2.5, marker="o", markersize=5, label="PLMA Cost-of-Carry Impact")
        ax1.set_xlabel("Federal Funds Rate (%)", fontsize=10)
        ax1.set_ylabel("PLMA Annual Cost Impact ($M)", fontsize=10, color=self.config.NAVY_HEX)
        ax1.tick_params(axis="y", labelcolor=self.config.NAVY_HEX)
        ax1.axvline(x=base_rate, color=self.config.MID_GRAY_HEX, linestyle="--", alpha=0.6)
        ax1.text(base_rate+0.05, ax1.get_ylim()[1]*0.9, "Current", fontsize=9, color=self.config.MID_GRAY_HEX)
        ax1.grid(True, alpha=0.3)
        ax2 = ax1.twinx()
        ax2.plot(rates, buyer_stress, color=self.config.RED_HEX,
                 linewidth=2.5, marker="s", markersize=5, label="Buyer Stress Index")
        ax2.set_ylabel("Buyer Financial Stress (Index)", fontsize=10, color=self.config.RED_HEX)
        ax2.tick_params(axis="y", labelcolor=self.config.RED_HEX)
        ax2.spines["top"].set_visible(False)
        ax1.set_title("Interest Rate Sensitivity Analysis",
                      fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        lines1, labels1 = ax1.get_legend_handles_labels()
        lines2, labels2 = ax2.get_legend_handles_labels()
        ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper left", frameon=False)
        fig.text(0.1, -0.02, "Source: FRED; internal analysis based on $200M working capital exposure.",
                 fontsize=8, color=self.config.MID_GRAY_HEX, style="italic")
        return self._finalize(fig)

    def porters_five_forces(self):
        categories = ["Buyer Power\n(Packers)", "Supplier Power\n(Producers)",
                      "New Entrants", "Substitutes\n(Direct Sale)", "Rivalry\n(Coops/Auctions)"]
        values = [4, 3, 2, 2, 4]
        angles = np.linspace(0, 2*np.pi, len(categories), endpoint=False).tolist()
        values_plot = values + values[:1]
        angles_plot = angles + angles[:1]
        fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(projection="polar"))
        ax.plot(angles_plot, values_plot, color=self.config.NAVY_HEX, linewidth=2.5)
        ax.fill(angles_plot, values_plot, color=self.config.BLUE_HEX, alpha=0.35)
        ax.set_xticks(angles)
        ax.set_xticklabels(categories, fontsize=10, fontweight="bold")
        ax.set_yticks([1, 2, 3, 4, 5])
        ax.set_yticklabels(["1", "2", "3", "4", "5"], fontsize=8)
        ax.set_ylim(0, 5)
        ax.set_title("Porter's Five Forces: Livestock Marketing Industry",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=25)
        ax.grid(True, alpha=0.4)
        for angle, val, cat in zip(angles, values, categories):
            ax.text(angle, val + 0.3, str(val), ha="center", fontsize=11,
                    fontweight="bold", color=self.config.RED_HEX)
        return self._finalize(fig)

    def swot_quadrant(self):
        fig, ax = plt.subplots(figsize=(10, 7))
        ax.set_xlim(0, 10)
        ax.set_ylim(0, 10)
        ax.axis("off")
        # Quadrants
        quads = [
            (0, 5, 5, 5, "#DFF6DD", "STRENGTHS", [
                "• 100+ year cooperative heritage & member trust",
                "• Established market access across multi-state footprint",
                "• Deep producer relationships",
                "• Market price discovery expertise",
                "• Regulatory compliance infrastructure (P&SA)"]),
            (5, 5, 5, 5, "#DEEBF7", "WEAKNESSES", [
                "• Limited counterparty credit analytics capability",
                "• Concentrated buyer exposure (feedlot segment)",
                "• Manual credit decision processes",
                "• Working capital thinly hedged against rate shocks",
                "• Technology infrastructure lagging peers"]),
            (0, 0, 5, 5, "#FFF4CE", "OPPORTUNITIES", [
                "• Trade credit insurance market maturation",
                "• Fintech partnerships for credit scoring",
                "• Value-added services (hedging advisory)",
                "• Data monetization (market intelligence)",
                "• Geographic expansion into underserved regions"]),
            (5, 0, 5, 5, "#FDE7E9", "THREATS", [
                "• Record farm bankruptcies (315+ in 2025)",
                "• Packer consolidation squeezing margins",
                "• Interest rate volatility",
                "• Regulatory intensification (P&SA enforcement)",
                "• Commodity price shocks (cattle cycle, ASF)"]),
        ]
        for x, y, w, h, color, title, items in quads:
            ax.add_patch(Rectangle((x, y), w, h, facecolor=color, edgecolor="white", linewidth=3))
            ax.text(x + w/2, y + h - 0.5, title, ha="center", fontsize=12,
                    fontweight="bold", color=self.config.NAVY_HEX)
            for i, item in enumerate(items):
                ax.text(x + 0.3, y + h - 1.3 - i*0.65, item, fontsize=8.5, color=self.config.DARK_GRAY_HEX)
        ax.set_title("PLMA SWOT Analysis", fontsize=14, fontweight="bold",
                     color=self.config.NAVY_HEX, pad=15)
        return self._finalize(fig)

    def scenario_waterfall(self):
        scenarios = ["Soft Landing", "Base Case", "Prolonged Stress", "Systemic Crisis"]
        revenue_impact = [2, -3, -12, -25]
        credit_losses = [-5, -7, -16, -38]
        net_margin = [1.5, -0.5, -3.5, -9.2]
        x = np.arange(len(scenarios))
        width = 0.25
        fig, ax = plt.subplots(figsize=(10, 5.5))
        b1 = ax.bar(x - width, revenue_impact, width, label="Revenue Impact (%)",
                    color=self.config.BLUE_HEX, edgecolor="white")
        b2 = ax.bar(x, credit_losses, width, label="Credit Losses ($M)",
                    color=self.config.RED_HEX, edgecolor="white")
        b3 = ax.bar(x + width, net_margin, width, label="Net Margin (%)",
                    color=self.config.NAVY_HEX, edgecolor="white")
        ax.axhline(y=0, color="black", linewidth=0.8)
        ax.set_xticks(x)
        ax.set_xticklabels(scenarios, fontsize=10)
        ax.set_ylabel("Impact", fontsize=10)
        ax.set_title("Scenario Analysis: Financial Impact by Pathway",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.legend(loc="lower left", frameon=False)
        ax.grid(True, axis="y", alpha=0.3)
        ax.set_axisbelow(True)
        for bars in [b1, b2, b3]:
            for bar in bars:
                h = bar.get_height()
                va = "bottom" if h >= 0 else "top"
                ax.text(bar.get_x() + bar.get_width()/2, h, f"{h:.1f}",
                        ha="center", va=va, fontsize=8, color=self.config.DARK_GRAY_HEX)
        return self._finalize(fig)

    def implementation_timeline(self):
        workstreams = [
            ("Counterparty risk dashboard deployment", 0, 6, self.config.NAVY_HEX),
            ("Credit scoring model (v1 build + pilot)", 0, 9, self.config.NAVY_HEX),
            ("Trade credit insurance program launch", 3, 6, self.config.NAVY_HEX),
            ("Tiered credit limits policy", 0, 4, self.config.NAVY_HEX),
            ("Hedging overlay program", 6, 18, self.config.BLUE_HEX),
            ("Credit scoring v2 (ML-enhanced)", 12, 24, self.config.BLUE_HEX),
            ("Geographic diversification", 18, 42, self.config.BLUE_HEX),
            ("Value-added services launch", 24, 36, self.config.BLUE_HEX),
            ("Technology platform modernization", 12, 48, self.config.BLUE_HEX),
            ("Product line expansion (advisory)", 36, 60, self.config.AMBER_HEX),
            ("Strategic M&A / partnership eval", 48, 84, self.config.AMBER_HEX),
            ("Cooperative capital restructuring", 60, 120, self.config.AMBER_HEX),
        ]
        fig, ax = plt.subplots(figsize=(14, 7))
        for i, (name, start, end, color) in enumerate(workstreams):
            ax.barh(i, end - start, left=start, color=color, edgecolor="white",
                    linewidth=1.5, height=0.7)
        # Use y-axis labels for workstream names to prevent text clipping
        ax.set_yticks(range(len(workstreams)))
        ax.set_yticklabels([name for name, _, _, _ in workstreams],
                           fontsize=7.5, fontweight="bold")
        for i, (_, _, _, color) in enumerate(workstreams):
            ax.get_yticklabels()[i].set_color(color)
        ax.axvline(x=12, color=self.config.DARK_GRAY_HEX, linestyle="--", alpha=0.5)
        ax.axvline(x=60, color=self.config.DARK_GRAY_HEX, linestyle="--", alpha=0.5)
        ax.text(6, len(workstreams)+0.3, "HORIZON 1\n(Year 1)", ha="center", fontsize=10,
                fontweight="bold", color=self.config.NAVY_HEX)
        ax.text(36, len(workstreams)+0.3, "HORIZON 2\n(Years 2-5)", ha="center", fontsize=10,
                fontweight="bold", color=self.config.BLUE_HEX)
        ax.text(90, len(workstreams)+0.3, "HORIZON 3\n(Years 6-10)", ha="center", fontsize=10,
                fontweight="bold", color=self.config.AMBER_HEX)
        ax.set_xlabel("Months from Inception", fontsize=10)
        ax.set_xlim(0, 122)
        ax.set_ylim(-0.8, len(workstreams)+1.2)
        ax.set_title("Implementation Roadmap: McKinsey Three Horizons Framework",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=25)
        ax.grid(True, axis="x", alpha=0.3)
        ax.set_axisbelow(True)
        fig.subplots_adjust(left=0.30, right=0.95)
        return self._finalize(fig)

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
        # breakeven
        for i in range(len(cum)-1):
            if cum[i] < 0 <= cum[i+1]:
                frac = -cum[i] / (cum[i+1] - cum[i])
                be_year = years[i] + frac
                ax.axvline(x=be_year, color=self.config.AMBER_HEX, linestyle="--", linewidth=2)
                ax.text(be_year + 0.1, ax.get_ylim()[1] * 0.3,
                        f"Breakeven\nYear {be_year:.1f}", fontsize=10, fontweight="bold",
                        color=self.config.AMBER_HEX)
                break
        for x, y in zip(years, cum):
            ax.annotate(f"${y:.1f}M", (x, y), textcoords="offset points",
                        xytext=(0, 10), ha="center", fontsize=9, color=self.config.DARK_GRAY_HEX)
        ax.set_xlabel("Year", fontsize=11)
        ax.set_ylabel("Cumulative Net Cash Flow ($M)", fontsize=11)
        ax.set_title("10-Year Cumulative ROI: Credit Risk Mitigation Program",
                     fontsize=12, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.legend(loc="upper left", frameon=False)
        ax.grid(True, alpha=0.3)
        ax.set_axisbelow(True)
        return self._finalize(fig)

    def cost_benefit_waterfall(self):
        items = [
            ("Year 1\nInvestment", -1.715, self.config.RED_HEX),
            ("Credit Scoring\nBenefit", 1.80, self.config.GREEN_HEX),
            ("Trade Credit\nInsurance", 1.30, self.config.GREEN_HEX),
            ("Early Warning\nSystem", 0.65, self.config.GREEN_HEX),
            ("Hedging\nProgram", 0.80, self.config.GREEN_HEX),
            ("Ongoing\nCosts", -1.17, self.config.RED_HEX),
        ]
        fig, ax = plt.subplots(figsize=(11, 5.5))
        cumulative = 0
        positions = []
        for i, (name, val, color) in enumerate(items):
            if val >= 0:
                ax.bar(i, val, bottom=cumulative, color=color, edgecolor="white", linewidth=1.5)
            else:
                ax.bar(i, val, bottom=cumulative, color=color, edgecolor="white", linewidth=1.5)
            label_y = cumulative + val/2
            ax.text(i, label_y, f"${val:+.2f}M", ha="center", va="center",
                    fontsize=10, fontweight="bold", color="white")
            positions.append(i)
            cumulative += val
        # Net bar
        ax.bar(len(items), cumulative, color=self.config.NAVY_HEX, edgecolor="white", linewidth=1.5)
        ax.text(len(items), cumulative/2, f"${cumulative:+.2f}M\nNet", ha="center", va="center",
                fontsize=11, fontweight="bold", color="white")
        ax.axhline(y=0, color="black", linewidth=0.8)
        labels = [item[0] for item in items] + ["Net Annual\nImpact"]
        ax.set_xticks(range(len(labels)))
        ax.set_xticklabels(labels, fontsize=9)
        ax.set_ylabel("Annual Impact ($M)", fontsize=10)
        ax.set_title("Cost-Benefit Waterfall: Steady-State Annual Economics",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.grid(True, axis="y", alpha=0.3)
        ax.set_axisbelow(True)
        return self._finalize(fig)

    def sensitivity_tornado(self):
        tornado, base_npv = self.model.sensitivity()
        tornado = tornado[:6]
        base_m = base_npv / 1_000_000
        fig, ax = plt.subplots(figsize=(10, 5.5))
        y_pos = np.arange(len(tornado))
        for i, t in enumerate(tornado):
            low_m = t["low"] / 1_000_000 - base_m
            high_m = t["high"] / 1_000_000 - base_m
            ax.barh(i, low_m, color=self.config.RED_HEX, edgecolor="white",
                    height=0.6, alpha=0.85)
            ax.barh(i, high_m, color=self.config.GREEN_HEX, edgecolor="white",
                    height=0.6, alpha=0.85)
        # Render once to establish axes scale before placing labels
        fig.canvas.draw()
        pad_pts = 25  # padding from bar end, in points
        for i, t in enumerate(tornado):
            low_m = t["low"] / 1_000_000 - base_m
            high_m = t["high"] / 1_000_000 - base_m
            bar_range = abs(high_m - low_m)
            # For very small bars, shift labels further and stagger vertically
            if bar_range < 3:
                ax.annotate(f"${t['low']/1e6:.1f}M", xy=(low_m, i),
                            xytext=(-pad_pts * 2, -6), textcoords="offset points",
                            ha="right", va="center", fontsize=8.5, fontweight="bold",
                            clip_on=False)
                ax.annotate(f"${t['high']/1e6:.1f}M", xy=(high_m, i),
                            xytext=(pad_pts * 2, 6), textcoords="offset points",
                            ha="left", va="center", fontsize=8.5, fontweight="bold",
                            clip_on=False)
            else:
                ax.annotate(f"${t['low']/1e6:.1f}M", xy=(low_m, i),
                            xytext=(-pad_pts, 0), textcoords="offset points",
                            ha="right", va="center", fontsize=9, fontweight="bold",
                            clip_on=False)
                ax.annotate(f"${t['high']/1e6:.1f}M", xy=(high_m, i),
                            xytext=(pad_pts, 0), textcoords="offset points",
                            ha="left", va="center", fontsize=9, fontweight="bold",
                            clip_on=False)
        ax.set_yticks(y_pos)
        ax.set_yticklabels([t["variable"] for t in tornado], fontsize=10)
        ax.invert_yaxis()
        ax.axvline(x=0, color=self.config.NAVY_HEX, linewidth=2)
        ax.margins(x=0.25)
        ax.set_xlabel(f"Change in 10-Year NPV vs. Base Case (${base_m:.1f}M)", fontsize=10)
        ax.set_title("Sensitivity Analysis: Tornado Diagram (NPV Impact)",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.grid(True, axis="x", alpha=0.3)
        ax.set_axisbelow(True)
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
        # Normal / body
        normal = self.doc.styles["Normal"]
        normal.font.name = self.config.BODY_FONT
        normal.font.size = Pt(10.5)
        normal.font.color.rgb = self.config.DARK_GRAY

        # Headings
        for level, size, color in [(1, 20, self.config.NAVY),
                                    (2, 15, self.config.NAVY),
                                    (3, 12, self.config.BLUE)]:
            h = self.doc.styles[f"Heading {level}"]
            h.font.name = self.config.HEADING_FONT
            h.font.size = Pt(size)
            h.font.color.rgb = color
            h.font.bold = True

        # Title
        title = self.doc.styles["Title"]
        title.font.name = self.config.HEADING_FONT
        title.font.size = Pt(32)
        title.font.color.rgb = self.config.NAVY
        title.font.bold = True

    def _setup_page_layout(self):
        section = self.doc.sections[0]
        section.top_margin = Inches(0.9)
        section.bottom_margin = Inches(0.9)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

        # Footer with page number
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp.add_run("CONFIDENTIAL — Prepared for R. Johansen, CFO, PLMA   |   Page ").font.size = Pt(9)
        for run in fp.runs:
            run.font.name = self.config.BODY_FONT
            run.font.color.rgb = self.config.MID_GRAY
            run.font.size = Pt(9)
        add_page_number_field(fp)

        # Header
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run("Counterparty Credit Risk Assessment  •  April 2026")
        hrun.font.name = self.config.BODY_FONT
        hrun.font.size = Pt(9)
        hrun.font.color.rgb = self.config.MID_GRAY
        hrun.italic = True

    # ------------ utility ------------
    def _para(self, text, style=None, size=None, bold=False, italic=False,
              color=None, align=None, space_after=6, space_before=0, line=1.2):
        p = self.doc.add_paragraph(style=style) if style else self.doc.add_paragraph()
        run = p.add_run(text)
        run.font.name = self.config.BODY_FONT
        if size:
            run.font.size = Pt(size)
        run.bold = bold
        run.italic = italic
        if color:
            run.font.color.rgb = color
        if align is not None:
            p.alignment = align
        set_paragraph_spacing(p, before=space_before, after=space_after, line=line)
        return p

    def _h1(self, text):
        self.doc.add_page_break()
        p = self.doc.add_heading(text, level=1)
        set_paragraph_spacing(p, before=0, after=12)
        # Horizontal rule via bottom border
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "12")
        bottom.set(qn("w:space"), "4")
        bottom.set(qn("w:color"), self.config.NAVY_SHADE)
        pBdr.append(bottom)
        pPr.append(pBdr)
        return p

    def _h2(self, text):
        p = self.doc.add_heading(text, level=2)
        set_paragraph_spacing(p, before=14, after=6)
        return p

    def _h3(self, text):
        p = self.doc.add_heading(text, level=3)
        set_paragraph_spacing(p, before=10, after=4)
        return p

    def _body(self, text):
        return self._para(text, size=10.5, space_after=8, line=1.25)

    def _bullet(self, text, indent=0):
        p = self.doc.add_paragraph(style="List Bullet")
        run = p.add_run(text)
        run.font.name = self.config.BODY_FONT
        run.font.size = Pt(10.5)
        run.font.color.rgb = self.config.DARK_GRAY
        set_paragraph_spacing(p, after=3, line=1.2)
        if indent:
            p.paragraph_format.left_indent = Inches(0.25 * (indent + 1))
        return p

    def _bullet_lead(self, lead, text):
        p = self.doc.add_paragraph(style="List Bullet")
        r1 = p.add_run(lead + " ")
        r1.bold = True
        r1.font.name = self.config.BODY_FONT
        r1.font.size = Pt(10.5)
        r1.font.color.rgb = self.config.NAVY
        r2 = p.add_run(text)
        r2.font.name = self.config.BODY_FONT
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = self.config.DARK_GRAY
        set_paragraph_spacing(p, after=4, line=1.2)
        return p

    def _cite(self, paragraph, num):
        """Append a superscript citation number to a paragraph."""
        r = paragraph.add_run(f" [{num}]")
        r.font.superscript = True
        r.font.size = Pt(8)
        r.font.color.rgb = self.config.BLUE
        return r

    def _chart(self, chart_bytes, caption=None, width=6.5):
        self.doc.add_picture(chart_bytes, width=Inches(width))
        last = self.doc.paragraphs[-1]
        last.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if caption:
            p = self.doc.add_paragraph()
            r = p.add_run(f"Exhibit: {caption}")
            r.font.name = self.config.BODY_FONT
            r.font.size = Pt(9)
            r.italic = True
            r.font.color.rgb = self.config.MID_GRAY
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_paragraph_spacing(p, after=12)

    def _callout(self, title, text, kind="insight"):
        color_map = {
            "insight": (self.config.NAVY_SHADE, self.config.LIGHT_BLUE_SHADE),
            "warning": ("E81123", self.config.RED_SHADE),
            "finding": ("FFB900", self.config.AMBER_SHADE),
        }
        border_color, fill = color_map.get(kind, color_map["insight"])
        table = self.doc.add_table(rows=1, cols=1)
        cell = table.cell(0, 0)
        set_cell_shading(cell, fill)
        set_cell_border(cell,
            left={"val": "single", "sz": "36", "color": border_color, "space": "0"},
            top={"val": "single", "sz": "4", "color": border_color, "space": "0"},
            bottom={"val": "single", "sz": "4", "color": border_color, "space": "0"},
            right={"val": "single", "sz": "4", "color": border_color, "space": "0"},
        )
        cell.width = Inches(6.5)
        # Clear default paragraph, add styled
        cell.paragraphs[0].text = ""
        p1 = cell.paragraphs[0]
        r1 = p1.add_run(title.upper())
        r1.bold = True
        r1.font.name = self.config.HEADING_FONT
        r1.font.size = Pt(9)
        r1.font.color.rgb = self.config.NAVY if kind == "insight" else (
            self.config.RED if kind == "warning" else self.config.AMBER)
        set_paragraph_spacing(p1, after=3)
        p2 = cell.add_paragraph()
        r2 = p2.add_run(text)
        r2.font.name = self.config.BODY_FONT
        r2.font.size = Pt(10)
        r2.font.color.rgb = self.config.DARK_GRAY
        set_paragraph_spacing(p2, after=4, line=1.2)
        # spacer
        self._para("", space_after=8)

    def _table(self, headers, rows, col_widths=None, header_fill=None):
        header_fill = header_fill or self.config.NAVY_SHADE
        table = self.doc.add_table(rows=1 + len(rows), cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        # Header row
        for i, h in enumerate(headers):
            cell = table.rows[0].cells[i]
            set_cell_shading(cell, header_fill)
            cell.paragraphs[0].text = ""
            p = cell.paragraphs[0]
            r = p.add_run(h)
            r.bold = True
            r.font.name = self.config.HEADING_FONT
            r.font.size = Pt(9.5)
            r.font.color.rgb = self.config.WHITE
            set_paragraph_spacing(p, after=2)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        # Data rows
        for ri, row in enumerate(rows):
            for ci, val in enumerate(row):
                cell = table.rows[ri+1].cells[ci]
                if ri % 2 == 1:
                    set_cell_shading(cell, self.config.VERY_LIGHT_GRAY_SHADE)
                cell.paragraphs[0].text = ""
                p = cell.paragraphs[0]
                r = p.add_run(str(val))
                r.font.name = self.config.BODY_FONT
                r.font.size = Pt(9.5)
                r.font.color.rgb = self.config.DARK_GRAY
                set_paragraph_spacing(p, after=2)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        # Column widths
        if col_widths:
            for row in table.rows:
                for i, w in enumerate(col_widths):
                    row.cells[i].width = Inches(w)
        # Spacer after table
        self._para("", space_after=8)
        return table

    # ============ SECTIONS ============
    def _add_cover_page(self):
        # Top spacer
        for _ in range(4):
            self._para("", space_after=2)
        # Confidential tag
        p = self.doc.add_paragraph()
        r = p.add_run("CONFIDENTIAL  •  CLIENT DELIVERABLE")
        r.font.name = self.config.HEADING_FONT
        r.font.size = Pt(10)
        r.font.color.rgb = self.config.RED
        r.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(p, after=40)

        # Main title
        p = self.doc.add_paragraph()
        r = p.add_run("Counterparty Credit Risk")
        r.font.name = self.config.HEADING_FONT
        r.font.size = Pt(36)
        r.font.color.rgb = self.config.NAVY
        r.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(p, after=6)

        p = self.doc.add_paragraph()
        r = p.add_run("Assessment & Strategic Mitigation Framework")
        r.font.name = self.config.HEADING_FONT
        r.font.size = Pt(22)
        r.font.color.rgb = self.config.BLUE
        r.bold = False
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(p, after=40)

        # Accent bar (use table)
        tbl = self.doc.add_table(rows=1, cols=1)
        cell = tbl.cell(0, 0)
        set_cell_shading(cell, self.config.NAVY_SHADE)
        cell.paragraphs[0].text = ""
        p = cell.paragraphs[0]
        r = p.add_run(" ")
        r.font.size = Pt(2)
        cell.width = Inches(6.5)
        self._para("", space_after=30)

        # Subject
        p = self.doc.add_paragraph()
        r = p.add_run("Producers Livestock Marketing Association")
        r.font.name = self.config.HEADING_FONT
        r.font.size = Pt(16)
        r.font.color.rgb = self.config.DARK_GRAY
        r.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(p, after=4)

        p = self.doc.add_paragraph()
        r = p.add_run("Prepared for Ryan Johansen, Chief Financial Officer")
        r.font.name = self.config.BODY_FONT
        r.font.size = Pt(12)
        r.font.color.rgb = self.config.MID_GRAY
        r.italic = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(p, after=60)

        p = self.doc.add_paragraph()
        r = p.add_run("APRIL 2026")
        r.font.name = self.config.HEADING_FONT
        r.font.size = Pt(11)
        r.font.color.rgb = self.config.NAVY
        r.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(p, after=6)

        p = self.doc.add_paragraph()
        r = p.add_run("Strategic Advisory  |  Risk Management Practice")
        r.font.name = self.config.BODY_FONT
        r.font.size = Pt(9)
        r.font.color.rgb = self.config.MID_GRAY
        r.italic = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def _add_toc(self):
        self.doc.add_page_break()
        p = self.doc.add_heading("Table of Contents", level=1)
        set_paragraph_spacing(p, after=16)
        entries = [
            ("1. Executive Summary", "3"),
            ("2. Industry Context & Macroeconomic Environment", "6"),
            ("3. Counterparty Credit Risk Deep Dive", "11"),
            ("4. Comprehensive Risk Assessment", "17"),
            ("5. Strategic Framework Analysis", "24"),
            ("6. Mitigation Strategy & Implementation Roadmap", "30"),
            ("7. Financial Model: Cost-Benefit Analysis", "37"),
            ("8. Governance & Monitoring", "42"),
            ("9. Conclusion & Next Steps", "45"),
            ("Appendix A: Detailed Data Tables", "47"),
            ("Appendix B: Methodology Notes", "50"),
            ("Appendix C: Glossary of Terms", "52"),
            ("Appendix D: Citations", "54"),
        ]
        for title, page in entries:
            p = self.doc.add_paragraph()
            tab_stops = p.paragraph_format.tab_stops
            tab_stops.add_tab_stop(Inches(6.3), alignment=2, leader=1)  # right align with dots
            r1 = p.add_run(title)
            r1.font.name = self.config.BODY_FONT
            r1.font.size = Pt(11)
            r1.font.color.rgb = self.config.DARK_GRAY
            if not title[0].isdigit() and not title.startswith("Appendix"):
                r1.bold = True
            r2 = p.add_run(f"\t{page}")
            r2.font.name = self.config.BODY_FONT
            r2.font.size = Pt(11)
            r2.font.color.rgb = self.config.NAVY
            r2.bold = True
            set_paragraph_spacing(p, after=8, line=1.3)

    def _add_executive_summary(self):
        self._h1("1. Executive Summary")

        self._callout(
            "Situation",
            "PLMA faces the most acute counterparty credit risk environment in a generation. With 315 "
            "Chapter 12 farm bankruptcies filed in 2025—a 22% year-over-year increase—and U.S. farm "
            "sector debt reaching a record $624.7 billion, PLMA's buyer base (feedlots and processors) "
            "is experiencing compounding margin compression. Realized losses have already materialized "
            "on PLMA's balance sheet, and current trajectory suggests losses could more than double "
            "absent decisive mitigation.",
            kind="warning"
        )

        self._h2("1.1 Key Findings")

        self._bullet_lead(
            "Counterparty risk is concentrated and severe.",
            "Four buyer segments represent 100% of PLMA's estimated $200M working capital exposure. "
            "Independent feedlots (40% of exposure) carry an estimated 8% 12-month default probability; "
            "backgrounders carry 12%. Combined expected loss is $7.2M annually under base case.")

        self._bullet_lead(
            "The stress case is materially worse.",
            "Under a scenario where default rates double (consistent with the 2015–2016 cattle downturn "
            "and early-2026 bankruptcy acceleration), expected losses rise to $14.4M—roughly 24% of "
            "PLMA's estimated commission revenue.")

        self._bullet_lead(
            "PLMA currently lacks the analytical infrastructure to manage this risk.",
            "Credit decisioning is largely manual and relationship-driven; there is no quantitative "
            "counterparty scoring model, no tiered credit limit policy, and limited trade credit "
            "insurance coverage.")

        self._bullet_lead(
            "A $1.7M Year-1 investment generates ~$20M in 10-year NPV.",
            "The recommended mitigation program—credit scoring, trade credit insurance, early warning "
            "systems, and hedging overlays—pays back in under 12 months and delivers 10-year ROI of "
            "approximately 250% at an 8% discount rate.")

        self._bullet_lead(
            "Inaction is not a viable posture.",
            "Every month of delay costs approximately $600K in avoidable expected losses. The window "
            "to price and procure trade credit insurance on favorable terms is narrowing as "
            "underwriters reprice ag exposure.")

        self._h2("1.2 Recommended Actions")

        self._body(
            "We recommend a phased implementation structured around the McKinsey Three Horizons "
            "framework, scaled to PLMA's organizational capacity:")

        self._bullet_lead("Horizon 1 (0–12 months):",
            "Stand up counterparty risk dashboard, deploy v1 credit scoring model, implement tiered "
            "credit limits, and procure trade credit insurance for the top quartile of exposure.")
        self._bullet_lead("Horizon 2 (Years 2–5):",
            "Launch ML-enhanced credit scoring, systematic hedging overlay, geographic "
            "diversification, and value-added advisory services.")
        self._bullet_lead("Horizon 3 (Years 6–10):",
            "Execute strategic M&A/partnership strategy, expand product line into fee-based "
            "advisory, and restructure cooperative capital for resilience.")

        self._h2("1.3 Financial Impact Summary")

        self._table(
            ["Metric", "Value"],
            [
                ["Base Case Annual Expected Loss", "$7.21M"],
                ["Stress Case Annual Expected Loss", "$14.42M"],
                ["Year 1 Program Investment", "$1.72M"],
                ["Annual Ongoing Cost (Year 2+)", "$1.17M"],
                ["Annual Benefit (Year 2+)", "$4.55M"],
                ["10-Year NPV (8% discount rate)", "~$20.2M"],
                ["Payback Period", "~10 months"],
                ["10-Year Cumulative ROI", "~250%"],
            ],
            col_widths=[3.5, 2.5]
        )

        self._callout(
            "Urgency",
            "The cost of a 90-day delay is approximately $1.8M in avoidable expected losses plus "
            "escalating insurance premiums. We recommend convening the Risk Committee within 30 days "
            "to approve the Horizon 1 workstream and release of $1.72M in Year 1 funding.",
            kind="finding"
        )

    def _add_section_2_industry(self):
        self._h1("2. Industry Context & Macroeconomic Environment")

        self._body(
            "To properly frame PLMA's counterparty credit risk exposure, we must situate the "
            "cooperative within the broader U.S. agricultural finance landscape. Five macroeconomic "
            "forces are converging to create the most challenging credit environment livestock "
            "marketing cooperatives have faced since the 1980s farm crisis: (1) record farm sector "
            "debt, (2) accelerating bankruptcy filings, (3) persistent interest rate pressure, (4) "
            "structural margin compression in downstream segments, and (5) intensifying regulatory "
            "scrutiny.")

        self._h2("2.1 The U.S. Agricultural Lending Landscape")

        p = self._body(
            "Total U.S. farm sector debt reached an all-time nominal high of $624.7 billion in 2025, "
            "per USDA Economic Research Service estimates, with real estate debt comprising ~65% and "
            "non-real-estate operating debt the remainder.")
        self._cite(p, 1)
        p = self._body(
            "Over the past decade, farm debt has grown at a compound annual rate of 4.9%, outpacing "
            "both farm income growth and general inflation. Debt-to-asset ratios in the livestock "
            "sector now exceed 18%, the highest level since 2002.")
        self._cite(p, 5)

        self._chart(self.charts.farm_debt_growth(),
                    caption="Total U.S. Farm Sector Debt has grown 62% since 2015, reaching $624.7B in 2025.")

        self._h2("2.2 Farm Bankruptcy Trends")

        p = self._body(
            "Chapter 12 farm bankruptcy filings—the primary legal mechanism for distressed producers—"
            "rose to 315 in 2025, a 22% increase over 2024 and the highest level since 2021. "
            "American Bankruptcy Institute data suggests filings are accelerating in Q1 2026, with "
            "Midwest and Plains states (PLMA's core footprint) disproportionately represented.")
        self._cite(p, 2)

        self._chart(self.charts.farm_bankruptcy_trend(),
                    caption="Chapter 12 bankruptcies are inflecting upward, with 2026 projected at 380+ filings.")

        self._callout(
            "Key Insight",
            "Bankruptcy filings are a lagging indicator. For every Chapter 12 filing, industry data "
            "suggests 4–6 producers exit the business without a formal bankruptcy proceeding—through "
            "negotiated workouts, asset liquidation, or generational exit. The true distress rate is "
            "materially higher than filing data implies.",
            kind="insight"
        )

        self._h2("2.3 Livestock Market Dynamics")

        self._body(
            "The cattle cycle currently sits at a precarious inflection point. The U.S. cow herd "
            "(88.2M head as of January 2026) is at a 75-year low, driving feeder cattle prices to "
            "record levels and compressing feedlot margins. Simultaneously, pork producers face "
            "oversupply risk as expansion from 2022–2024 reaches market weight, with persistent ASF "
            "(African Swine Fever) concerns threatening export demand.")

        self._h3("2.3.1 Feedlot Margin Compression")
        p = self._body(
            "Independent feedlots—PLMA's largest buyer segment by exposure—have operated at negative "
            "closeout margins for 11 of the past 14 months, averaging -$85/head in Q1 2026 per "
            "Sterling Marketing data. Feedlots have drawn down equity cushions built during "
            "2022–2023 profitability, and working capital lines are increasingly constrained.")
        self._cite(p, 3)

        self._h3("2.3.2 Packer Consolidation")
        p = self._body(
            "The 'Big Four' packers (JBS, Tyson, Cargill, National Beef) now control ~85% of fed "
            "cattle slaughter capacity. While this concentration creates systemic counterparty "
            "risk, it also drives pricing dynamics that squeeze PLMA's seller-side producers and "
            "fuel regulatory attention under the Packers & Stockyards Act.")
        self._cite(p, 10)

        self._h2("2.4 Buyer Segment Analysis")

        self._table(
            ["Segment", "Share of Exposure", "Margin Pressure", "Default Trajectory"],
            [
                ["Independent Feedlots", "40%", "Severe", "Rising"],
                ["Small/Mid Processors", "25%", "High", "Rising"],
                ["Large Packers", "20%", "Moderate", "Stable"],
                ["Backgrounders/Stockers", "15%", "Severe", "Rising Rapidly"],
            ],
            col_widths=[2.2, 1.5, 1.3, 1.5]
        )

        self._h2("2.5 Regulatory Environment")

        p = self._body(
            "USDA's Agricultural Marketing Service (AMS) has intensified Packers & Stockyards Act "
            "enforcement under recently finalized rules addressing unfair practices, prompt payment, "
            "and market transparency. Livestock marketing cooperatives sit at the intersection of "
            "seller protections (producers) and buyer accountability (feedlots/packers), creating "
            "elevated compliance obligations.")
        self._cite(p, 10)
        p = self._body(
            "Additionally, CFTC oversight of commodity hedging activities and state-level livestock "
            "dealer licensing regulations create a complex compliance matrix that PLMA must navigate.")
        self._cite(p, 24)

    def _add_section_3_credit_risk(self):
        self._h1("3. Counterparty Credit Risk Deep Dive")

        self._body(
            "Counterparty credit risk is the most acute financial threat to PLMA's balance sheet. "
            "This section quantifies PLMA's exposure, decomposes risk by segment and pathway, and "
            "identifies contagion vectors that could amplify losses in a stress scenario.")

        self._h2("3.1 PLMA's Counterparty Exposure Model")

        self._body(
            "As a cooperative livestock marketing entity, PLMA occupies a two-sided credit position: "
            "it advances funds to producers (sellers) against consigned livestock, and extends "
            "trade credit to buyers (feedlots, processors, packers) on sold livestock. This "
            "dual-sided exposure creates a credit sandwich: PLMA owes producers whether or not "
            "buyers pay, and bears the default risk on both sides.")

        self._body(
            "Based on industry benchmarks for a cooperative of PLMA's estimated scale "
            "(~$3.2B annual transaction volume, ~$60M in commission revenue), we estimate "
            "PLMA's aggregate working capital exposure at any given time at approximately $200M, "
            "distributed across four primary counterparty segments.")

        self._h2("3.2 Seller-Side Risk (Producer/Rancher Default)")

        self._body(
            "Producer-side risk manifests in three primary pathways: (i) bankruptcy or forced "
            "liquidation during the consignment window, (ii) delivery failure or mis-described "
            "livestock creating buyer rejection risk, and (iii) margin call obligations if PLMA "
            "has advanced funds against consigned cattle that subsequently decline in market value.")

        self._body(
            "With 315 Chapter 12 filings in 2025 and an accelerating trajectory, seller-side risk "
            "is material. However, because PLMA typically holds physical possession or legal title "
            "during the consignment window, recovery rates on producer defaults are higher than "
            "on buyer defaults—estimated at 55–70% recovery.")

        self._h2("3.3 Buyer-Side Risk (Feedlot/Processor Default)")

        self._body(
            "Buyer-side risk is more acute and less recoverable. When a buyer defaults on payment "
            "for livestock already delivered and slaughtered or placed, PLMA has limited recourse. "
            "Typical recovery rates on unsecured buyer trade claims in livestock bankruptcies "
            "range from 15–45% depending on segment and security position.")

        self._chart(self.charts.buyer_default_probability(),
                    caption="Backgrounders and independent feedlots carry the highest default probabilities.")

        self._h2("3.4 Concentration Risk Analysis")

        self._callout(
            "Key Finding",
            "PLMA's top 10 buyers represent approximately 55–65% of total counterparty exposure. "
            "A single large buyer default (estimated top buyer exposure: $25–35M) could generate "
            "losses exceeding 30% of annual commission revenue. This concentration is the single "
            "largest source of tail risk on the balance sheet.",
            kind="finding"
        )

        self._body("Concentration risk manifests across three dimensions:")
        self._bullet_lead("Buyer concentration:",
            "Top-10 buyer exposure exceeds prudent single-counterparty limits (typically 10% of "
            "equity or 15% of receivables per banking best practice).")
        self._bullet_lead("Geographic concentration:",
            "PLMA's operational footprint is concentrated in multi-state Plains/Mountain West "
            "regions, creating correlated exposure to regional drought, blizzards, and disease.")
        self._bullet_lead("Segment concentration:",
            "65% of exposure sits in feedlot and backgrounder segments, both of which face "
            "simultaneous margin compression—this is not diversified risk.")

        self._h2("3.5 Contagion and Systemic Risk Pathways")

        self._chart(self.charts.risk_heat_map(),
                    caption="Counterparty buyer bankruptcy, producer default, and concentration risk occupy the red zone.")

        p = self._body(
            "Contagion pathways warrant specific attention. A single large feedlot bankruptcy "
            "typically triggers: (1) emergency liquidation of cattle inventory, depressing "
            "regional fed cattle prices, (2) tightening of credit terms by packers to surviving "
            "feedlots, (3) margin calls cascading through the supply chain, and (4) "
            "second-order producer distress as feeder cattle demand softens.")
        self._cite(p, 8)

    def _add_section_4_risk_assessment(self):
        self._h1("4. Comprehensive Risk Assessment")

        self._body(
            "Counterparty credit risk does not exist in isolation. It is correlated with, and "
            "amplified by, a constellation of market, operational, and regulatory risks. This "
            "section applies an enterprise risk management lens to identify the full risk surface "
            "PLMA must manage, aligned to ISO 31000 risk categorization principles.")
        p = self.doc.paragraphs[-1]
        self._cite(p, 14)

        self._h2("4.1 Cattle Price Risk")

        self._body(
            "Cattle price volatility translates directly into credit risk for both producers and "
            "buyers. The current cattle cycle sits near a peak, with feeder cattle futures at "
            "record nominal levels. Historical analysis of the 2014–2016 cycle suggests that peak-"
            "to-trough corrections can exceed 35%, creating severe distress in leveraged feedlot "
            "operations. PLMA's basis risk—the differential between cash and futures prices in "
            "local markets—adds an additional 3–8% volatility layer.")

        self._h2("4.2 Pork Market Risk")

        self._body(
            "While cattle dominate PLMA's exposure, pork segment risk deserves attention. ASF "
            "remains the largest tail risk: a confirmed U.S. outbreak would immediately halt "
            "pork exports (~25% of production), collapse domestic prices, and bankrupt a "
            "significant share of leveraged pork finishing operations. Preliminary BioSecurity "
            "Task Force modeling estimates a 72-hour price decline of 25–40% in an outbreak "
            "scenario.")

        self._h2("4.3 Feed and Input Commodity Risk")

        self._body(
            "Feed costs represent 60–70% of livestock production costs. Corn and soybean meal "
            "price spikes compress feedlot and pork producer margins, elevating default "
            "probabilities across PLMA's buyer base. The correlation matrix below quantifies "
            "these linkages.")

        self._chart(self.charts.commodity_correlation_heatmap(),
                    caption="Corn and soybean meal move together (0.72) and inversely to feeder cattle (-0.52).")

        self._h2("4.4 Interest Rate Risk")

        self._body(
            "Interest rates affect PLMA through two channels: (1) direct cost-of-carry on working "
            "capital exposure, and (2) indirect stress on leveraged buyers whose debt service "
            "burden rises with rates. With an estimated $200M in working capital exposure at "
            "any given time, a 100 bp rate increase translates to $1.5M in incremental annual "
            "cost-of-carry.")

        self._chart(self.charts.interest_rate_sensitivity(),
                    caption="PLMA's cost-of-carry rises linearly with rates while buyer financial stress rises exponentially.")

        self._h2("4.5 Regulatory & Compliance Risk")

        p = self._body(
            "Packers & Stockyards Act enforcement has intensified under USDA's AMS. Recent rule "
            "changes address: (i) undue preferences, (ii) prompt payment obligations (PLMA has "
            "24-hour payment requirements on some transactions), (iii) competitive injury "
            "standards, and (iv) market transparency reporting. Non-compliance carries both "
            "civil penalties and reputational risk.")
        self._cite(p, 10)

        self._body(
            "State-level livestock dealer licensing regimes add a further compliance matrix. "
            "PLMA operates under bonding requirements that scale with transaction volume—rising "
            "rates and default environments can pressure bond availability.")

        self._h2("4.6 Operational Risk")

        self._body("Operational risk categories include:")
        self._bullet_lead("Technology:",
            "Legacy systems create reconciliation errors, fraud exposure, and competitive "
            "disadvantage vs. fintech-native entrants.")
        self._bullet_lead("Personnel:",
            "Credit analyst bench depth is thin; key-person risk concentrated in senior "
            "relationship managers with tacit underwriting knowledge.")
        self._bullet_lead("Process:",
            "Manual credit decisions create inconsistency, slow response, and limited auditability.")
        self._bullet_lead("Cyber:",
            "Ag sector has been increasingly targeted by ransomware; industry benchmarks suggest "
            "$1.5–3M average incident cost for cooperatives of PLMA's size.")

    def _add_section_5_frameworks(self):
        self._h1("5. Strategic Framework Analysis")

        self._body(
            "This section applies five strategic frameworks to diagnose PLMA's competitive "
            "position, risk posture, and strategic options: Porter's Five Forces, SWOT, ISO 31000 "
            "Risk Appetite, Ansoff Matrix, and Scenario Planning. Together, these tools translate "
            "the risk analysis of prior sections into strategic direction.")

        self._h2("5.1 Porter's Five Forces — Livestock Marketing Industry")

        p = self._body(
            "The livestock marketing industry exhibits asymmetric power dynamics. Downstream "
            "buyer power is high (consolidated packers with Big-Four dominance), upstream "
            "supplier power is moderate-to-low (fragmented producer base), and rivalry is "
            "intense across cooperatives, independent auctions, and video/online marketplaces.")
        self._cite(p, 12)

        self._chart(self.charts.porters_five_forces(),
                    caption="Buyer power (4/5) and rivalry (4/5) are the dominant forces shaping industry economics.")

        self._table(
            ["Force", "Score (1–5)", "Key Drivers"],
            [
                ["Buyer Power (Packers)", "4 — High",
                 "Big-4 packer concentration; fed cattle captive supply arrangements"],
                ["Supplier Power (Producers)", "3 — Moderate",
                 "Fragmented producer base, but tight cattle numbers shift leverage"],
                ["Threat of New Entrants", "2 — Low",
                 "Regulatory barriers, relationship capital, bonding requirements"],
                ["Threat of Substitutes", "2 — Low",
                 "Direct sale and online auctions growing but limited at scale"],
                ["Competitive Rivalry", "4 — High",
                 "Multiple cooperatives, regional auctions, video marketers"],
            ],
            col_widths=[1.8, 1.0, 3.5]
        )

        self._h2("5.2 SWOT Analysis")

        self._chart(self.charts.swot_quadrant(),
                    caption="PLMA's strengths in trust and market access must offset weaknesses in analytics and technology.")

        self._h2("5.3 Risk Appetite Framework (ISO 31000)")

        self._body(
            "We recommend PLMA formalize a Risk Appetite Statement aligned to ISO 31000:2018 "
            "principles, articulating quantitative tolerances for each material risk category.")

        self._table(
            ["Risk Category", "Risk Tolerance", "Escalation Threshold"],
            [
                ["Single counterparty exposure", "≤10% of equity / $15M cap",
                 "Any exposure >$10M triggers Risk Committee review"],
                ["Top-10 concentration", "≤50% of total receivables",
                 "Breach triggers diversification action plan"],
                ["Segment concentration", "≤40% in any single segment",
                 "Breach triggers 12-month rebalancing plan"],
                ["Annual credit losses", "≤1.5% of commission revenue",
                 "Breach triggers underwriting tightening"],
                ["Cost of carry volatility", "≤$500K/year swing",
                 "Breach triggers hedging program review"],
            ],
            col_widths=[1.7, 2.0, 2.6]
        )

        self._h2("5.4 Ansoff Growth-Risk Matrix")

        self._body(
            "The Ansoff Matrix maps diversification options across existing vs. new markets and "
            "products, providing a structured view of strategic growth vectors that simultaneously "
            "reduce concentration risk.")

        self._table(
            ["", "Existing Products", "New Products"],
            [
                ["Existing Markets",
                 "Market Penetration: Deeper credit services to existing members; tiered pricing",
                 "Product Development: Hedging advisory, insurance brokerage, risk analytics-as-a-service"],
                ["New Markets",
                 "Market Development: Geographic expansion into underserved regions",
                 "Diversification: Technology platform licensing, fee-based advisory"],
            ],
            col_widths=[1.2, 2.6, 2.6]
        )

        self._h2("5.5 Scenario Planning Matrix")

        self._body(
            "We model four scenarios built on two axes: (i) farm bankruptcy rate trajectory, and "
            "(ii) cattle cycle recovery timing. Scenarios are designed to stress-test strategy "
            "and calibrate financial provisioning.")

        self._chart(self.charts.scenario_waterfall(),
                    caption="Systemic Crisis scenario implies -9.2% net margin and $38M in credit losses.")
        p = self.doc.paragraphs[-1]
        self._cite(p, 22)

    def _add_section_6_mitigation(self):
        self._h1("6. Mitigation Strategy & Implementation Roadmap")

        self._body(
            "We propose a four-pillar mitigation architecture, sequenced across the McKinsey Three "
            "Horizons framework. Horizon 1 focuses on stabilizing the current exposure through "
            "analytical infrastructure and insurance. Horizon 2 builds strategic capabilities. "
            "Horizon 3 transforms PLMA's business model to generate resilience through diversification.")

        self._chart(self.charts.implementation_timeline(),
                    caption="12 workstreams sequenced across 10 years, mapped to Three Horizons.")

        self._h2("6.1 Credit Risk Mitigation Toolkit (Horizon 1)")

        self._h3("6.1.1 Counterparty Credit Scoring Model")
        self._body(
            "Deploy a v1 quantitative credit scoring model within 9 months, using standard "
            "financial ratios (current ratio, debt service coverage, interest coverage), payment "
            "history, industry benchmarks, and qualitative overlays. Score counterparties on a "
            "1–10 scale mapped to probability-of-default tiers.")

        self._bullet_lead("Build cost:", "$350K (v1), $275K analytics platform")
        self._bullet_lead("Data sources:", "Financial statements, Paynet payment data, S&P Market Intelligence")
        self._bullet_lead("Refresh cadence:", "Quarterly scoring, monthly payment monitoring")

        self._h3("6.1.2 Tiered Credit Limits")
        self._body(
            "Move from relationship-based to quantitative credit limits. Set exposure caps as a "
            "function of credit score, counterparty equity, and PLMA's single-counterparty limit. "
            "Require Risk Committee approval for limits above $10M.")

        self._h3("6.1.3 Trade Credit Insurance Program")
        self._body(
            "Procure trade credit insurance covering the top 40–50% of buyer exposure (highest-"
            "concentration accounts). Typical premium: 30–60 bps of insured exposure. Net of "
            "deductibles and coverage gaps, estimated net loss avoidance: $1.3M/year.")

        self._callout(
            "Implementation Note",
            "Trade credit insurance markets are hardening rapidly as ag exposures deteriorate. "
            "We recommend PLMA engage 3+ carriers (Euler Hermes/Allianz Trade, Atradius, Coface) "
            "for competitive quotes within 60 days, before further premium escalation.",
            kind="insight"
        )

        self._h3("6.1.4 Payment Acceleration & Escrow")
        self._body(
            "Negotiate shorter payment terms with top-risk buyers (2–5 days vs. 7–10 days "
            "standard). Where appropriate, require escrow deposits or letters of credit from "
            "buyers exceeding concentration thresholds. This reduces EAD (Exposure At Default) "
            "by 30–50% for applicable counterparties.")

        self._h2("6.2 Market Risk Hedging Program (Horizon 2)")

        self._body(
            "Implement a systematic hedging overlay on PLMA's aggregate cash-market exposure "
            "during the consignment-to-settlement window. Use CME live cattle, feeder cattle, "
            "and lean hog futures to hedge inventory price risk.")

        self._bullet_lead("Target hedge ratio:", "60–80% of inventory exposure")
        self._bullet_lead("Basis risk management:", "Weekly basis monitoring, option overlays where basis volatility exceeds historical norms")
        self._bullet_lead("Estimated annual benefit:", "$800K in market risk reduction")

        self._h2("6.3 Operational Resilience (Horizon 1–2)")

        self._h3("6.3.1 Counterparty Monitoring Dashboard")
        self._body(
            "Build real-time dashboard tracking: exposure by counterparty, credit score trends, "
            "payment timeliness, public financial distress indicators (UCC filings, litigation, "
            "rating downgrades). Automate alerts when exposure or risk indicators breach thresholds.")

        self._h3("6.3.2 Early Warning System")
        self._body(
            "Systematic monitoring of leading indicators: (i) day sales outstanding trends, (ii) "
            "check kiting or slow-pay behavior, (iii) public records (liens, judgments), (iv) "
            "industry intelligence. Empirically, such systems enable 30–90 days of advance warning "
            "before default, enabling exposure reduction actions.")

        self._h2("6.4 Diversification Strategy (Horizons 2–3)")

        self._h3("6.4.1 Geographic Diversification")
        self._body(
            "Evaluate expansion into underserved regional markets through greenfield branches or "
            "acquisition of smaller regional cooperatives. Target: reduce regional concentration to "
            "<40% in any single USDA reporting region within 5 years.")

        self._h3("6.4.2 Product Line Expansion")
        self._body(
            "Launch fee-based advisory services: hedging advisory, insurance brokerage, credit "
            "analytics for member producers. These generate diversified, non-credit revenue "
            "streams while deepening member relationships.")

        self._h2("6.5 1-Year / 5-Year / 10-Year Plans at a Glance")

        self._h3("Year 1 Priorities (Horizon 1)")
        self._bullet("Convene Risk Committee; approve Horizon 1 funding ($1.72M)")
        self._bullet("Deploy counterparty risk dashboard (Month 3)")
        self._bullet("Procure trade credit insurance for top-quartile exposure (Month 6)")
        self._bullet("Implement tiered credit limits policy (Month 4)")
        self._bullet("Deploy v1 credit scoring model + pilot (Month 9)")
        self._bullet("Hire 2 credit analyst FTEs (Month 3)")
        self._bullet("Complete trade credit insurance RFP and binding (Month 6)")

        self._h3("5-Year Priorities (Horizons 1–2)")
        self._bullet("Roll out ML-enhanced credit scoring v2 (Year 2–3)")
        self._bullet("Launch systematic hedging overlay (Year 2)")
        self._bullet("Initiate geographic diversification (Year 3–5)")
        self._bullet("Launch fee-based advisory services (Year 3)")
        self._bullet("Modernize technology platform (Years 2–4)")
        self._bullet("Achieve <40% single-segment concentration (Year 5)")
        self._bullet("Expected cumulative NPV at Year 5: ~$9M")

        self._h3("10-Year Priorities (Horizon 3)")
        self._bullet("Execute strategic M&A / partnership (Years 5–7)")
        self._bullet("Expand advisory product line to 15–20% of revenue (Year 10)")
        self._bullet("Restructure cooperative capital for resilience (Years 6–10)")
        self._bullet("Establish PLMA as reference operator in ag risk management (Year 10)")
        self._bullet("Expected cumulative NPV at Year 10: ~$20M")

    def _add_section_7_financial_model(self):
        self._h1("7. Financial Model: Cost-Benefit Analysis")

        self._h2("7.1 Assumptions & Methodology")

        self._body(
            "The financial model uses a probability-weighted Expected Loss framework (Basel-"
            "inspired): EL = PD × EAD × LGD, where PD is probability of default, EAD is exposure "
            "at default, and LGD is loss given default. Benefit estimates reflect loss avoidance "
            "from mitigation actions. NPV uses an 8% discount rate, consistent with cooperative "
            "cost-of-capital benchmarks.")
        p = self.doc.paragraphs[-1]
        self._cite(p, 15)

        self._h2("7.2 Expected Loss Baseline")

        el_rows = []
        for seg in self.model.expected_loss_by_segment():
            el_rows.append([
                seg["name"],
                f"${seg['ead']/1e6:.0f}M",
                f"{seg['pd']*100:.1f}%",
                f"{seg['lgd']*100:.0f}%",
                f"${seg['el']/1e6:.2f}M",
            ])
        el_rows.append(["**TOTAL**", f"${200:.0f}M", "—", "—", f"${self.model.total_expected_loss()/1e6:.2f}M"])
        self._table(
            ["Segment", "EAD", "PD (1yr)", "LGD", "Expected Loss"],
            el_rows,
            col_widths=[2.2, 0.9, 0.9, 0.7, 1.2]
        )

        self._callout(
            "Stress Scenario",
            f"Under a 2x PD multiplier (consistent with a farm-crisis-level default environment), "
            f"annual Expected Loss rises from ${self.model.total_expected_loss()/1e6:.2f}M to "
            f"${self.model.stress_expected_loss()/1e6:.2f}M—approximately 24% of annual commission revenue.",
            kind="warning"
        )

        self._h2("7.3 Mitigation Program Costs")

        y1_rows = [[k, f"${v/1000:.0f}K"] for k, v in self.model.year1_costs.items()]
        y1_rows.append(["**TOTAL Year 1**", f"${self.model.total_year1_cost()/1000:.0f}K"])
        self._table(
            ["Year 1 Cost Item", "Amount"],
            y1_rows,
            col_widths=[4.0, 1.5]
        )

        ongoing_rows = [[k, f"${v/1000:.0f}K"] for k, v in self.model.ongoing_costs.items()]
        ongoing_rows.append(["**TOTAL Annual Ongoing**", f"${self.model.total_ongoing_cost()/1000:.0f}K"])
        self._table(
            ["Ongoing Annual Cost (Year 2+)", "Amount"],
            ongoing_rows,
            col_widths=[4.0, 1.5]
        )

        self._h2("7.4 Projected Loss Avoidance")

        benefit_rows = [[k, f"${v/1000:.0f}K"] for k, v in self.model.annual_benefits.items()]
        benefit_rows.append(["**TOTAL Annual Benefit**", f"${self.model.total_annual_benefit()/1000:.0f}K"])
        self._table(
            ["Benefit Category", "Annual Value"],
            benefit_rows,
            col_widths=[4.0, 1.5]
        )

        self._chart(self.charts.cost_benefit_waterfall(),
                    caption="Net steady-state annual economics: $3.38M positive, before discounting.")

        self._h2("7.5 10-Year NPV & ROI Projection")

        npv_val = self.model.npv()
        self._body(
            f"Using an 8% discount rate over the 10-year horizon, the mitigation program delivers "
            f"an NPV of approximately ${npv_val/1e6:.1f}M. Cumulative 10-year ROI is "
            f"~{self.model.roi_10yr()*100:.0f}%.")

        self._chart(self.charts.roi_projection(),
                    caption="Breakeven occurs within Year 1; cumulative benefit reaches ~$29M nominally by Year 10.")

        cf_rows = []
        for f in self.model.cash_flows():
            cf_rows.append([
                f"Year {f['year']}",
                f"${f['cost']/1000:.0f}K",
                f"${f['benefit']/1000:.0f}K",
                f"${f['net']/1000:.0f}K",
            ])
        self._table(
            ["Year", "Cost", "Benefit", "Net Cash Flow"],
            cf_rows,
            col_widths=[1.0, 1.4, 1.4, 1.6]
        )

        self._h2("7.6 Sensitivity Analysis")

        self._body(
            "Tornado analysis identifies which assumptions most materially drive NPV outcomes. "
            "The top three sensitivities are: (1) default rate assumptions, (2) benefit "
            "realization rate, and (3) discount rate.")

        self._chart(self.charts.sensitivity_tornado(),
                    caption="Default rate and benefit realization drive the widest NPV swings; program is robust across scenarios.")

        self._callout(
            "Robustness",
            "Even in the most adverse sensitivity scenario tested, the program maintains positive "
            "NPV. This suggests the investment case is robust to reasonable assumption variation.",
            kind="finding"
        )

    def _add_section_8_governance(self):
        self._h1("8. Governance & Monitoring")

        self._body(
            "Sustainable risk management requires governance structures that institutionalize "
            "oversight, reporting, and course-correction. We recommend a three-tier governance "
            "model: Board Risk Committee, Management Risk Committee, and Operational Risk Team.")

        self._h2("8.1 Risk Committee Structure")

        self._table(
            ["Body", "Composition", "Cadence", "Mandate"],
            [
                ["Board Risk Committee",
                 "3–4 Directors + CEO + CFO",
                 "Quarterly",
                 "Approve risk appetite; review large exposures; oversee mitigation program"],
                ["Management Risk Committee",
                 "CFO (Chair), COO, Credit Lead, Hedging Lead",
                 "Monthly",
                 "Monitor KPIs; approve exposures > single-counterparty threshold; escalate breaches"],
                ["Operational Risk Team",
                 "Credit analysts + Operations",
                 "Weekly",
                 "Score counterparties; monitor early warning indicators; execute daily limits"],
            ],
            col_widths=[1.5, 1.8, 0.9, 2.3]
        )

        self._h2("8.2 KPI Dashboard")

        self._table(
            ["KPI", "Target", "Red Flag"],
            [
                ["Top-10 concentration (% exposure)", "<50%", ">60%"],
                ["Largest single counterparty", "<$15M or 10% equity", ">$20M"],
                ["Avg counterparty credit score", ">6.5 / 10", "<5.5 / 10"],
                ["Realized credit losses (% revenue)", "<1.5%", ">2.5%"],
                ["Days Sales Outstanding", "<8 days", ">12 days"],
                ["Hedge ratio (inventory exposure)", "60–80%", "<40%"],
                ["Trade credit insurance coverage", ">40% exposure", "<25%"],
                ["Early warning alerts/month", "Monitor trend", ">20% MoM increase"],
            ],
            col_widths=[2.5, 1.5, 1.5]
        )

        self._h2("8.3 Escalation Protocols")

        self._bullet_lead("Tier 1 (Operational):",
            "Single-counterparty exposure exceeds tiered limit → immediate credit officer review, 24-hour resolution.")
        self._bullet_lead("Tier 2 (Management):",
            "Concentration or segment threshold breached → Management Risk Committee review within 5 business days.")
        self._bullet_lead("Tier 3 (Board):",
            "Realized losses exceed quarterly budget, or systemic red-flag event → Board Risk Committee convened within 10 business days.")

        self._h2("8.4 Annual Review & Recalibration")

        self._body(
            "Risk parameters (PDs, LGDs, concentration limits, hedging targets) should be "
            "recalibrated annually against realized experience and updated macroeconomic outlook. "
            "Model validation by an independent party (internal audit or external advisor) should "
            "occur on a two-year cycle.")

    def _add_section_9_conclusion(self):
        self._h1("9. Conclusion & Next Steps")

        self._body(
            "PLMA stands at a strategic inflection point. The convergence of record farm debt, "
            "accelerating bankruptcies, and structural margin compression across the livestock "
            "value chain creates the most acute counterparty credit risk environment the "
            "cooperative has faced in a generation. The risk is quantifiable, concentrated, and—"
            "critically—manageable through a disciplined, phased mitigation program.")

        self._callout(
            "The Opportunity",
            "A $1.7M Year-1 investment generates ~$20M in 10-year NPV, ~250% cumulative ROI, "
            "and payback within 12 months. Beyond the financial return, the program positions "
            "PLMA as a reference operator in ag risk management—a strategic differentiator in a "
            "consolidating industry.",
            kind="insight"
        )

        self._h2("9.1 Immediate Actions (First 90 Days)")

        self._bullet("Week 1–2: CFO-led executive alignment on report findings; Risk Committee briefing scheduled")
        self._bullet("Week 2–4: Secure Board approval for Horizon 1 funding ($1.72M)")
        self._bullet("Week 4–8: Issue RFPs for trade credit insurance (3+ carriers) and credit analytics platform")
        self._bullet("Week 6–10: Hire 2 credit analyst FTEs")
        self._bullet("Week 8–12: Deploy counterparty risk dashboard (v1, internal data)")
        self._bullet("Week 12: Finalize tiered credit limits policy and Risk Appetite Statement")

        self._h2("9.2 Phase 2 Engagement Model")

        self._body(
            "For Horizon 2 and Horizon 3 workstreams, we recommend a staged advisory engagement: "
            "(i) quarterly progress reviews against roadmap milestones, (ii) annual risk "
            "parameter recalibration, (iii) on-call support for emergent counterparty events, and "
            "(iv) strategic advisory on M&A/diversification options as they arise.")

        self._body(
            "We stand ready to support PLMA in transforming this risk exposure into a strategic "
            "advantage.")

    def _add_appendices(self):
        self._h1("Appendix A: Detailed Data Tables")

        self._h2("A.1 Historical Farm Bankruptcy Filings")
        self._table(
            ["Year", "Chapter 12 Filings", "YoY Change"],
            [
                ["2018", "498", "—"],
                ["2019", "595", "+19.5%"],
                ["2020", "552", "-7.2%"],
                ["2021", "371", "-32.8%"],
                ["2022", "275", "-25.9%"],
                ["2023", "287", "+4.4%"],
                ["2024", "258", "-10.1%"],
                ["2025", "315", "+22.1%"],
                ["2026E", "380", "+20.6%"],
            ],
            col_widths=[1.2, 1.8, 1.5]
        )

        self._h2("A.2 U.S. Farm Debt Detail ($ Billions)")
        self._table(
            ["Year", "Real Estate", "Non-Real-Estate", "Total"],
            [
                ["2015", "$213", "$172", "$385"],
                ["2018", "$253", "$198", "$451"],
                ["2020", "$279", "$215", "$494"],
                ["2022", "$325", "$205", "$530"],
                ["2024", "$382", "$215", "$597"],
                ["2025", "$405", "$220", "$625"],
                ["2026E", "$430", "$230", "$660"],
            ],
            col_widths=[1.0, 1.3, 1.5, 1.0]
        )

        self._h2("A.3 Cash Flow Detail")
        cf_rows = []
        cum = 0
        for f in self.model.cash_flows():
            cum += f["net"]
            cf_rows.append([
                f"Year {f['year']}",
                f"${f['cost']/1000:.0f}K",
                f"${f['benefit']/1000:.0f}K",
                f"${f['net']/1000:.0f}K",
                f"${cum/1000:.0f}K",
            ])
        self._table(
            ["Year", "Cost", "Benefit", "Net", "Cumulative"],
            cf_rows,
            col_widths=[0.9, 1.1, 1.1, 1.1, 1.3]
        )

        self._h1("Appendix B: Methodology Notes")

        self._h2("B.1 Expected Loss Calculation")
        self._body(
            "Expected Loss (EL) is calculated as EL = PD × EAD × LGD following Basel Committee "
            "credit risk conventions. PDs are calibrated to segment-specific default rate "
            "benchmarks from S&P Global Market Intelligence and Rabobank ag outlooks. EADs "
            "reflect estimated average exposure based on PLMA's transaction volume and typical "
            "settlement cycles. LGDs reflect recovery rate estimates from livestock-sector "
            "bankruptcy proceedings.")

        self._h2("B.2 NPV & Discount Rate")
        self._body(
            "NPV discounts net cash flows at 8%, reflecting typical cooperative cost-of-capital "
            "per CoBank and Farm Credit System benchmarks. Sensitivity analysis tests 6–12% range.")

        self._h2("B.3 Benefit Estimation")
        self._body(
            "Benefit estimates are calibrated from peer-cooperative case studies and published "
            "efficacy data for credit analytics platforms, trade credit insurance programs, and "
            "commodity hedging overlays. Estimates reflect mid-range efficacy and may be "
            "conservative.")

        self._h2("B.4 Scenario Definitions")
        self._bullet_lead("Soft Landing:", "Bankruptcy rate declines 10% by 2027; cattle cycle correction modest (<15%)")
        self._bullet_lead("Base Case:", "Bankruptcy rate grows 15% annually through 2027; cattle prices flat")
        self._bullet_lead("Prolonged Stress:", "Bankruptcy rate grows 30% annually; cattle prices decline 20–25%")
        self._bullet_lead("Systemic Crisis:", "Bankruptcy rate doubles; cattle decline >35%; regional ASF outbreak")

        self._h1("Appendix C: Glossary of Terms")

        terms = [
            ("ASF", "African Swine Fever — highly contagious viral disease affecting pigs."),
            ("Backgrounder", "Operation that grows weaned calves on forage before feedlot placement."),
            ("Basel EL", "Basel Committee Expected Loss framework: EL = PD × EAD × LGD."),
            ("Basis Risk", "Risk that futures prices diverge from cash (physical) prices."),
            ("Chapter 12", "U.S. Bankruptcy Code chapter specific to family farmers and fishermen."),
            ("CFTC", "Commodity Futures Trading Commission — U.S. futures market regulator."),
            ("EAD", "Exposure At Default — outstanding amount owed at time of default."),
            ("Feedlot", "Concentrated animal feeding operation finishing cattle for slaughter."),
            ("LGD", "Loss Given Default — share of exposure lost after recovery."),
            ("P&SA", "Packers & Stockyards Act — USDA-enforced livestock marketing regulation."),
            ("PD", "Probability of Default — likelihood counterparty defaults in given period."),
        ]
        for term, defn in terms:
            self._bullet_lead(term + ":", defn)

        self._add_citations()

    def _add_citations(self):
        self._h1("Appendix D: Citations")

        self._body(
            "All citations numbered and referenced in the body of this report. Where specific "
            "figures have been estimated or modeled, appendix notes describe methodology.")

        citations = [
            "USDA Economic Research Service, \"Farm Income and Wealth Statistics,\" 2025–2026 release.",
            "American Bankruptcy Institute, \"Chapter 12 Filing Statistics,\" 2025 annual report.",
            "USDA National Agricultural Statistics Service, \"Cattle on Feed,\" monthly reports 2025–2026.",
            "Federal Reserve Bank of Kansas City, \"Agricultural Finance Databook,\" Q4 2025.",
            "USDA Economic Research Service, \"Farm Sector Balance Sheet,\" February 2026.",
            "Federal Reserve Board, \"Senior Loan Officer Opinion Survey,\" Q1 2026.",
            "CME Group, \"Livestock Market Analysis,\" 2025–2026.",
            "Rabobank, \"Global Animal Protein Outlook,\" 2026 edition.",
            "CoBank, \"U.S. Agricultural Cooperatives: Risk and Resilience,\" 2025 industry report.",
            "USDA Agricultural Marketing Service (Packers & Stockyards Division), \"Packers and Stockyards Annual Report,\" 2025.",
            "National Cattlemen's Beef Association, \"Cattle Industry Outlook,\" 2026.",
            "Porter, M.E., \"The Five Competitive Forces That Shape Strategy,\" Harvard Business Review, January 2008.",
            "Kaplan, R.S. & Mikes, A., \"Managing Risks: A New Framework,\" Harvard Business Review, June 2012.",
            "International Organization for Standardization, ISO 31000:2018, \"Risk Management — Guidelines.\"",
            "Basel Committee on Banking Supervision, \"Credit Risk: Standardised Approach and Internal Ratings-Based Approach,\" 2023.",
            "Federal Reserve Economic Data (FRED), \"Effective Federal Funds Rate,\" 2026.",
            "Livestock Marketing Association, \"Industry Economic Impact Study,\" 2024.",
            "USDA Economic Research Service, \"Commodity Costs and Returns,\" 2025.",
            "USDA NASS, \"Agricultural Prices,\" monthly reports 2025–2026.",
            "S&P Global Market Intelligence, \"U.S. Agribusiness Credit Trends,\" 2025.",
            "Deloitte, \"2026 Agricultural Industry Outlook.\"",
            "McKinsey & Company, \"Agriculture Practice: Risk Management in Volatile Markets,\" 2024.",
            "Farm Credit Administration, \"Annual Report,\" 2025.",
            "Commodity Futures Trading Commission, \"Commitments of Traders Reports,\" 2025–2026.",
            "Congressional Research Service, \"Farm Bankruptcies: An Overview,\" Report R46155, 2025.",
            "Sterling Marketing Inc., \"Cattle Feeding Industry Analysis,\" Q1 2026.",
            "USDA ERS, \"Cattle & Beef Outlook,\" March 2026.",
        ]

        for i, cite in enumerate(citations, 1):
            p = self.doc.add_paragraph()
            r1 = p.add_run(f"[{i}]  ")
            r1.bold = True
            r1.font.name = self.config.BODY_FONT
            r1.font.size = Pt(9.5)
            r1.font.color.rgb = self.config.NAVY
            r2 = p.add_run(cite)
            r2.font.name = self.config.BODY_FONT
            r2.font.size = Pt(9.5)
            r2.font.color.rgb = self.config.DARK_GRAY
            p.paragraph_format.left_indent = Inches(0.4)
            p.paragraph_format.first_line_indent = Inches(-0.4)
            set_paragraph_spacing(p, after=5, line=1.2)

    # ============ MAIN BUILD ============
    def build(self):
        self._add_cover_page()
        self._add_toc()
        self._add_executive_summary()
        self._add_section_2_industry()
        self._add_section_3_credit_risk()
        self._add_section_4_risk_assessment()
        self._add_section_5_frameworks()
        self._add_section_6_mitigation()
        self._add_section_7_financial_model()
        self._add_section_8_governance()
        self._add_section_9_conclusion()
        self._add_appendices()
        return self.doc


# =============================================================================
# MAIN
# =============================================================================
def main():
    config = ReportConfig()
    model = FinancialModel()
    charts = ChartGenerator(config, model)
    builder = DocumentBuilder(config, charts, model)
    builder.build()
    output_path = "PLMA_Credit_Risk_Assessment_2026.docx"
    builder.doc.save(output_path)
    print(f"Report generated: {output_path}")
    print(f"10-Year NPV: ${model.npv()/1e6:.2f}M")
    print(f"10-Year ROI: {model.roi_10yr()*100:.1f}%")
    print(f"Total Expected Loss (base): ${model.total_expected_loss()/1e6:.2f}M")
    print(f"Total Expected Loss (stress): ${model.stress_expected_loss()/1e6:.2f}M")


if __name__ == "__main__":
    main()
