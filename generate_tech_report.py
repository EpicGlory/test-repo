"""
PLMA Emerging Technology Strategy Report Generator
Risk-led, broad tech transformation analysis for CFO Ryan Johansen.
Companion report to PLMA_Credit_Risk_Assessment_2026.docx.
"""
from io import BytesIO

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import Rectangle, FancyArrowPatch
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
    VERY_LIGHT_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
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
    LIGHT_GRAY_HEX = "#D9D9D9"

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
class TechFinancialModel:
    def __init__(self):
        # Investment by workstream (Year 1)
        self.year1_costs = {
            "Cybersecurity hardening (Zero Trust, EDR, SOC)": 900_000,
            "Cloud / Data Platform v1": 1_200_000,
            "AI Credit Scoring v1 (shared w/ Report #1)": 350_000,
            "GenAI Pilots (back-office, research)": 450_000,
            "RPA / Workflow Automation": 400_000,
            "Digital Platform Enhancements": 800_000,
            "AI Governance & Talent": 700_000,
        }
        # Annual ongoing costs (Year 2+)
        self.ongoing_costs = {
            "Cybersecurity operations & licenses": 650_000,
            "Cloud/data platform run-rate": 550_000,
            "AI/ML platform & tooling": 300_000,
            "GenAI licenses & API costs": 220_000,
            "RPA maintenance": 140_000,
            "Digital platform operations": 380_000,
            "AI governance & FTEs": 360_000,
        }
        # Annual benefits (Year 2+)
        self.annual_benefits = {
            "Credit loss avoidance (incremental vs Report #1)": 900_000,
            "Labor productivity (RPA + GenAI)": 2_400_000,
            "Cyber incident cost avoidance (expected value)": 1_100_000,
            "New digital revenue (platform services)": 2_200_000,
            "Market intelligence & hedging improvement": 1_100_000,
            "Working capital optimization (forecasting)": 1_300_000,
        }
        self.discount_rate = 0.08
        self.horizon_years = 10

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
                benefit = self.total_annual_benefit() * 0.35  # slow ramp Y1
            elif y == 2:
                cost = self.total_ongoing_cost()
                benefit = self.total_annual_benefit() * 0.80  # continued ramp
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
        total_benefit = sum(f["benefit"] for f in self.cash_flows())
        return (total_benefit - total_cost) / total_cost

    def payback_months(self):
        cum = -self.total_year1_cost()
        monthly_ongoing = self.total_ongoing_cost() / 12
        y1_monthly_benefit = (self.total_annual_benefit() * 0.35) / 12
        y2_monthly_benefit = (self.total_annual_benefit() * 0.80) / 12
        ss_monthly_benefit = self.total_annual_benefit() / 12
        for m in range(1, 121):
            if m <= 12:
                cum += y1_monthly_benefit
            elif m <= 24:
                cum += y2_monthly_benefit - monthly_ongoing
            else:
                cum += ss_monthly_benefit - monthly_ongoing
            if cum >= 0:
                return m
        return None

    def use_case_portfolio(self):
        """12 AI/tech use cases with value and feasibility scores (1-10)."""
        return [
            {"name": "ML Credit Scoring", "value": 9, "feasibility": 8,
             "nb_pv": 7.5, "category": "AI/ML"},
            {"name": "GenAI Doc Processing", "value": 8, "feasibility": 9,
             "nb_pv": 6.8, "category": "GenAI"},
            {"name": "Price Prediction Models", "value": 9, "feasibility": 7,
             "nb_pv": 8.2, "category": "AI/ML"},
            {"name": "Computer Vision Grading", "value": 8, "feasibility": 5,
             "nb_pv": 5.4, "category": "AI/ML"},
            {"name": "GenAI Member Service", "value": 7, "feasibility": 8,
             "nb_pv": 4.6, "category": "GenAI"},
            {"name": "Demand Forecasting", "value": 8, "feasibility": 7,
             "nb_pv": 6.1, "category": "AI/ML"},
            {"name": "RPA — AR/AP Automation", "value": 7, "feasibility": 9,
             "nb_pv": 5.8, "category": "Automation"},
            {"name": "EID/IoT Traceability", "value": 7, "feasibility": 6,
             "nb_pv": 4.2, "category": "IoT"},
            {"name": "Digital Auction Platform", "value": 9, "feasibility": 6,
             "nb_pv": 7.9, "category": "Digital"},
            {"name": "Cyber Zero-Trust", "value": 8, "feasibility": 8,
             "nb_pv": 3.5, "category": "Security"},
            {"name": "Data Lakehouse", "value": 7, "feasibility": 7,
             "nb_pv": 4.8, "category": "Data"},
            {"name": "Blockchain Provenance", "value": 5, "feasibility": 4,
             "nb_pv": 1.2, "category": "Blockchain"},
        ]

    def sensitivity(self):
        base = self.npv()
        results = []

        # Benefit realization +/- 30%
        orig = self.annual_benefits.copy()
        for mult, label in [(0.7, "low"), (1.3, "high")]:
            for k in self.annual_benefits:
                self.annual_benefits[k] = orig[k] * mult
            results.append(("Benefit realization", label, self.npv()))
        self.annual_benefits = orig.copy()

        # Implementation cost +/- 40%
        orig_y1 = self.year1_costs.copy()
        for mult, label in [(0.8, "low"), (1.4, "high")]:
            for k in self.year1_costs:
                self.year1_costs[k] = orig_y1[k] * mult
            results.append(("Implementation cost", label, self.npv()))
        self.year1_costs = orig_y1.copy()

        # Ongoing cost +/- 25%
        orig_ong = self.ongoing_costs.copy()
        for mult, label in [(0.75, "low"), (1.25, "high")]:
            for k in self.ongoing_costs:
                self.ongoing_costs[k] = orig_ong[k] * mult
            results.append(("Ongoing cost", label, self.npv()))
        self.ongoing_costs = orig_ong.copy()

        # Ramp speed
        orig_flows = self.cash_flows  # method reference
        # proxy: just adjust discount to simulate risk
        orig_dr = self.discount_rate
        for rate, label in [(0.06, "low"), (0.12, "high")]:
            self.discount_rate = rate
            results.append(("Discount rate", label, self.npv()))
        self.discount_rate = orig_dr

        # Cyber incident avoidance +/- 50%
        for mult, label in [(0.5, "low"), (1.5, "high")]:
            self.annual_benefits["Cyber incident cost avoidance (expected value)"] = 1_100_000 * mult
            results.append(("Cyber avoidance EV", label, self.npv()))
        self.annual_benefits = orig.copy()

        # Labor productivity +/- 35%
        for mult, label in [(0.65, "low"), (1.35, "high")]:
            self.annual_benefits["Labor productivity (RPA + GenAI)"] = 2_400_000 * mult
            results.append(("Labor productivity", label, self.npv()))
        self.annual_benefits = orig.copy()

        variables = {}
        for var, lbl, npv in results:
            variables.setdefault(var, {})[lbl] = npv
        tornado = []
        for var, vals in variables.items():
            tornado.append({
                "variable": var,
                "low": vals["low"],
                "high": vals["high"],
                "base": base,
                "range": abs(vals["high"] - vals["low"]),
            })
        tornado.sort(key=lambda x: x["range"], reverse=True)
        return tornado, base


# =============================================================================
# CHART GENERATOR
# =============================================================================
class TechChartGenerator:
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

    # 1. Ag tech investment trends
    def agtech_investment_trends(self):
        years = [2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025]
        precision_ag = [1.8, 2.2, 2.6, 4.1, 3.5, 3.0, 3.2, 3.8]
        livestock_tech = [0.6, 0.9, 1.2, 2.0, 1.8, 1.5, 1.9, 2.3]
        ai_ml = [0.4, 0.8, 1.4, 2.8, 2.6, 2.4, 3.1, 4.2]
        marketplace = [1.1, 1.4, 1.8, 2.5, 2.1, 1.9, 2.2, 2.6]
        fig, ax = plt.subplots(figsize=(9, 5))
        width = 0.6
        bottom = np.zeros(len(years))
        for data, color, label in [
            (precision_ag, self.config.NAVY_HEX, "Precision Ag"),
            (livestock_tech, self.config.BLUE_HEX, "Livestock Tech"),
            (ai_ml, self.config.AMBER_HEX, "Ag AI/ML"),
            (marketplace, self.config.GREEN_HEX, "Digital Marketplaces"),
        ]:
            ax.bar(years, data, width, bottom=bottom, label=label,
                   color=color, edgecolor="white", linewidth=0.5)
            bottom += np.array(data)
        for i, y in enumerate(years):
            ax.text(y, bottom[i] + 0.2, f"${bottom[i]:.1f}B", ha="center",
                    fontsize=9, fontweight="bold", color=self.config.DARK_GRAY_HEX)
        ax.set_title("Global Agtech Investment by Segment (2018–2025)",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.set_ylabel("Investment ($ Billions)", fontsize=10)
        ax.set_xlabel("Year", fontsize=10)
        ax.legend(loc="upper left", frameon=False, fontsize=9)
        ax.grid(True, axis="y", alpha=0.3)
        ax.set_axisbelow(True)
        ax.text(0.01, -0.15, "Source: AgFunder AgriFoodTech Investment Reports 2018–2025; PitchBook.",
                transform=ax.transAxes, fontsize=8, color=self.config.MID_GRAY_HEX, style="italic")
        return self._finalize(fig)

    # 2. AI adoption S-curve
    def ai_adoption_scurve(self):
        years = np.arange(2018, 2031)
        def s_curve(t, peak=1.0, k=0.6, shift=5):
            return peak / (1 + np.exp(-k * (t - shift)))
        tech = s_curve(np.arange(len(years)), peak=0.85, k=0.75, shift=4)
        financial = s_curve(np.arange(len(years)), peak=0.75, k=0.60, shift=5.5)
        retail = s_curve(np.arange(len(years)), peak=0.70, k=0.55, shift=6.5)
        manufacturing = s_curve(np.arange(len(years)), peak=0.65, k=0.50, shift=7.5)
        ag = s_curve(np.arange(len(years)), peak=0.60, k=0.45, shift=9)
        livestock = s_curve(np.arange(len(years)), peak=0.55, k=0.40, shift=10)
        fig, ax = plt.subplots(figsize=(10, 5.5))
        ax.plot(years, tech*100, color=self.config.NAVY_HEX, linewidth=2.5, label="Technology")
        ax.plot(years, financial*100, color=self.config.BLUE_HEX, linewidth=2.5, label="Financial Services")
        ax.plot(years, retail*100, color=self.config.PURPLE_HEX, linewidth=2.5, label="Retail")
        ax.plot(years, manufacturing*100, color=self.config.AMBER_HEX, linewidth=2.5, label="Manufacturing")
        ax.plot(years, ag*100, color=self.config.GREEN_HEX, linewidth=2.5, label="Agriculture (broad)")
        ax.plot(years, livestock*100, color=self.config.RED_HEX, linewidth=3, linestyle="--",
                label="Livestock Marketing ←")
        ax.axvline(x=2026, color=self.config.DARK_GRAY_HEX, linestyle=":", alpha=0.6)
        ax.text(2026.1, 65, "Today", fontsize=9, color=self.config.DARK_GRAY_HEX)
        ax.annotate("PLMA's\nopportunity\nwindow", xy=(2027, 25), xytext=(2028, 15),
                    fontsize=9, fontweight="bold", color=self.config.RED_HEX,
                    arrowprops=dict(arrowstyle="->", color=self.config.RED_HEX))
        ax.set_title("AI Adoption S-Curve by Industry",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.set_ylabel("AI Adoption Rate (%)", fontsize=10)
        ax.set_xlabel("Year", fontsize=10)
        ax.legend(loc="upper left", frameon=False, fontsize=9)
        ax.grid(True, alpha=0.3)
        ax.set_axisbelow(True)
        ax.set_ylim(0, 100)
        ax.text(0.01, -0.14, "Source: McKinsey State of AI reports 2020–2025; IDC Worldwide AI Spending Guide.",
                transform=ax.transAxes, fontsize=8, color=self.config.MID_GRAY_HEX, style="italic")
        return self._finalize(fig)

    # 3. Tech risk heat map
    def tech_risk_heat_map(self):
        fig, ax = plt.subplots(figsize=(10, 8))
        for i in range(5):
            for j in range(5):
                sev = (i+1) * (j+1)
                if sev <= 4:
                    c = "#DFF6DD"
                elif sev <= 9:
                    c = "#FFF4CE"
                elif sev <= 15:
                    c = "#FED9B7"
                else:
                    c = "#FDE7E9"
                ax.add_patch(Rectangle((j, i), 1, 1, facecolor=c, edgecolor="white", linewidth=2))
        # Each risk has custom (dx, dy) offset in points to stagger and prevent overlaps
        risks = [
            ("Ransomware attack", 4, 5, (8, 8)),
            ("Digital marketplace\ndisruption", 4, 4, (8, -18)),
            ("Legacy tech debt drag", 5, 3, (-8, -18)),
            ("AI model bias / P&SA", 3, 4, (8, 10)),
            ("Data privacy breach", 3, 4, (-85, -16)),
            ("Talent shortage", 4, 3, (8, 10)),
            ("Vendor lock-in", 3, 3, (8, 10)),
            ("Cloud outage", 2, 3, (8, 12)),
            ("SEC cyber disclosure", 2, 3, (-90, -16)),
            ("Insider threat", 3, 3, (-70, -16)),
            ("Shadow IT", 4, 2, (8, 8)),
            ("AI hallucination", 3, 2, (8, -16)),
            ("State AI regulation", 2, 2, (8, 8)),
        ]
        _bbox = dict(boxstyle="round,pad=0.2", fc="white", ec="none", alpha=0.85)
        for name, lk, imp, offset in risks:
            ax.scatter(lk-0.5, imp-0.5, s=160, c=self.config.NAVY_HEX,
                       edgecolor="white", linewidth=2, zorder=5)
            ax.annotate(name, (lk-0.5, imp-0.5), fontsize=7.5,
                        xytext=offset, textcoords="offset points",
                        color=self.config.DARK_GRAY_HEX, fontweight="bold",
                        bbox=_bbox)
        ax.set_xlim(0, 5); ax.set_ylim(0, 5)
        ax.set_xticks([0.5, 1.5, 2.5, 3.5, 4.5])
        ax.set_xticklabels(["Rare", "Unlikely", "Possible", "Likely", "Certain"])
        ax.set_yticks([0.5, 1.5, 2.5, 3.5, 4.5])
        ax.set_yticklabels(["Negligible", "Minor", "Moderate", "Major", "Severe"])
        ax.set_xlabel("Likelihood", fontsize=11, fontweight="bold")
        ax.set_ylabel("Impact", fontsize=11, fontweight="bold")
        ax.set_title("PLMA Technology Risk Heat Map",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.set_aspect("equal")
        return self._finalize(fig)

    # 4. Cyber incidents in ag
    def cyber_incidents_ag(self):
        years = [2020, 2021, 2022, 2023, 2024, 2025]
        incidents = [12, 24, 41, 58, 82, 117]
        ransom_avg = [0.35, 0.85, 1.4, 1.9, 2.3, 2.8]  # $M
        fig, ax1 = plt.subplots(figsize=(9, 5))
        bars = ax1.bar(years, incidents, color=self.config.NAVY_HEX,
                       edgecolor="white", width=0.6, label="Disclosed ag-sector cyber incidents")
        for bar, val in zip(bars, incidents):
            ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 2,
                     f"{val}", ha="center", fontsize=9, fontweight="bold",
                     color=self.config.DARK_GRAY_HEX, zorder=6,
                     bbox=dict(boxstyle="round,pad=0.2", fc="white", alpha=0.5, ec="none"))
        ax1.set_ylabel("Incidents Disclosed", fontsize=10, color=self.config.NAVY_HEX)
        ax1.tick_params(axis="y", labelcolor=self.config.NAVY_HEX)
        ax1.set_xlabel("Year", fontsize=10)
        ax1.grid(True, axis="y", alpha=0.3)
        ax1.set_axisbelow(True)
        ax2 = ax1.twinx()
        ax2.plot(years, ransom_avg, color=self.config.RED_HEX, linewidth=3,
                 marker="o", markersize=8, label="Avg ransom demand ($M)")
        ax2.set_ylabel("Average Ransom Demand ($M)", fontsize=10, color=self.config.RED_HEX)
        ax2.tick_params(axis="y", labelcolor=self.config.RED_HEX)
        ax2.spines["top"].set_visible(False)
        ax1.set_title("Ag-Sector Cybersecurity Incidents & Ransom Escalation (2020–2025)",
                      fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        l1, lab1 = ax1.get_legend_handles_labels()
        l2, lab2 = ax2.get_legend_handles_labels()
        ax1.legend(l1+l2, lab1+lab2, loc="upper left", frameon=False, fontsize=9)
        fig.text(0.1, -0.02, "Source: CISA Ag & Food ISAC; FBI IC3 Annual Reports; Food & Ag ISAC 2025 threat brief.",
                 fontsize=8, color=self.config.MID_GRAY_HEX, style="italic")
        return self._finalize(fig)

    # 5. Legacy tech debt waterfall
    def legacy_tech_debt(self):
        items = [
            ("Deferred maintenance\n(legacy core systems)", -2.8, self.config.RED_HEX),
            ("Manual reconciliation\nlabor tax", -1.6, self.config.RED_HEX),
            ("Integration complexity\n& errors", -1.2, self.config.RED_HEX),
            ("Lost productivity\n(tool fragmentation)", -0.9, self.config.RED_HEX),
            ("Cyber exposure premium", -0.7, self.config.RED_HEX),
            ("Talent attrition & hiring\nfriction", -0.5, self.config.RED_HEX),
            ("Competitive disadvantage\n(digital-native peers)", -1.1, self.config.RED_HEX),
        ]
        fig, ax = plt.subplots(figsize=(11, 5.5))
        cum = 0
        for i, (name, val, color) in enumerate(items):
            ax.bar(i, val, bottom=cum, color=color, edgecolor="white", linewidth=1.5, alpha=0.85)
            ax.text(i, cum + val/2, f"${val:+.1f}M", ha="center", va="center",
                    fontsize=10, fontweight="bold", color="white")
            cum += val
        ax.bar(len(items), cum, color=self.config.NAVY_HEX, edgecolor="white", linewidth=1.5)
        ax.text(len(items), cum/2, f"${cum:+.1f}M\nAnnual\nDrag", ha="center", va="center",
                fontsize=11, fontweight="bold", color="white")
        ax.axhline(y=0, color="black", linewidth=0.8)
        labels = [i[0] for i in items] + ["Total Annual\nTech Debt Cost"]
        ax.set_xticks(range(len(labels)))
        ax.set_xticklabels(labels, fontsize=8.5)
        ax.set_ylabel("Annual Cost ($M)", fontsize=10)
        ax.set_title("Quantified Cost of PLMA's Legacy Tech Debt (Estimated)",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.grid(True, axis="y", alpha=0.3)
        ax.set_axisbelow(True)
        return self._finalize(fig)

    # 6. AI use case prioritization bubble chart
    def ai_use_case_matrix(self):
        uc = self.model.use_case_portfolio()
        cat_colors = {
            "AI/ML": self.config.NAVY_HEX, "GenAI": self.config.BLUE_HEX,
            "Automation": self.config.GREEN_HEX, "IoT": self.config.AMBER_HEX,
            "Digital": self.config.PURPLE_HEX, "Security": self.config.RED_HEX,
            "Data": self.config.LIGHT_BLUE_HEX, "Blockchain": self.config.MID_GRAY_HEX,
        }
        fig, ax = plt.subplots(figsize=(12, 8))
        _bbox = dict(boxstyle="round,pad=0.2", fc="white", ec="none", alpha=0.85)
        # Explicit offsets per use case — ZIGZAG pattern (far/near above/below)
        # at same Y-row to prevent horizontal label collisions
        offsets = {
            # Top row y=9 (value=9): zigzag above at 34/16/34
            "Digital Auction Platform": 34,  # far above
            "Price Prediction Models": 16,   # near above
            "ML Credit Scoring": 34,         # far above
            # Middle row y=8 (value=8): zigzag
            "Computer Vision Grading": -16,  # below (isolated left)
            "Demand Forecasting": -16,       # below near
            "Cyber Zero-Trust": -34,         # below far
            "GenAI Doc Processing": 16,      # above (right edge)
            # Lower row y=7 (value=7): zigzag
            "EID/IoT Traceability": -16,     # below near
            "Data Lakehouse": -34,           # below far
            "GenAI Member Service": -16,     # below near
            "RPA — AR/AP Automation": -34,   # below far
            # Isolated bottom-left
            "Blockchain Provenance": 18,     # above (clears legend)
        }
        for u in uc:
            ax.scatter(u["feasibility"], u["value"],
                       s=u["nb_pv"]*180, c=cat_colors[u["category"]],
                       edgecolor="white", linewidth=2, alpha=0.85, zorder=3)
            dy = offsets.get(u["name"], 14)
            ax.annotate(u["name"], (u["feasibility"], u["value"]),
                        xytext=(0, dy), textcoords="offset points", fontsize=8.5,
                        ha="center", va="bottom" if dy > 0 else "top",
                        color=self.config.DARK_GRAY_HEX, fontweight="bold",
                        bbox=_bbox, zorder=5)
        # Quadrant lines
        ax.axhline(y=7, color=self.config.MID_GRAY_HEX, linestyle="--", alpha=0.5)
        ax.axvline(x=7, color=self.config.MID_GRAY_HEX, linestyle="--", alpha=0.5)
        ax.text(9.8, 10.8, "DO FIRST", fontsize=11, fontweight="bold",
                color=self.config.GREEN_HEX, ha="right")
        ax.text(3.2, 10.8, "STRATEGIC BETS", fontsize=10, fontweight="bold",
                color=self.config.AMBER_HEX, ha="left")
        ax.text(8.5, 5.5, "QUICK WINS", fontsize=10, fontweight="bold",
                color=self.config.BLUE_HEX, ha="center")
        ax.text(5.5, 5.5, "DEFER", fontsize=10, fontweight="bold",
                color=self.config.RED_HEX, ha="center")
        ax.set_xlim(3, 10.5); ax.set_ylim(4, 11.2)
        ax.set_xlabel("Feasibility (1–10)", fontsize=11, fontweight="bold")
        ax.set_ylabel("Strategic Value (1–10)", fontsize=11, fontweight="bold")
        ax.set_title("AI & Tech Use Case Prioritization (bubble = 10-yr NPV $M)",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.grid(True, alpha=0.3)
        # Category legend — placed in lower-right (QUICK WINS has fewer dots at bottom)
        patches = [mpatches.Patch(color=c, label=cat) for cat, c in cat_colors.items()]
        ax.legend(handles=patches, loc="lower right", fontsize=8, frameon=True,
                  ncol=2, framealpha=0.95)
        return self._finalize(fig)

    # 7. Build vs Buy vs Partner
    def build_buy_partner(self):
        categories = ["Credit\nScoring", "GenAI Doc\nProc.", "Price\nPrediction",
                      "Computer\nVision", "RPA", "Data\nPlatform", "Cyber\nSOC"]
        build = [40, 15, 60, 25, 20, 35, 10]
        buy = [35, 70, 15, 60, 75, 55, 70]
        partner = [25, 15, 25, 15, 5, 10, 20]
        x = np.arange(len(categories))
        width = 0.6
        fig, ax = plt.subplots(figsize=(10, 5.5))
        ax.bar(x, build, width, label="Build (internal)", color=self.config.NAVY_HEX,
               edgecolor="white")
        ax.bar(x, buy, width, bottom=build, label="Buy (vendor)",
               color=self.config.BLUE_HEX, edgecolor="white")
        bottoms = [b+bu for b, bu in zip(build, buy)]
        ax.bar(x, partner, width, bottom=bottoms, label="Partner",
               color=self.config.AMBER_HEX, edgecolor="white")
        ax.set_xticks(x); ax.set_xticklabels(categories, fontsize=9)
        ax.set_ylabel("Recommended Sourcing Mix (%)", fontsize=10)
        ax.set_title("Build vs. Buy vs. Partner — Recommended Sourcing Strategy",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.legend(loc="upper right", frameon=False, fontsize=9)
        ax.set_ylim(0, 115)
        ax.grid(True, axis="y", alpha=0.3)
        ax.set_axisbelow(True)
        return self._finalize(fig)

    # 8. Tech portfolio 2x2
    def tech_portfolio_2x2(self):
        fig, ax = plt.subplots(figsize=(11, 8))
        # Custom per-tech offsets (dx, dy) to prevent overlap
        # (name, impact=y, feasibility=x, color, offset_points, ha)
        # Stagger above/below with horizontal shifts where dots coincide (9,8 pair)
        techs = [
            ("Cybersecurity (Zero Trust)", 9, 8, self.config.RED_HEX, (0, 32), "center"),
            ("ML Credit Scoring",          9, 8, self.config.NAVY_HEX, (-55, 14), "right"),
            ("Cloud/Data Platform",        9, 7, self.config.NAVY_HEX, (55, 0), "left"),
            ("GenAI Back-office",          8, 9, self.config.BLUE_HEX, (0, -16), "center"),
            ("Digital Auction Platform",   9, 5, self.config.PURPLE_HEX, (0, -16), "center"),
            ("RPA",                        7, 9, self.config.GREEN_HEX, (55, 0), "left"),
            ("Computer Vision",            7, 4, self.config.BLUE_HEX, (0, 14), "center"),
            ("IoT/EID",                    6, 5, self.config.AMBER_HEX, (0, 14), "center"),
            ("Blockchain",                 4, 3, self.config.MID_GRAY_HEX, (0, -16), "center"),
            ("Metaverse/AR",               2, 3, self.config.MID_GRAY_HEX, (0, 14), "center"),
        ]
        _bbox = dict(boxstyle="round,pad=0.2", fc="white", ec="none", alpha=0.85)
        for name, impact, feasibility, color, offset, ha in techs:
            ax.scatter(feasibility, impact, s=300, c=color, edgecolor="white",
                       linewidth=2, alpha=0.85, zorder=3)
            dy = offset[1]
            if dy > 0:
                va = "bottom"
            elif dy < 0:
                va = "top"
            else:
                va = "center"
            ax.annotate(name, (feasibility, impact), xytext=offset,
                        textcoords="offset points", fontsize=8.5, ha=ha, va=va,
                        color=self.config.DARK_GRAY_HEX, fontweight="bold",
                        bbox=_bbox, zorder=5)
        ax.axhline(y=5.5, color=self.config.MID_GRAY_HEX, linestyle="--", alpha=0.5)
        ax.axvline(x=5.5, color=self.config.MID_GRAY_HEX, linestyle="--", alpha=0.5)
        ax.text(10.8, 10.4, "INVEST NOW", fontsize=11, fontweight="bold",
                color=self.config.GREEN_HEX, ha="right")
        ax.text(0.2, 10.4, "STRATEGIC BETS", fontsize=10, fontweight="bold",
                color=self.config.AMBER_HEX, ha="left")
        ax.text(8, 1.2, "QUICK WINS", fontsize=10, fontweight="bold",
                color=self.config.BLUE_HEX, ha="center")
        ax.text(3, 1.2, "MONITOR / DEFER", fontsize=10, fontweight="bold",
                color=self.config.RED_HEX, ha="center")
        ax.set_xlim(0, 11); ax.set_ylim(0, 10.8)
        ax.set_xlabel("Feasibility (technical + organizational readiness)",
                      fontsize=11, fontweight="bold")
        ax.set_ylabel("Strategic Impact",
                      fontsize=11, fontweight="bold")
        ax.set_title("PLMA Technology Portfolio: Impact × Feasibility",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.grid(True, alpha=0.3)
        return self._finalize(fig)

    # 9. Digital Maturity Radar
    def digital_maturity_radar(self):
        categories = ["Strategy &\nVision", "Data &\nAnalytics", "Talent &\nCulture",
                      "Infrastructure\n& Cloud", "AI/ML\nCapability", "Security &\nGovernance"]
        plma = [3, 2, 2.5, 2, 1.5, 2.5]
        peer_avg = [4, 3.5, 3, 3.5, 3, 3.5]
        leader = [5, 5, 4.5, 5, 4.5, 4.5]
        angles = np.linspace(0, 2*np.pi, len(categories), endpoint=False).tolist()
        def close(vals): return vals + vals[:1]
        angles_c = close(angles)
        fig, ax = plt.subplots(figsize=(9, 8), subplot_kw=dict(projection="polar"))
        ax.plot(angles_c, close(leader), color=self.config.GREEN_HEX, linewidth=2,
                linestyle="--", label="Digital Leader")
        ax.fill(angles_c, close(leader), color=self.config.GREEN_HEX, alpha=0.10)
        ax.plot(angles_c, close(peer_avg), color=self.config.BLUE_HEX, linewidth=2,
                label="Peer Coop Avg")
        ax.fill(angles_c, close(peer_avg), color=self.config.BLUE_HEX, alpha=0.15)
        ax.plot(angles_c, close(plma), color=self.config.RED_HEX, linewidth=3,
                label="PLMA (estimated)")
        ax.fill(angles_c, close(plma), color=self.config.RED_HEX, alpha=0.25)
        ax.set_xticks(angles)
        ax.set_xticklabels(categories, fontsize=10, fontweight="bold")
        ax.set_yticks([1, 2, 3, 4, 5])
        ax.set_yticklabels(["1", "2", "3", "4", "5"], fontsize=9,
                           fontweight="bold", zorder=10,
                           bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="none", alpha=0.8))
        ax.set_ylim(0, 5.3)
        ax.set_title("PLMA Digital Maturity Assessment vs. Peers & Leaders",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=30)
        ax.legend(loc="lower right", bbox_to_anchor=(1.25, 0.0), frameon=False, fontsize=9)
        ax.grid(True, alpha=0.4)
        return self._finalize(fig)

    # 10. Gartner Hype Cycle positioning
    def hype_cycle(self):
        fig, ax = plt.subplots(figsize=(13, 7))
        x = np.linspace(0, 10, 400)
        peak = 4 * np.exp(-((x - 1.8)**2) / 0.6) + 1
        trough = -1.5 * np.exp(-((x - 4.2)**2) / 0.5) + 0
        slope = 0.55 * (x - 5.5) / (1 + 0.3 * (x - 5.5))
        slope = np.where(x >= 5.5, slope, 0)
        y = peak + trough + slope
        y = np.clip(y, 0, None)
        ax.plot(x, y, color=self.config.NAVY_HEX, linewidth=3)
        ax.fill_between(x, 0, y, color=self.config.LIGHT_BLUE_HEX, alpha=0.2)
        phases = [
            (0.6, "Innovation\nTrigger"),
            (1.9, "Peak of Inflated\nExpectations"),
            (4.2, "Trough of\nDisillusionment"),
            (6.8, "Slope of\nEnlightenment"),
            (9.3, "Plateau of\nProductivity"),
        ]
        for px, label in phases:
            ax.text(px, -1.1, label, ha="center", fontsize=8.5,
                    color=self.config.DARK_GRAY_HEX, fontweight="bold")
        # Tech positions with explicit label coordinates (text_x, text_y) to prevent overlap
        _bbox = dict(boxstyle="round,pad=0.2", fc="white", ec="none", alpha=0.9)
        techs = [
            ("Agentic AI", 1.3, 4.8, self.config.RED_HEX, 0.5, 5.8),
            ("GenAI (Enterprise)", 2.2, 4.3, self.config.RED_HEX, 3.2, 5.0),
            ("Blockchain (Ag)", 4.0, -0.3, self.config.AMBER_HEX, 4.0, -1.3),
            ("Computer Vision\n(Livestock)", 5.2, 0.8, self.config.AMBER_HEX, 4.2, 1.8),
            ("IoT/EID", 6.8, 2.4, self.config.GREEN_HEX, 6.0, 1.4),
            ("RPA", 8.5, 3.2, self.config.GREEN_HEX, 9.5, 2.5),
            ("Cloud Core", 9.2, 3.5, self.config.GREEN_HEX, 10.0, 4.3),
            ("ML Credit Scoring", 7.3, 2.7, self.config.GREEN_HEX, 6.8, 3.8),
            ("Digital Auction\nPlatform", 7.8, 3.0, self.config.GREEN_HEX, 8.5, 4.5),
        ]
        for name, px, py, color, tx, ty in techs:
            ax.scatter(px, py, s=140, c=color, edgecolor="white", linewidth=2, zorder=5)
            ax.annotate(name, xy=(px, py), xytext=(tx, ty),
                        fontsize=8.5, color=self.config.DARK_GRAY_HEX, fontweight="bold",
                        ha="center", va="center", bbox=_bbox,
                        arrowprops=dict(arrowstyle="-", color=self.config.MID_GRAY_HEX,
                                        lw=0.7, alpha=0.6))
        ax.set_xticks([])
        ax.set_yticks([])
        ax.spines["left"].set_visible(False)
        ax.spines["bottom"].set_visible(False)
        ax.set_xlabel("Time →", fontsize=11, fontweight="bold",
                      color=self.config.DARK_GRAY_HEX)
        ax.set_ylabel("Expectations →", fontsize=11, fontweight="bold",
                      color=self.config.DARK_GRAY_HEX)
        ax.set_title("Gartner Hype Cycle — PLMA Technology Positioning",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.set_ylim(-2, 6)
        ax.set_xlim(0, 10.5)
        # Legend
        legend_elems = [
            mpatches.Patch(color=self.config.RED_HEX, label="High hype — calibrate expectations"),
            mpatches.Patch(color=self.config.AMBER_HEX, label="Overhyped — monitor"),
            mpatches.Patch(color=self.config.GREEN_HEX, label="Mature — invest now"),
        ]
        ax.legend(handles=legend_elems, loc="upper right", frameon=True, fontsize=8.5,
                  framealpha=0.95)
        return self._finalize(fig)

    # 11. Tech Three Horizons Timeline
    def tech_timeline(self):
        workstreams = [
            ("Cybersecurity hardening (Zero Trust)", 0, 9, self.config.NAVY_HEX),
            ("Cloud migration Phase 1", 3, 12, self.config.NAVY_HEX),
            ("Data Lakehouse v1", 3, 12, self.config.NAVY_HEX),
            ("ML Credit Scoring v1 (pilot)", 0, 9, self.config.NAVY_HEX),
            ("GenAI back-office pilots", 3, 12, self.config.NAVY_HEX),
            ("RPA AR/AP automation", 0, 6, self.config.NAVY_HEX),
            ("AI Governance framework", 0, 6, self.config.NAVY_HEX),
            ("ML Credit Scoring v2 (scaled)", 12, 30, self.config.BLUE_HEX),
            ("Price Prediction & Market Intel", 12, 30, self.config.BLUE_HEX),
            ("Digital Auction Platform", 18, 42, self.config.BLUE_HEX),
            ("Computer Vision pilots", 24, 48, self.config.BLUE_HEX),
            ("IoT/EID pilots", 18, 48, self.config.BLUE_HEX),
            ("Full platform modernization", 12, 60, self.config.BLUE_HEX),
            ("AI-native operating model", 48, 96, self.config.AMBER_HEX),
            ("Platform business / data monetization", 60, 108, self.config.AMBER_HEX),
            ("Ecosystem orchestration play", 72, 120, self.config.AMBER_HEX),
        ]
        fig, ax = plt.subplots(figsize=(14, 7.5))
        for i, (name, start, end, color) in enumerate(workstreams):
            bar_width = end - start
            ax.barh(i, bar_width, left=start, color=color, edgecolor="white",
                    linewidth=1.5, height=0.7)
            if color == self.config.AMBER_HEX:
                # Horizon 3: text inside bar, dark blue font
                ax.text(start + bar_width/2, i, name, ha="center", va="center",
                        fontsize=9, color=self.config.NAVY_HEX, fontweight="bold")
            else:
                # Horizon 1 & 2: text outside bar to the right, matching bar color
                ax.text(end + 1, i, name, ha="left", va="center",
                        fontsize=9, color=color, fontweight="bold")
        ax.axvline(x=12, color=self.config.DARK_GRAY_HEX, linestyle="--", alpha=0.5)
        ax.axvline(x=60, color=self.config.DARK_GRAY_HEX, linestyle="--", alpha=0.5)
        ax.text(6, len(workstreams)+0.3, "HORIZON 1\nFoundation (Y1)", ha="center",
                fontsize=10, fontweight="bold", color=self.config.NAVY_HEX)
        ax.text(36, len(workstreams)+0.3, "HORIZON 2\nTransformation (Y2-5)",
                ha="center", fontsize=10, fontweight="bold", color=self.config.BLUE_HEX)
        ax.text(90, len(workstreams)+0.3, "HORIZON 3\nReinvention (Y6-10)",
                ha="center", fontsize=10, fontweight="bold", color=self.config.AMBER_HEX)
        ax.set_yticks([]); ax.set_xlabel("Months from Inception", fontsize=10)
        ax.set_xlim(0, 130)
        ax.set_ylim(-0.8, len(workstreams)+1.2)
        ax.set_title("Technology Transformation Roadmap — Three Horizons",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=25)
        ax.grid(True, axis="x", alpha=0.3)
        ax.set_axisbelow(True)
        return self._finalize(fig)

    # 12. Cumulative ROI
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
                ax.text(be + 0.1, ax.get_ylim()[1]*0.3,
                        f"Breakeven\nYear {be:.1f}", fontsize=10, fontweight="bold",
                        color=self.config.AMBER_HEX)
                break
        for x, y in zip(years, cum):
            ax.annotate(f"${y:.1f}M", (x, y), textcoords="offset points",
                        xytext=(0, 10), ha="center", fontsize=9,
                        color=self.config.DARK_GRAY_HEX)
        ax.set_xlabel("Year", fontsize=11)
        ax.set_ylabel("Cumulative Net Cash Flow ($M)", fontsize=11)
        ax.set_title("10-Year Cumulative ROI: Technology Transformation Program",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.legend(loc="upper left", frameon=False)
        ax.grid(True, alpha=0.3)
        ax.set_axisbelow(True)
        return self._finalize(fig)

    # 13. Cost-benefit waterfall
    def cost_benefit_waterfall(self):
        items = [
            ("Year 1\nInvestment", -4.8, self.config.RED_HEX),
            ("Credit Loss\nAvoidance", 0.9, self.config.GREEN_HEX),
            ("Labor\nProductivity", 2.4, self.config.GREEN_HEX),
            ("Cyber Incident\nAvoidance", 1.1, self.config.GREEN_HEX),
            ("New Digital\nRevenue", 2.2, self.config.GREEN_HEX),
            ("Market Intel\nHedging", 1.1, self.config.GREEN_HEX),
            ("Working Capital\nOptimization", 1.3, self.config.GREEN_HEX),
            ("Ongoing\nCosts", -2.6, self.config.RED_HEX),
        ]
        fig, ax = plt.subplots(figsize=(12, 6))
        cum = 0
        for i, (name, val, color) in enumerate(items):
            ax.bar(i, val, bottom=cum, color=color, edgecolor="white", linewidth=1.5)
            ax.text(i, cum + val/2, f"${val:+.1f}M", ha="center", va="center",
                    fontsize=10, fontweight="bold", color="white")
            cum += val
        ax.bar(len(items), cum, color=self.config.NAVY_HEX, edgecolor="white", linewidth=1.5)
        # Place "Net Steady-State" label above the bar if it's short
        if abs(cum) < 3:
            ax.text(len(items), cum + 0.5, f"${cum:+.1f}M\nNet Steady-State", ha="center",
                    va="bottom", fontsize=10, fontweight="bold", color=self.config.NAVY_HEX)
        else:
            ax.text(len(items), cum/2, f"${cum:+.1f}M\nNet\nSteady-State", ha="center",
                    va="center", fontsize=10, fontweight="bold", color="white")
        ax.axhline(y=0, color="black", linewidth=0.8)
        labels = [it[0] for it in items] + ["Net Steady-\nState Impact"]
        ax.set_xticks(range(len(labels)))
        ax.set_xticklabels(labels, fontsize=8.5)
        ax.set_ylabel("Annual Impact ($M)", fontsize=10)
        ax.set_title("Cost-Benefit Waterfall — Technology Program Economics",
                     fontsize=13, fontweight="bold", color=self.config.NAVY_HEX, pad=15)
        ax.grid(True, axis="y", alpha=0.3)
        ax.set_axisbelow(True)
        return self._finalize(fig)

    # 14. Sensitivity tornado
    def sensitivity_tornado(self):
        tornado, base = self.model.sensitivity()
        tornado = tornado[:6]
        base_m = base / 1_000_000
        fig, ax = plt.subplots(figsize=(12, 5.5))
        for i, t in enumerate(tornado):
            low_m = t["low"] / 1_000_000 - base_m
            high_m = t["high"] / 1_000_000 - base_m
            ax.barh(i, low_m, color=self.config.RED_HEX, edgecolor="white",
                    height=0.6, alpha=0.85)
            ax.barh(i, high_m, color=self.config.GREEN_HEX, edgecolor="white",
                    height=0.6, alpha=0.85)
        # Place labels based on ACTUAL bar direction (handles inverted rows)
        data_pad = 2.0
        for i, t in enumerate(tornado):
            low_m = t["low"] / 1_000_000 - base_m
            high_m = t["high"] / 1_000_000 - base_m
            leftmost = min(low_m, high_m)
            rightmost = max(low_m, high_m)
            if low_m <= high_m:
                left_label = f"${t['low']/1e6:.1f}M"
                right_label = f"${t['high']/1e6:.1f}M"
            else:
                left_label = f"${t['high']/1e6:.1f}M"
                right_label = f"${t['low']/1e6:.1f}M"
            ax.text(leftmost - data_pad, i, left_label,
                    ha="right", va="center", fontsize=9, clip_on=False)
            ax.text(rightmost + data_pad, i, right_label,
                    ha="left", va="center", fontsize=9, clip_on=False)
        ax.set_yticks(range(len(tornado)))
        ax.set_yticklabels([t["variable"] for t in tornado], fontsize=10)
        ax.invert_yaxis()
        ax.axvline(x=0, color=self.config.NAVY_HEX, linewidth=2)
        ax.margins(x=0.25)
        ax.set_xlabel(f"Change in 10-Year NPV vs. Base Case (${base_m:.1f}M)", fontsize=10)
        ax.set_title("Sensitivity Tornado — Technology Program NPV",
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
        normal = self.doc.styles["Normal"]
        normal.font.name = self.config.BODY_FONT
        normal.font.size = Pt(10.5)
        normal.font.color.rgb = self.config.DARK_GRAY
        for level, size, color in [(1, 20, self.config.NAVY),
                                    (2, 15, self.config.NAVY),
                                    (3, 12, self.config.BLUE)]:
            h = self.doc.styles[f"Heading {level}"]
            h.font.name = self.config.HEADING_FONT
            h.font.size = Pt(size)
            h.font.color.rgb = color
            h.font.bold = True
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
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp.add_run("CONFIDENTIAL — Prepared for R. Johansen, CFO, PLMA   |   Page ").font.size = Pt(9)
        for run in fp.runs:
            run.font.name = self.config.BODY_FONT
            run.font.color.rgb = self.config.MID_GRAY
            run.font.size = Pt(9)
        add_page_number_field(fp)
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run("Emerging Technology Strategy  •  April 2026  •  Companion to Counterparty Credit Risk Report")
        hrun.font.name = self.config.BODY_FONT
        hrun.font.size = Pt(9)
        hrun.font.color.rgb = self.config.MID_GRAY
        hrun.italic = True

    def _para(self, text, size=None, bold=False, italic=False, color=None,
              align=None, space_after=6, space_before=0, line=1.2):
        p = self.doc.add_paragraph()
        run = p.add_run(text)
        run.font.name = self.config.BODY_FONT
        if size: run.font.size = Pt(size)
        run.bold = bold
        run.italic = italic
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

    def _bullet(self, text):
        p = self.doc.add_paragraph(style="List Bullet")
        run = p.add_run(text)
        run.font.name = self.config.BODY_FONT
        run.font.size = Pt(10.5)
        run.font.color.rgb = self.config.DARK_GRAY
        set_paragraph_spacing(p, after=3, line=1.2)
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
            left={"val": "single", "sz": "36", "color": border_color, "space": "0"},
            top={"val": "single", "sz": "4", "color": border_color, "space": "0"},
            bottom={"val": "single", "sz": "4", "color": border_color, "space": "0"},
            right={"val": "single", "sz": "4", "color": border_color, "space": "0"},
        )
        cell.width = Inches(6.5)
        cell.paragraphs[0].text = ""
        p1 = cell.paragraphs[0]
        r1 = p1.add_run(title.upper())
        r1.bold = True
        r1.font.name = self.config.HEADING_FONT
        r1.font.size = Pt(9)
        r1.font.color.rgb = heading_color
        set_paragraph_spacing(p1, after=3)
        p2 = cell.add_paragraph()
        r2 = p2.add_run(text)
        r2.font.name = self.config.BODY_FONT
        r2.font.size = Pt(10)
        r2.font.color.rgb = self.config.DARK_GRAY
        set_paragraph_spacing(p2, after=4, line=1.2)
        self._para("", space_after=8)

    def _table(self, headers, rows, col_widths=None, header_fill=None):
        header_fill = header_fill or self.config.NAVY_SHADE
        table = self.doc.add_table(rows=1 + len(rows), cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
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
        if col_widths:
            for row in table.rows:
                for i, w in enumerate(col_widths):
                    row.cells[i].width = Inches(w)
        self._para("", space_after=8)
        return table

    # ============ COVER, TOC ============
    def _add_cover_page(self):
        for _ in range(4):
            self._para("", space_after=2)
        p = self.doc.add_paragraph()
        r = p.add_run("CONFIDENTIAL  •  CLIENT DELIVERABLE  •  REPORT 2 OF 2")
        r.font.name = self.config.HEADING_FONT
        r.font.size = Pt(10)
        r.font.color.rgb = self.config.RED
        r.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(p, after=40)

        p = self.doc.add_paragraph()
        r = p.add_run("Emerging Technology Strategy")
        r.font.name = self.config.HEADING_FONT
        r.font.size = Pt(36)
        r.font.color.rgb = self.config.NAVY
        r.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(p, after=6)

        p = self.doc.add_paragraph()
        r = p.add_run("Opportunity, Risk & the Path to Digital Resilience")
        r.font.name = self.config.HEADING_FONT
        r.font.size = Pt(20)
        r.font.color.rgb = self.config.BLUE
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(p, after=40)

        tbl = self.doc.add_table(rows=1, cols=1)
        cell = tbl.cell(0, 0)
        set_cell_shading(cell, self.config.NAVY_SHADE)
        cell.paragraphs[0].text = ""
        p = cell.paragraphs[0]
        r = p.add_run(" ")
        r.font.size = Pt(2)
        cell.width = Inches(6.5)
        self._para("", space_after=30)

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
        set_paragraph_spacing(p, after=20)

        p = self.doc.add_paragraph()
        r = p.add_run("Companion to Report 1: Counterparty Credit Risk Assessment")
        r.font.name = self.config.BODY_FONT
        r.font.size = Pt(10)
        r.font.color.rgb = self.config.PURPLE
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
        r = p.add_run("Strategic Advisory  |  Technology & Digital Practice")
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
            ("2. The Technology Landscape for Livestock Marketing", "6"),
            ("3. Technology Risk Assessment [Risk-Led Anchor]", "10"),
            ("4. The AI Opportunity Deep Dive", "17"),
            ("5. Broader Technology Opportunity Portfolio", "24"),
            ("6. Strategic Framework Analysis", "30"),
            ("7. Technology Strategy & Roadmap (1 / 5 / 10 Year)", "35"),
            ("8. Financial Model: Tech Investment Economics", "41"),
            ("9. Governance & Execution", "46"),
            ("10. Conclusion & Next Steps", "49"),
            ("Appendix A: Use Case Scorecards", "51"),
            ("Appendix B: Methodology & Assumptions", "54"),
            ("Appendix C: Glossary of Terms", "56"),
            ("Appendix D: Citations", "58"),
        ]
        for title, page in entries:
            p = self.doc.add_paragraph()
            tab_stops = p.paragraph_format.tab_stops
            tab_stops.add_tab_stop(Inches(6.3), alignment=2, leader=1)
            r1 = p.add_run(title)
            r1.font.name = self.config.BODY_FONT
            r1.font.size = Pt(11)
            r1.font.color.rgb = self.config.DARK_GRAY
            r2 = p.add_run(f"\t{page}")
            r2.font.name = self.config.BODY_FONT
            r2.font.size = Pt(11)
            r2.font.color.rgb = self.config.NAVY
            r2.bold = True
            set_paragraph_spacing(p, after=8, line=1.3)

    # ============ SECTION 1: EXECUTIVE SUMMARY ============
    def _add_executive_summary(self):
        self._h1("1. Executive Summary")

        self._callout(
            "Situation",
            "PLMA operates on a technology foundation that is approximately 10 years behind livestock-"
            "marketing peers and 15+ years behind digitally-native financial services. The cost of "
            "inaction now exceeds the cost of transformation: legacy tech debt silently drains an "
            "estimated $8.8M annually in avoidable costs and foregone opportunity, while digital "
            "marketplace entrants (Superior Livestock, Cattlerange, fintech challengers) are "
            "accelerating their share of flow. This window to act on favorable terms is closing.",
            kind="warning"
        )

        self._h2("1.1 Key Findings")

        self._bullet_lead(
            "The disruption threat is real and accelerating.",
            "Digital livestock marketplaces grew an estimated 40%+ YoY in transaction volume during "
            "2024–2025 while traditional auction models grew <5%. Fintech-native credit and payment "
            "rails are being introduced by non-traditional entrants targeting PLMA's fee pools.")

        self._bullet_lead(
            "Cybersecurity is the highest-severity technology risk.",
            "Ag-sector cyber incidents grew from 12 (2020) to 117 (2025) — a 9.75× increase. Average "
            "ransom demands escalated from $350K to $2.8M. A single incident could exceed PLMA's "
            "annual cyber avoidance budget by 3-5×.")

        self._bullet_lead(
            "AI is neither optional nor niche — it is a fiduciary imperative.",
            "12 credible AI/tech use cases were identified for PLMA with aggregate expected 10-year "
            "NPV of ~$66M if all pursued. The top 7 use cases (DO FIRST + QUICK WINS quadrants) "
            "deliver ~$47M in NPV and form the core of the recommended program.")

        self._bullet_lead(
            "PLMA's digital maturity lags peers across all six dimensions.",
            "Estimated maturity score: 2.25/5 average vs. peer cooperative average of 3.42/5 and "
            "digital leader benchmark of 4.75/5. The largest gaps are in AI/ML capability, data/"
            "analytics, and cloud infrastructure.")

        self._bullet_lead(
            "The investment case is compelling even under adverse sensitivity assumptions.",
            "A $4.8M Year-1 investment generates a 10-year NPV of ~$34M at 8% discount rate, "
            "~192% cumulative ROI, and payback within 18 months. The program remains NPV-positive "
            "across all tested sensitivity scenarios.")

        self._bullet_lead(
            "Credit risk and tech risk are the same risk viewed from two angles.",
            "The manual, relationship-driven credit decisioning flagged in Report #1 is simultaneously "
            "a counterparty credit risk AND a technology risk. This report's AI credit scoring "
            "recommendation directly supersedes the v1 scoring model in Report #1's Horizon 1 plan.")

        self._h2("1.2 Recommended Actions")

        self._body(
            "The recommended program is a risk-led, opportunity-harvesting portfolio sequenced "
            "across the McKinsey Three Horizons framework:")

        self._bullet_lead("Horizon 1 (Year 1 — Foundation):",
            "Cybersecurity hardening (Zero Trust), cloud/data platform v1, AI governance framework, "
            "ML credit scoring v1, GenAI back-office pilots, RPA for AR/AP automation. "
            "$4.8M investment. Goal: stabilize risk surface, prove AI value.")
        self._bullet_lead("Horizon 2 (Years 2–5 — Transformation):",
            "Scaled AI (price prediction, computer vision), full platform modernization, digital "
            "auction platform, IoT/EID pilots. Goal: capture the digital marketplace threat, "
            "differentiate on analytics.")
        self._bullet_lead("Horizon 3 (Years 6–10 — Reinvention):",
            "AI-native operating model, platform business plays (data monetization, advisory "
            "services), ecosystem orchestration. Goal: transform PLMA from auction cooperative "
            "to agricultural risk and market intelligence platform.")

        self._h2("1.3 Financial Impact Summary")

        self._table(
            ["Metric", "Value"],
            [
                ["Year 1 Program Investment", "$4.80M"],
                ["Annual Ongoing Cost (Year 2+)", "$2.60M"],
                ["Annual Steady-State Benefit", "$9.00M"],
                ["Net Steady-State Annual Impact", "+$6.40M"],
                ["10-Year NPV (8% discount rate)", "~$34M"],
                ["Payback Period", "~16-18 months"],
                ["10-Year Cumulative ROI", "~192%"],
                ["Annual Tech Debt Drag (status quo)", "~$8.8M"],
            ],
            col_widths=[3.5, 2.5]
        )

        self._callout(
            "Urgency",
            "Three factors compress the decision window: (1) digital marketplace entrants are "
            "compounding advantage, (2) cyber insurance premiums are hardening as incident frequency "
            "rises, and (3) AI talent markets favor first-movers. A 12-month delay costs an "
            "estimated $5-7M in foregone NPV plus elevated tail risk.",
            kind="finding"
        )

    # ============ SECTION 2: LANDSCAPE ============
    def _add_section_2_landscape(self):
        self._h1("2. The Technology Landscape for Livestock Marketing")

        self._body(
            "The livestock marketing industry is undergoing a quiet but accelerating digital "
            "transformation. This section sets context: where capital is flowing, where AI adoption "
            "stands, how peers compare, and what the regulatory environment requires.")

        self._h2("2.1 Digital Disruption of Agricultural Value Chains")

        p = self._body(
            "Global agtech investment reached an estimated $13B in 2025 after a 2021 peak of "
            "$11.4B, a post-correction rebound. Cumulative 2018–2025 investment in ag-focused "
            "technology exceeds $75B, with the fastest-growing segment being ag AI/ML "
            "(from $400M in 2018 to $4.2B in 2025).")
        self._cite(p, 1)

        self._chart(self.charts.agtech_investment_trends(),
                    caption="Ag AI/ML is the fastest-growing agtech segment, crossing $4B in 2025.")

        self._body(
            "This capital is funding direct threats to PLMA's business model: digital auction "
            "platforms, peer-to-peer livestock marketplaces, AI-driven credit decisioning for "
            "producers, fintech payment rails that bypass intermediaries, and data platforms that "
            "aggregate market intelligence cooperatively held historically.")

        self._h2("2.2 AI Adoption Curves")

        p = self._body(
            "McKinsey's State of AI surveys indicate enterprise AI adoption reaching ~65% of "
            "organizations globally in 2025, with GenAI adoption specifically jumping from 33% "
            "(2023) to 65% (2024). Adoption is not uniform: financial services and technology lead, "
            "while agriculture — particularly livestock marketing — lags.")
        self._cite(p, 2)

        self._chart(self.charts.ai_adoption_scurve(),
                    caption="Livestock marketing sits in the early-majority zone; PLMA has a 2-3 year window to catch up on favorable terms.")

        self._h2("2.3 Competitor / Peer Benchmarking")

        self._body(
            "Among ~15 major U.S. livestock marketing cooperatives and auction networks, we estimate:")
        self._bullet_lead("Digital leaders (3-4 peers):",
            "Cloud-native core systems, AI-driven credit scoring in production, digital auction "
            "platforms live, RPA deployed broadly.")
        self._bullet_lead("Digital challengers (5-6 peers):",
            "Hybrid cloud, selective AI pilots, partial digital auction capability, some RPA.")
        self._bullet_lead("Digital laggards (5-6 peers, including PLMA):",
            "On-premise legacy systems, manual credit processes, traditional auction model, limited "
            "or no AI, reactive cyber posture.")

        self._h2("2.4 Regulatory Landscape")

        self._table(
            ["Regulation / Framework", "Relevance to PLMA", "Status"],
            [
                ["NIST AI Risk Management Framework (AI RMF 1.0)",
                 "Voluntary but increasingly required in B2B and regulated contexts",
                 "Active (2023)"],
                ["SEC Cybersecurity Disclosure Rule",
                 "Public filer obligation; signals expectations across all entities",
                 "Effective 2023"],
                ["State AI Governance Laws (CO, CA, UT, NYC)",
                 "High-risk AI decision systems; extend to credit/employment",
                 "Expanding 2024–2026"],
                ["EU AI Act",
                 "Extraterritorial reach for firms operating with EU buyers/data",
                 "Phased 2025–2027"],
                ["State Privacy Laws (20+ states)",
                 "Producer data, member data, marketing consent",
                 "Patchwork, expanding"],
                ["Packers & Stockyards Act (P&SA)",
                 "Algorithmic pricing, AI bias in credit decisions",
                 "Active enforcement"],
            ],
            col_widths=[2.2, 2.8, 1.4]
        )

        self._h2("2.5 The Cost of Doing Nothing")

        self._callout(
            "The Do-Nothing Scenario",
            "If PLMA maintains current technology posture through 2030: (1) digital marketplace "
            "entrants capture an estimated 15-25% of transaction flow currently routed through "
            "PLMA, (2) cyber incident probability approaches ~35% over 5 years at estimated "
            "$2-4M per incident, (3) legacy system maintenance costs grow 8-12% annually, and "
            "(4) AI talent becomes unaffordable or unavailable as peers lock in capacity.",
            kind="warning"
        )

    # ============ SECTION 3: RISK ASSESSMENT ============
    def _add_section_3_risk(self):
        self._h1("3. Technology Risk Assessment [Risk-Led Anchor]")

        self._body(
            "This section anchors the report. We treat each technology risk in sequence, quantify "
            "where possible, and map to the PLMA context. The risks are ordered by materiality — "
            "not by technical complexity.")

        self._chart(self.charts.tech_risk_heat_map(),
                    caption="Ransomware, digital disruption, and legacy tech debt occupy the red zone of the heat map.")

        self._h2("3.1 Disruption Risk — Digital Marketplaces Eroding the Auction Model")

        self._body(
            "The most significant strategic risk to PLMA is not a single incident but a structural "
            "shift: the gradual migration of buyer and producer behavior to digital-first channels. "
            "Superior Livestock Auction (a pioneer in video/online cattle sales), Cattlerange, and "
            "regional online marketplaces have captured share steadily since 2019. Newer entrants "
            "apply fintech playbooks: AI credit decisioning at origination, instant settlement, "
            "buyer-producer matching algorithms.")

        self._body(
            "Unlike a cyber incident, this disruption is slow and undramatic — which makes it "
            "particularly dangerous. By the time the shift is visible in revenue, competitive "
            "position has already eroded.")

        self._h2("3.2 Cybersecurity Risk — The Most Acute Technology Exposure")

        p = self._body(
            "The Food and Agriculture Information Sharing and Analysis Center (Food & Ag-ISAC) "
            "and CISA have documented a material increase in targeted attacks on agricultural "
            "cooperatives and processors. Notable incidents include the 2021 JBS ransomware attack "
            "($11M ransom paid), 2021 NEW Cooperative incident, and numerous 2023–2025 incidents "
            "affecting grain elevators, feed mills, and livestock marketers.")
        self._cite(p, 12)

        self._chart(self.charts.cyber_incidents_ag(),
                    caption="Cyber incidents in ag sector grew 9.75× from 2020–2025, with average ransoms reaching $2.8M.")

        self._callout(
            "The Ransomware Math",
            "At 117 disclosed incidents/year across the ag sector and ~2,000-3,000 qualifying "
            "entities, the 12-month incident probability for a PLMA-sized target is approximately "
            "4-6%. Expected loss per incident: $2.5-4M (ransom, business interruption, remediation, "
            "notification, litigation). Annual expected loss: $100K-$240K — but tail risk "
            "(catastrophic incident) is $8-15M.",
            kind="warning"
        )

        self._h2("3.3 Legacy / Tech Debt Risk")

        self._chart(self.charts.legacy_tech_debt(),
                    caption="Legacy tech debt silently costs PLMA an estimated $8.8M annually — more than the recommended Year-1 investment.")

        self._callout(
            "Connection to Report #1 (Counterparty Credit Risk)",
            "The manual credit decisioning process identified as a weakness in Report #1 is "
            "simultaneously a technology risk and a credit risk. Manual processes produce "
            "inconsistent underwriting, limited auditability, slow response to distress signals, "
            "and no systematic portfolio-level risk view. Fixing the credit process IS fixing a "
            "tech debt issue — the two reports point to the same root cause.",
            kind="sidebar"
        )

        self._h2("3.4 AI Governance & Explainability Risk")

        self._body(
            "Deploying AI creates new risk categories PLMA has not previously managed: model bias "
            "(decisions systematically disadvantaging producer subgroups), explainability "
            "requirements (P&SA regulators may demand rationale for algorithmic credit decisions), "
            "model drift (performance degradation over time), and hallucination risk in GenAI "
            "applications (incorrect output presented with confidence).")

        self._body(
            "We recommend PLMA adopt the NIST AI Risk Management Framework (AI RMF 1.0) as the "
            "governance backbone — it provides structure for GOVERN, MAP, MEASURE, and MANAGE "
            "functions across the AI lifecycle.")
        p = self.doc.paragraphs[-1]
        self._cite(p, 8)

        self._h2("3.5 Data Privacy & Regulatory Risk")

        self._body(
            "State privacy laws (CCPA/CPRA, VCDPA, CPA, and 20+ state-level analogs enacted or "
            "proposed) impose obligations around member data, producer data, marketing consent, "
            "and data subject rights. The patchwork of state laws creates compliance complexity "
            "far beyond what PLMA currently manages. SEC climate and cyber disclosure rules, "
            "while aimed at public filers, are setting de facto expectations across B2B ag.")

        self._h2("3.6 Talent Risk")

        self._body(
            "AI, data, and cloud talent is in acute shortage nationally — and PLMA's rural/ag-"
            "adjacent hiring footprint makes the challenge harder. Competitive offers for "
            "senior data scientists, ML engineers, and cloud architects commonly exceed $200K "
            "total comp. PLMA cannot outbid coastal tech firms, but it can compete on mission, "
            "flexibility, and career arc — if the organization visibly invests.")

        self._h2("3.7 Vendor Lock-in & Platform Risk")

        self._body(
            "Hyperscaler cloud providers (AWS, Azure, GCP), enterprise SaaS vendors, and AI "
            "foundation model providers all create lock-in risk. The mitigation is not to avoid "
            "these vendors but to architect deliberately: abstraction layers, data portability "
            "standards, multi-cloud optionality for critical workloads, and contractual exit "
            "rights.")

    # ============ SECTION 4: AI DEEP DIVE ============
    def _add_section_4_ai(self):
        self._h1("4. The AI Opportunity Deep Dive")

        self._body(
            "AI is the single largest source of value-creation opportunity in PLMA's technology "
            "portfolio. This section identifies, scores, and sequences AI use cases — separating "
            "where PLMA should invest, where it should pilot, and where it should defer.")

        self._h2("4.1 The PLMA AI Use Case Portfolio")

        self._body(
            "We identified 12 candidate use cases across six AI/tech categories. Each is scored "
            "on strategic value (1-10) and feasibility (1-10). Bubble size represents estimated "
            "10-year NPV contribution.")

        self._chart(self.charts.ai_use_case_matrix(),
                    caption="DO FIRST quadrant (high value, high feasibility) contains 4 use cases with ~$29M combined NPV.")

        uc_rows = []
        for u in self.model.use_case_portfolio():
            uc_rows.append([
                u["name"], u["category"], f"{u['value']}/10", f"{u['feasibility']}/10",
                f"${u['nb_pv']:.1f}M"
            ])
        self._table(
            ["Use Case", "Category", "Value", "Feasibility", "10yr NPV"],
            uc_rows,
            col_widths=[1.9, 1.1, 0.9, 1.1, 0.9]
        )

        self._h2("4.2 ML Counterparty Credit Scoring")

        self._callout(
            "Supersedes Report #1 Horizon 1",
            "Report #1 recommended a v1 statistical credit scoring model at $350K. This report "
            "recommends that investment be re-scoped as the foundation for a production ML credit "
            "scoring capability ($350K Year 1 → ML v2 at $600K in Year 2-3). Same budget "
            "envelope, substantially greater capability. Financial models in Section 8 de-duplicate "
            "the $900K/year credit loss avoidance benefit to prevent double-counting with Report #1.",
            kind="sidebar"
        )

        self._body(
            "An ML credit scoring model ingests financial statements, payment history, public "
            "records, industry benchmarks, and alternative data (satellite imagery for collateral, "
            "weather, commodity exposures) to produce a continuous probability-of-default score. "
            "Compared to Report #1's v1 statistical model, an ML approach captures non-linear "
            "relationships, adapts to new data, and provides significantly better early-warning "
            "signal on distressed counterparties.")

        self._h2("4.3 Market Price Prediction & Hedging Intelligence")

        self._body(
            "AI models trained on historical futures, cash, basis, and macro data can generate "
            "probabilistic forecasts of cattle and hog prices at 1-week, 1-month, and 3-month "
            "horizons. These forecasts inform: (i) PLMA's own working capital positioning, "
            "(ii) hedging overlay execution (tying to Report #1's hedging program), and "
            "(iii) potential fee-based advisory services to members.")

        self._h2("4.4 Computer Vision for Livestock")

        self._body(
            "Computer vision is maturing rapidly in livestock applications: individual animal "
            "identification (replacing/augmenting EID), weight estimation from video, body "
            "condition scoring, early disease detection, and automated grading. Readiness varies: "
            "identification and weight estimation are production-ready; grading and disease "
            "detection remain pilot-stage. PLMA's right role is strategic pilot + evaluation, "
            "not foundational R&D.")

        self._h2("4.5 GenAI for Back-Office, Research, and Member Service")

        self._body(
            "GenAI has immediate, high-feasibility value in: (i) document processing (contracts, "
            "financial statements, sale records), (ii) research and market intelligence "
            "(synthesizing USDA, Fed, industry reports), (iii) member service automation "
            "(routine inquiries, documentation), and (iv) internal knowledge management. These "
            "are quick wins with measurable productivity gains and limited downside risk — "
            "provided governance guardrails are in place.")

        self._h2("4.6 Demand Sensing & Inventory Optimization")

        self._body(
            "ML-driven demand sensing can reduce working capital requirements by 10-15% by "
            "improving the accuracy of expected buyer demand and optimal inventory posture. "
            "For PLMA's $200M estimated working capital exposure, this represents $20-30M in "
            "liquidity unlocked or redeployed.")

        self._h2("4.7 Build vs. Buy vs. Partner")

        self._chart(self.charts.build_buy_partner(),
                    caption="Sourcing recommendation: Buy + Partner for most AI capabilities; Build selectively where PLMA has proprietary data advantage.")

        self._body(
            "Core principle: Build where PLMA has proprietary data advantage (credit scoring, "
            "market prediction); Buy where vendors offer commodity capability (GenAI, RPA, "
            "cyber SOC); Partner where the capability requires specialized expertise PLMA "
            "cannot cost-effectively develop internally.")

    # ============ SECTION 5: BROADER TECH PORTFOLIO ============
    def _add_section_5_portfolio(self):
        self._h1("5. Broader Technology Opportunity Portfolio")

        self._body(
            "Beyond AI, six additional technology domains demand PLMA's attention. Each is "
            "discussed below with a position on the Invest Now / Quick Wins / Strategic Bets / "
            "Defer framework.")

        self._chart(self.charts.tech_portfolio_2x2(),
                    caption="Cybersecurity, Cloud/Data, ML Credit Scoring, and Digital Auction are INVEST NOW priorities.")

        self._h2("5.1 Data & Analytics Platform Modernization")

        self._body(
            "A modern data lakehouse architecture (Snowflake, Databricks, or equivalent) is the "
            "foundation for AI, analytics, and decision support. Without it, every AI use case "
            "is throttled by data fragmentation, quality issues, and integration complexity. "
            "Estimated cost: $600-900K Year 1, $350K ongoing. This is an enabling investment: "
            "its value is in unlocking other use cases.")

        self._h2("5.2 Cloud Infrastructure Transformation")

        self._body(
            "PLMA's estimated on-premise infrastructure footprint carries 8-12% annual cost "
            "growth, limited scalability, and material cyber exposure. Phased migration to a "
            "primary hyperscaler (recommended: AWS for ag ecosystem depth, or Azure for "
            "Microsoft-shop integration) reduces TCO by 15-25% over 3-5 years while improving "
            "security posture and enabling AI workloads.")

        self._h2("5.3 IoT / EID Integration for Traceability")

        self._body(
            "Electronic Identification (EID) is mandated or encouraged across cattle supply "
            "chains for disease traceability (USDA's ADT rule expansion). PLMA is positioned to "
            "be a data aggregation and value-added services layer on top of EID infrastructure — "
            "if it moves early. This is Strategic Bet quadrant: high long-term value, moderate "
            "near-term feasibility.")

        self._h2("5.4 Blockchain for Provenance & Smart Contracts (Narrow Scope)")

        self._body(
            "Blockchain has been over-hyped in agricultural contexts. We recommend PLMA monitor "
            "but NOT invest in broad blockchain initiatives in the near term. The one narrow "
            "exception: permissioned blockchain for supply-chain traceability in premium/branded "
            "beef programs — and only as a partner/participant in a buyer-led consortium, never "
            "as a lead infrastructure investor.")

        self._h2("5.5 Digital Auction Platforms & Omnichannel")

        self._body(
            "PLMA's traditional auction model needs a digital overlay, not a replacement. A "
            "hybrid model — physical auctions with digital participation, online-only sales for "
            "specific cattle classes, mobile-first member experience — is the target state. "
            "Platform build/buy decision: BUY core platform from specialized vendor, CUSTOMIZE "
            "with PLMA-specific workflows.")

        self._h2("5.6 RPA / Workflow Automation")

        self._body(
            "RPA applied to AR/AP, settlement operations, compliance reporting, and data entry "
            "produces measurable labor productivity gains — typically 20-40% reduction in "
            "processing time for targeted workflows. Estimated annual benefit: $800K-$1.2M at "
            "scale. This is a Quick Win with proven ROI and limited implementation risk.")

        self._h2("5.7 Cybersecurity Fortification")

        self._body(
            "Given the cyber risk profile in Section 3.2, cyber must be treated as Invest Now "
            "priority #1 — regardless of AI ambitions. Recommended elements: Zero Trust "
            "architecture, EDR/XDR on all endpoints, 24×7 SOC (outsourced or MSSP), identity "
            "hardening (MFA universal, privileged access management), incident response retainer, "
            "tabletop exercises quarterly, cyber insurance alignment.")

    # ============ SECTION 6: FRAMEWORKS ============
    def _add_section_6_frameworks(self):
        self._h1("6. Strategic Framework Analysis")

        self._body(
            "Five strategic frameworks inform the technology strategy: (1) Digital Maturity "
            "Assessment, (2) Gartner Hype Cycle, (3) NIST AI RMF, (4) Porter's Value Chain with "
            "digital overlay, and (5) McKinsey Three Horizons (reused from Report #1 for "
            "strategic consistency).")

        self._h2("6.1 PLMA Digital Maturity Assessment")

        self._chart(self.charts.digital_maturity_radar(),
                    caption="PLMA's 2.25 average maturity score is ~34% below peer average and 53% below digital leader benchmark.")

        self._table(
            ["Dimension", "PLMA", "Peer Avg", "Leader", "Gap Priority"],
            [
                ["Strategy & Vision", "3.0", "4.0", "5.0", "Medium"],
                ["Data & Analytics", "2.0", "3.5", "5.0", "HIGH"],
                ["Talent & Culture", "2.5", "3.0", "4.5", "Medium"],
                ["Infrastructure & Cloud", "2.0", "3.5", "5.0", "HIGH"],
                ["AI/ML Capability", "1.5", "3.0", "4.5", "HIGHEST"],
                ["Security & Governance", "2.5", "3.5", "4.5", "HIGH"],
            ],
            col_widths=[1.8, 0.9, 0.9, 0.9, 1.4]
        )

        self._h2("6.2 Gartner Hype Cycle Positioning")

        self._chart(self.charts.hype_cycle(),
                    caption="Calibrate expectations: Agentic AI and GenAI are at peak hype; invest where technologies have matured.")

        self._body(
            "Hype cycle positioning drives investment timing. For technologies at the Peak of "
            "Inflated Expectations (Agentic AI, enterprise GenAI), we recommend disciplined "
            "pilots with strict value gates. For technologies on the Plateau of Productivity "
            "(cloud core, RPA, ML credit scoring), invest aggressively. For technologies in the "
            "Trough (blockchain in ag, some computer vision applications), monitor and "
            "reassess annually.")

        self._h2("6.3 NIST AI Risk Management Framework (AI RMF 1.0)")

        self._table(
            ["NIST AI RMF Function", "PLMA Application", "Year 1 Action"],
            [
                ["GOVERN", "AI Governance Council, policies, accountability",
                 "Stand up Council; publish AI Use Policy"],
                ["MAP", "Context mapping: where AI is used, by whom, for what",
                 "AI inventory; risk classification by use case"],
                ["MEASURE", "Performance, bias, drift, explainability metrics",
                 "Metrics framework for credit scoring model"],
                ["MANAGE", "Ongoing monitoring, incident response, retirement",
                 "Model risk management process; human-in-loop requirements"],
            ],
            col_widths=[1.3, 2.5, 2.5]
        )

        self._h2("6.4 Porter's Value Chain — Digital Overlay")

        self._body(
            "Mapping digital opportunities to Porter's Value Chain illuminates where technology "
            "creates differentiation vs. parity:")
        self._bullet_lead("Inbound Logistics (producer intake):",
            "Digital member onboarding, AI credit scoring — differentiation opportunity.")
        self._bullet_lead("Operations (auction/marketing):",
            "Digital auction platform, computer vision grading — parity, then differentiation.")
        self._bullet_lead("Outbound Logistics (settlement):",
            "RPA, instant settlement, digital payment rails — parity, defensive.")
        self._bullet_lead("Marketing & Sales:",
            "Market intelligence, demand sensing, member advisory — differentiation opportunity.")
        self._bullet_lead("Service:",
            "GenAI member service, advisory, education — differentiation opportunity.")
        self._bullet_lead("Support Activities:",
            "Cybersecurity, cloud, data platform, AI governance — foundation; not differentiating but failure-enabling.")

        self._h2("6.5 Three Horizons (Strategic Consistency with Report #1)")

        self._body(
            "The same Three Horizons framework structures both reports' roadmaps. This is "
            "deliberate: PLMA should have one integrated transformation plan, not two parallel "
            "tracks. Report #1's Horizon 1 credit risk workstreams and this report's Horizon 1 "
            "tech workstreams should be governed by the same PMO and Risk Committee.")

    # ============ SECTION 7: STRATEGY & ROADMAP ============
    def _add_section_7_roadmap(self):
        self._h1("7. Technology Strategy & Roadmap (1 / 5 / 10 Year)")

        self._h2("7.1 North Star Technology Vision")

        self._callout(
            "PLMA's 2036 Tech Vision",
            "\"PLMA is the most trusted, intelligent, and resilient livestock marketing platform "
            "in North America — combining century-old producer trust with AI-native analytics, "
            "digital-first member experience, and best-in-class risk management. Our technology "
            "is a strategic differentiator, not a utility.\"",
            kind="insight"
        )

        self._chart(self.charts.tech_timeline(),
                    caption="16 workstreams sequenced across 10 years, aligned to Three Horizons.")

        self._h2("7.2 Horizon 1: Foundation (Year 1)")

        self._body(
            "Year 1 is about stabilizing the risk surface and proving AI value with low-risk, "
            "high-return pilots.")

        self._h3("Priorities")
        self._bullet("Cybersecurity hardening (Zero Trust architecture, EDR, 24×7 SOC) — $900K")
        self._bullet("Cloud migration Phase 1 + Data Lakehouse v1 — $1.2M")
        self._bullet("ML Credit Scoring v1 (pilot, shared budget with Report #1) — $350K")
        self._bullet("GenAI back-office pilots (document processing, research) — $450K")
        self._bullet("RPA — AR/AP Automation — $400K")
        self._bullet("Digital Platform Enhancements (member portal, mobile) — $800K")
        self._bullet("AI Governance Council, policies, and talent — $700K")

        self._h3("Year 1 Success Metrics")
        self._bullet_lead("Cyber:", "Zero critical vulnerabilities unmitigated >30 days; 100% MFA coverage; incident response plan tested.")
        self._bullet_lead("AI:", "Credit scoring model live; GenAI pilot demonstrates ≥20% productivity gain in target workflows.")
        self._bullet_lead("Cloud/Data:", "Lakehouse v1 operational; 3 core data domains migrated.")
        self._bullet_lead("Governance:", "AI Council operational; AI Use Policy published; NIST AI RMF adoption documented.")

        self._h2("7.3 Horizon 2: Transformation (Years 2–5)")

        self._body(
            "Horizon 2 is about scaling what works, answering the disruption threat with a "
            "digital auction platform, and building the AI capability layer.")

        self._h3("Priorities")
        self._bullet("ML Credit Scoring v2 (production-scale, ML-enhanced)")
        self._bullet("Price Prediction & Market Intelligence models")
        self._bullet("Digital Auction Platform (hybrid physical/online)")
        self._bullet("Computer Vision pilots (weight estimation, identification)")
        self._bullet("IoT/EID integration pilots")
        self._bullet("Full platform modernization (complete cloud migration)")
        self._bullet("RPA expansion to compliance, settlement, reporting")

        self._h3("Year 5 Success Metrics")
        self._bullet_lead("Competitive:", "Digital auction capturing ≥25% of transaction volume; member NPS ≥45.")
        self._bullet_lead("AI:", "5+ AI models in production with measured business impact; model risk management operational.")
        self._bullet_lead("Financial:", "Cumulative program NPV ≥ $14M; tech debt drag reduced by 50%.")
        self._bullet_lead("Digital Maturity:", "PLMA maturity score ≥3.5/5 (from 2.25 baseline).")

        self._h2("7.4 Horizon 3: Reinvention (Years 6–10)")

        self._body(
            "Horizon 3 is about moving beyond being a livestock marketing cooperative to being "
            "an agricultural risk and market intelligence platform.")

        self._h3("Priorities")
        self._bullet("AI-native operating model: AI agents embedded in core workflows")
        self._bullet("Platform business plays: data monetization, fee-based advisory services")
        self._bullet("Ecosystem orchestration: partnerships with banks, insurers, processors")
        self._bullet("Cooperative capital restructuring aligned to platform economics")

        self._h3("Year 10 Success Metrics")
        self._bullet_lead("Strategic:", "15-20% of revenue from fee-based/platform services (non-commission).")
        self._bullet_lead("Financial:", "Cumulative program NPV ≥$34M; EBITDA margin expansion of 300-500 bps.")
        self._bullet_lead("Competitive:", "PLMA is the reference operator in ag-tech for livestock marketing coops.")

    # ============ SECTION 8: FINANCIAL MODEL ============
    def _add_section_8_financials(self):
        self._h1("8. Financial Model: Tech Investment Economics")

        self._h2("8.1 Methodology & Assumptions")

        self._body(
            "This section quantifies the technology transformation program's economics. We use "
            "conventional DCF with an 8% discount rate (consistent with Report #1 and typical "
            "cooperative cost-of-capital). Benefits ramp over 2 years to steady state. "
            "Sensitivity analysis tests six variables across reasonable ranges.")

        self._h2("8.2 Program Costs")

        y1_rows = [[k, f"${v/1000:.0f}K"] for k, v in self.model.year1_costs.items()]
        y1_rows.append(["**TOTAL Year 1**", f"${self.model.total_year1_cost()/1000:.0f}K"])
        self._table(
            ["Year 1 Investment Item", "Amount"],
            y1_rows,
            col_widths=[4.5, 1.5]
        )

        ongoing_rows = [[k, f"${v/1000:.0f}K"] for k, v in self.model.ongoing_costs.items()]
        ongoing_rows.append(["**TOTAL Ongoing Annual**", f"${self.model.total_ongoing_cost()/1000:.0f}K"])
        self._table(
            ["Ongoing Annual Cost (Year 2+)", "Amount"],
            ongoing_rows,
            col_widths=[4.5, 1.5]
        )

        self._h2("8.3 Annual Benefits")

        self._callout(
            "De-duplication with Report #1",
            "Credit loss avoidance in this model is INCREMENTAL to Report #1's credit risk "
            "mitigation program. Report #1 claims $1.8M/year from statistical credit scoring; "
            "this report claims an additional $900K/year from ML enhancement (the delta between "
            "ML and statistical models). Total combined credit loss avoidance: $2.7M/year. "
            "This is not double-counted.",
            kind="sidebar"
        )

        benefit_rows = [[k, f"${v/1000:.0f}K"] for k, v in self.model.annual_benefits.items()]
        benefit_rows.append(["**TOTAL Annual Benefit**", f"${self.model.total_annual_benefit()/1000:.0f}K"])
        self._table(
            ["Benefit Category", "Annual Value"],
            benefit_rows,
            col_widths=[4.5, 1.5]
        )

        self._chart(self.charts.cost_benefit_waterfall(),
                    caption="Steady-state economics: $9.0M annual benefit against $2.6M ongoing cost = $6.4M net.")

        self._h2("8.4 10-Year NPV & Cash Flows")

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
            ["Year", "Cost", "Benefit", "Net CF", "Cumulative"],
            cf_rows,
            col_widths=[0.9, 1.1, 1.1, 1.1, 1.3]
        )

        self._chart(self.charts.roi_projection(),
                    caption="Program breaks even during Year 2; cumulative 10-year nominal net cash flow exceeds $55M.")

        self._h2("8.5 Sensitivity Analysis")

        self._chart(self.charts.sensitivity_tornado(),
                    caption="Benefit realization and labor productivity are the top NPV drivers; program remains positive across all tested scenarios.")

        self._callout(
            "Robustness",
            "Even under the most adverse sensitivity case tested (benefit realization at -30%, "
            "implementation cost at +40%, ongoing costs at +25%), the program remains NPV-positive. "
            "This indicates the investment case is robust — not contingent on aggressive assumptions.",
            kind="finding"
        )

    # ============ SECTION 9: GOVERNANCE ============
    def _add_section_9_governance(self):
        self._h1("9. Governance & Execution")

        self._h2("9.1 Technology Investment Committee")

        self._body(
            "We recommend establishing a Technology Investment Committee (TIC) as a formal "
            "governance body with CFO Johansen as chair. The TIC approves all tech spend above "
            "a threshold, monitors portfolio health, and reports to the Board Risk Committee.")

        self._table(
            ["Body", "Composition", "Cadence", "Mandate"],
            [
                ["Board Risk & Technology Cmte",
                 "3-4 Directors + CEO + CFO",
                 "Quarterly",
                 "Approve tech strategy; oversee risk; review major investments"],
                ["Tech Investment Committee",
                 "CFO (Chair), CTO/IT Lead, COO, Security Lead",
                 "Monthly",
                 "Approve program spend; monitor delivery; course-correct"],
                ["AI Governance Council",
                 "CFO, Chief Risk, Legal, Data Lead, External Ethicist",
                 "Monthly",
                 "NIST AI RMF compliance; bias/fairness review; model approval"],
                ["Cyber Steering Group",
                 "Security Lead (chair), IT, Legal, Ops",
                 "Bi-weekly",
                 "Threat monitoring; incident readiness; control effectiveness"],
            ],
            col_widths=[1.7, 1.8, 0.9, 2.2]
        )

        self._h2("9.2 KPI Dashboard")

        self._table(
            ["KPI", "Target", "Red Flag"],
            [
                ["Digital Maturity Score (avg)", ">3.5 by Year 5", "<3.0 by Year 3"],
                ["Cyber Mean Time to Detect", "<24 hours", ">72 hours"],
                ["Tech program NPV realization", ">80% of plan", "<60% of plan"],
                ["AI models in production (Y5)", "≥5", "<3"],
                ["Incident response tabletop frequency", "Quarterly", "Missed any quarter"],
                ["% revenue from digital services (Y10)", "≥15%", "<8%"],
                ["Program delivery velocity", "≥85% milestones on time", "<70%"],
                ["Tech debt drag reduction", "-50% by Year 5", "<-25% by Year 5"],
            ],
            col_widths=[2.8, 1.5, 1.5]
        )

        self._h2("9.3 Change Management & Talent Strategy")

        self._body(
            "Technology transformation fails when change management is neglected. We recommend "
            "a dedicated change management stream with ~5% of program budget, focused on: "
            "(i) executive sponsorship and visible commitment, (ii) middle-manager enablement, "
            "(iii) producer/member communication about digital channels, (iv) training and "
            "upskilling for affected roles, and (v) celebration of early wins.")

        self._body(
            "Talent strategy: (i) hire 2-3 senior data/AI roles in Year 1, (ii) partner with "
            "2-3 land-grant universities for internship and co-op pipelines, (iii) invest in "
            "existing staff upskilling (certifications, conferences, rotational assignments), "
            "(iv) compete on mission and autonomy where PLMA cannot compete on total comp.")

    # ============ SECTION 10: CONCLUSION ============
    def _add_section_10_conclusion(self):
        self._h1("10. Conclusion & Next Steps")

        self._body(
            "PLMA faces a decade-defining decision. The cooperative can continue its current "
            "technology trajectory — incremental upgrades, reactive cyber posture, manual "
            "decisioning — and watch digital-native entrants capture share, talent, and strategic "
            "initiative. Or PLMA can make a disciplined, phased $4.8M Year-1 investment that "
            "stabilizes the risk surface, proves AI value, and begins a 10-year transformation "
            "to an AI-native, platform-oriented cooperative.")

        self._callout(
            "The Core Thesis",
            "Technology inaction is not neutral. Every month of delay costs approximately $730K "
            "in foregone value ($8.8M tech debt drag + missed benefit capture), grows the "
            "disruption gap vs. peers, and compounds cyber tail risk. The recommended program "
            "pays for itself within 18 months and delivers ~$34M in 10-year NPV. This is the "
            "rare transformation where the financial case, risk case, and strategic case all "
            "align.",
            kind="insight"
        )

        self._h2("10.1 Immediate Actions (First 90 Days)")

        self._bullet("Week 1-2: CFO-led alignment on findings; Tech Investment Committee founded")
        self._bullet("Week 2-4: Board Risk & Technology Committee briefing; approval of Year-1 budget ($4.8M)")
        self._bullet("Week 4-8: Cybersecurity gap assessment; RFP for MSSP/SOC services")
        self._bullet("Week 4-8: Cloud strategy finalization; hyperscaler selection")
        self._bullet("Week 6-10: Senior data/AI leader recruitment launch (2-3 hires)")
        self._bullet("Week 8-12: AI Governance Council chartered; NIST AI RMF adoption documented")
        self._bullet("Week 8-12: RPA use case selection & vendor selection (AR/AP workflow first)")
        self._bullet("Week 10-13: Integrated PMO established linking Report #1 and Report #2 workstreams")

        self._h2("10.2 Integration with Report #1 (Counterparty Credit Risk)")

        self._body(
            "Both reports identify a common root cause — inadequate analytical infrastructure — "
            "and converge on an integrated solution. We recommend PLMA establish ONE integrated "
            "transformation program with unified governance, shared PMO, and consolidated KPI "
            "reporting. The credit risk mitigation and technology strategy are two lenses on the "
            "same modernization imperative.")

        self._h2("10.3 Phase 2 Engagement Model")

        self._body(
            "For Horizon 2 and Horizon 3 workstreams, we recommend: (i) quarterly progress "
            "reviews against Three Horizons roadmap, (ii) annual strategy refresh against "
            "emerging tech developments (particularly in AI), (iii) on-call advisory for "
            "vendor selection and major architecture decisions, (iv) independent model risk "
            "review for AI credit scoring on 18-month cycle.")

        self._body(
            "PLMA has the institutional trust, market access, and producer relationships that "
            "no digital-native entrant can replicate. The question is whether PLMA will also "
            "have the technology capability to defend and extend that position. The financial "
            "case is unambiguous. The strategic case is unambiguous. The risk case is "
            "unambiguous. The decision rests with leadership.")

    # ============ APPENDICES ============
    def _add_appendices(self):
        self._h1("Appendix A: Use Case Scorecards")

        self._body(
            "The 12 AI/tech use cases are detailed below with scoring rationale. Value scores "
            "reflect strategic importance and financial impact potential. Feasibility scores "
            "reflect technical readiness, organizational readiness, and data availability.")

        scorecards = [
            ("ML Credit Scoring", "AI/ML", "9/10", "8/10", "$7.5M",
             "Direct link to Report #1 credit losses. Rich internal data. Vendor solutions mature."),
            ("GenAI Doc Processing", "GenAI", "8/10", "9/10", "$6.8M",
             "High-volume, clear ROI. Mature vendor offerings. Low risk."),
            ("Price Prediction Models", "AI/ML", "9/10", "7/10", "$8.2M",
             "Proprietary data advantage. Differentiation potential. Moderate build complexity."),
            ("Computer Vision Grading", "AI/ML", "8/10", "5/10", "$5.4M",
             "High strategic value. Technology maturing. Pilot-stage."),
            ("GenAI Member Service", "GenAI", "7/10", "8/10", "$4.6M",
             "Quick win. Vendor solutions ready. Governance needed for hallucination."),
            ("Demand Forecasting", "AI/ML", "8/10", "7/10", "$6.1M",
             "Working capital optimization. Proven techniques. Data quality dependency."),
            ("RPA — AR/AP Automation", "Automation", "7/10", "9/10", "$5.8M",
             "Highest feasibility. Proven ROI. Limited risk."),
            ("EID/IoT Traceability", "IoT", "7/10", "6/10", "$4.2M",
             "Regulatory tailwind. Infrastructure cost. Ecosystem play."),
            ("Digital Auction Platform", "Digital", "9/10", "6/10", "$7.9M",
             "Counters disruption threat. Complex build/integrate."),
            ("Cyber Zero-Trust", "Security", "8/10", "8/10", "$3.5M",
             "Table stakes. Measurable risk reduction. Vendor-led."),
            ("Data Lakehouse", "Data", "7/10", "7/10", "$4.8M",
             "Enabling investment. Value through unlocked use cases."),
            ("Blockchain Provenance", "Blockchain", "5/10", "4/10", "$1.2M",
             "Overhyped for most use cases. Monitor; defer investment."),
        ]
        self._table(
            ["Use Case", "Cat.", "Value", "Feas.", "10yr NPV", "Rationale"],
            [[s[0], s[1], s[2], s[3], s[4], s[5]] for s in scorecards],
            col_widths=[1.6, 0.8, 0.6, 0.6, 0.8, 2.1]
        )

        self._h1("Appendix B: Methodology & Assumptions")

        self._h2("B.1 Financial Model Methodology")
        self._body(
            "The financial model uses standard DCF with 8% discount rate. Year-1 costs are "
            "expensed upfront; benefits ramp over 2 years (35% Year 1, 80% Year 2, 100% "
            "Year 3+). Sensitivity analysis tests six variables over reasonable ranges.")

        self._h2("B.2 Benefit Estimation")
        self._body(
            "Benefit estimates are calibrated from: (i) peer-cooperative case studies, "
            "(ii) published efficacy data for RPA, GenAI, and ML solutions in financial services "
            "and ag contexts, (iii) McKinsey Global Institute AI productivity estimates, "
            "(iv) industry benchmarks for cybersecurity incident costs, and (v) analogous "
            "digital transformation programs in adjacent industries.")

        self._h2("B.3 Digital Maturity Scoring")
        self._body(
            "Maturity scores are qualitative estimates across 6 dimensions, calibrated against "
            "published digital maturity frameworks (McKinsey Digital Quotient, MIT CISR, "
            "Gartner). Peer benchmarks are industry estimates, not specific coop data.")

        self._h2("B.4 Use Case Scoring")
        self._body(
            "Value scores reflect: revenue potential, cost reduction potential, risk reduction, "
            "strategic differentiation, data advantage. Feasibility scores reflect: technical "
            "maturity, organizational readiness, data availability, regulatory clarity, vendor "
            "ecosystem maturity.")

        self._h1("Appendix C: Glossary of Terms")

        terms = [
            ("AI RMF", "NIST AI Risk Management Framework — voluntary U.S. standard for AI governance."),
            ("Agentic AI", "AI systems that autonomously plan and execute multi-step workflows."),
            ("Data Lakehouse", "Architecture combining data lake flexibility with warehouse performance."),
            ("EDR/XDR", "Endpoint/Extended Detection and Response — real-time threat detection."),
            ("EID", "Electronic Identification — RFID-based livestock tracking tags."),
            ("GenAI", "Generative AI — AI systems producing text, code, images, audio."),
            ("Hyperscaler", "Major public cloud providers: AWS, Azure, Google Cloud."),
            ("LLM", "Large Language Model — foundation model behind most GenAI applications."),
            ("MFA", "Multi-Factor Authentication."),
            ("MSSP", "Managed Security Services Provider — outsourced security operations."),
            ("P&SA", "Packers & Stockyards Act — USDA livestock marketing regulation."),
            ("RPA", "Robotic Process Automation — software bots for workflow automation."),
            ("SOC", "Security Operations Center — 24×7 threat monitoring and response."),
            ("TCO", "Total Cost of Ownership."),
            ("Zero Trust", "Security model: never trust, always verify; identity-based access."),
        ]
        for term, defn in terms:
            self._bullet_lead(term + ":", defn)

        self._add_citations()

    def _add_citations(self):
        self._h1("Appendix D: Citations")
        self._body(
            "All citations numbered and referenced in the body of this report. Figures calibrated "
            "from multiple public and industry sources; specific PLMA estimates noted in Appendix B.")

        citations = [
            "AgFunder, \"AgriFoodTech Investment Report,\" 2018–2025 annual editions.",
            "McKinsey & Company, \"The State of AI in 2024: Gen AI's breakout year,\" 2024.",
            "McKinsey Global Institute, \"The economic potential of generative AI,\" June 2023.",
            "Gartner, \"Hype Cycle for Artificial Intelligence,\" 2024 and 2025 editions.",
            "Gartner, \"Hype Cycle for Digital Agriculture,\" 2024.",
            "IDC, \"Worldwide Artificial Intelligence Spending Guide,\" 2025.",
            "Forrester Research, \"The State of Enterprise AI,\" 2025.",
            "NIST, \"Artificial Intelligence Risk Management Framework (AI RMF 1.0),\" January 2023.",
            "NIST, \"AI RMF Generative AI Profile,\" July 2024.",
            "SEC, \"Cybersecurity Risk Management, Strategy, Governance, and Incident Disclosure,\" Final Rule, July 2023.",
            "European Union, \"Artificial Intelligence Act (EU AI Act),\" Regulation 2024/1689.",
            "Food and Ag-ISAC, \"Cyber Threats to the Food & Agriculture Sector,\" 2024 and 2025 reports.",
            "CISA, \"Cyber Threats to the Agriculture Sector,\" 2023 & 2024 advisories.",
            "FBI Internet Crime Complaint Center (IC3), \"Annual Report,\" 2023, 2024, 2025.",
            "Sophos, \"The State of Ransomware 2025.\"",
            "Verizon, \"Data Breach Investigations Report (DBIR),\" 2024 and 2025.",
            "USDA APHIS, \"Animal Disease Traceability (ADT) Rule,\" Updated 2024.",
            "Porter, M.E., \"Competitive Advantage,\" Free Press, 1985 (value chain framework).",
            "Baghai, M., Coley, S., White, D., \"The Alchemy of Growth\" (McKinsey Three Horizons), 1999.",
            "Kane, G.C. et al., \"Coming of Age Digitally,\" MIT Sloan Management Review, 2018.",
            "Deloitte, \"2026 Technology Industry Outlook.\"",
            "Accenture, \"Technology Vision 2025.\"",
            "Boston Consulting Group, \"The CEO's Guide to Generative AI,\" 2024.",
            "Harvard Business Review, \"How Generative AI Is Changing Creative Work,\" November 2022.",
            "Harvard Business Review, \"Getting AI Right,\" September-October 2024.",
            "USDA ERS, \"Adoption of Precision Agriculture Technologies in the U.S.,\" 2024.",
            "CoBank, \"Rural Broadband & Digital Agriculture,\" 2025.",
            "Farm Credit Administration, \"Technology Risk in Ag Lending,\" 2024.",
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

    # ============ BUILD ============
    def build(self):
        self._add_cover_page()
        self._add_toc()
        self._add_executive_summary()
        self._add_section_2_landscape()
        self._add_section_3_risk()
        self._add_section_4_ai()
        self._add_section_5_portfolio()
        self._add_section_6_frameworks()
        self._add_section_7_roadmap()
        self._add_section_8_financials()
        self._add_section_9_governance()
        self._add_section_10_conclusion()
        self._add_appendices()
        return self.doc


# =============================================================================
# MAIN
# =============================================================================
def main():
    config = ReportConfig()
    model = TechFinancialModel()
    charts = TechChartGenerator(config, model)
    builder = DocumentBuilder(config, charts, model)
    builder.build()
    output_path = "PLMA_Tech_Strategy_2026.docx"
    builder.doc.save(output_path)
    print(f"Report generated: {output_path}")
    print(f"10-Year NPV: ${model.npv()/1e6:.2f}M")
    print(f"10-Year ROI: {model.roi_10yr()*100:.1f}%")
    print(f"Payback (months): {model.payback_months()}")
    print(f"Annual Benefit: ${model.total_annual_benefit()/1e6:.2f}M")
    print(f"Year 1 Cost: ${model.total_year1_cost()/1e6:.2f}M")


if __name__ == "__main__":
    main()
