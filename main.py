import math
import os
import platform
import subprocess
import json
from dataclasses import dataclass
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk


UOM_OPTIONS = ["MINR", "Assessed Spend %", "Nos."]
MONTH_OPTIONS = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
]

LOCATION_COLORS = {
    "Chennai": "#F2C94C",
    "Bhiwadi": "#F2994A",
    "Jhagadia": "#4A90E2",
}

CATEGORY_COLORS = {
    "COF": "#EAF4FF",
    "Working Capital": "#EEF9F0",
    "Responsible Purchase": "#FFF5E6",
    "RFX": "#F3EEFF",
    "E Auction": "#FFEFF3",
    "Prorating": "#EFFFF9",
    "Pro Contract": "#FFF4EC",
    "TOD": "#EEF4FF",
}


@dataclass(frozen=True)
class KPIRow:
    category: str
    section: str
    default_uom: str
    value_kind: str  # "number" or "percent"


KPI_ROWS = [
    KPIRow("COF", "COF Chennai", "MINR", "number"),
    KPIRow("COF", "COF Bhiwadi", "MINR", "number"),
    KPIRow("COF", "COF Jhagadia", "MINR", "number"),
    KPIRow("Working Capital", "Working Capital WGC", "MINR", "number"),
    KPIRow("Working Capital", "Working Capital Bhiwadi", "MINR", "number"),
    KPIRow("Working Capital", "Working Capital Jhagadia", "MINR", "number"),
    KPIRow("Responsible Purchase", "Responsible Purchase",
           "Assessed Spend %", "percent"),
    KPIRow("RFX", "RFX - Chennai", "Nos.", "number"),
    KPIRow("RFX", "RFX - Bhiwadi", "Nos.", "number"),
    KPIRow("RFX", "RFX - Jhagadia", "Nos.", "number"),
    KPIRow("E Auction", "E Auction - Chennai", "Nos.", "number"),
    KPIRow("E Auction", "Auction - Bhiwadi", "Nos.", "number"),
    KPIRow("E Auction", "E Auction - Jhagadia", "Nos.", "number"),
    KPIRow("Prorating", "Prorating - Chennai", "Nos.", "number"),
    KPIRow("Prorating", "Prorating - Bhiwadi", "Nos.", "number"),
    KPIRow("Prorating", "Prorating - Jhagadia", "Nos.", "number"),
    KPIRow("Pro Contract", "Pro Contract - Chennai", "Nos.", "number"),
    KPIRow("Pro Contract", "Pro Contract - Bhiwadi", "Nos.", "number"),
    KPIRow("Pro Contract", "Pro Rating - Jhagadia", "Nos.", "number"),
    KPIRow("TOD", "TOD - Chennai", "MINR", "number"),
    KPIRow("TOD", "TOD - Bhiwadi", "MINR", "number"),
    KPIRow("TOD", "TOD - Jhagadia", "MINR", "number"),
]


def _extract_location(section_name):
    s = section_name.lower()
    if "chennai" in s or "wgc" in s:
        return "Chennai"
    if "bhiwadi" in s:
        return "Bhiwadi"
    if "jhagadia" in s:
        return "Jhagadia"
    return "Other"


def _target_color(hex_color):
    color = hex_color.strip("#")
    r = int(color[0:2], 16)
    g = int(color[2:4], 16)
    b = int(color[4:6], 16)
    # Blend toward white for the target bar tone
    r = int(r + (255 - r) * 0.45)
    g = int(g + (255 - g) * 0.45)
    b = int(b + (255 - b) * 0.45)
    return f"#{r:02X}{g:02X}{b:02X}"


class PurchaseKPIDashboard(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Purchase KPI Dashboard")
        self.geometry("1280x790")
        self.minsize(1060, 680)
        self.configure(bg="#FFFFFF")

        self.row_fields = {}
        self.month_var = tk.StringVar(
            value=MONTH_OPTIONS[datetime.now().month - 1])
        self.year_var = tk.StringVar(value=str(datetime.now().year))

        self._build_ui()

    def _build_ui(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background="#FFFFFF")
        style.configure("TLabel", background="#FFFFFF", foreground="#1B1B1B")
        style.configure("Header.TLabel", background="#ECF2F8",
                        foreground="#1B1B1B", font=("Arial", 10, "bold"))
        style.configure("Title.TLabel", background="#FFFFFF",
                        foreground="#0E2F44", font=("Arial", 17, "bold"))
        style.configure("TButton", padding=(10, 7), font=("Arial", 10, "bold"))
        style.configure("TCombobox", fieldbackground="#FFFFFF",
                        background="#FFFFFF", foreground="#1B1B1B")

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        top_frame = ttk.Frame(self, padding=(14, 10))
        top_frame.grid(row=0, column=0, sticky="nsew")
        top_frame.grid_rowconfigure(2, weight=1)
        top_frame.grid_columnconfigure(0, weight=1)

        title = ttk.Label(
            top_frame, text="Purchase KPI Dashboard", style="Title.TLabel")
        title.grid(row=0, column=0, sticky="w", pady=(0, 10))

        period_row = ttk.Frame(top_frame)
        period_row.grid(row=1, column=0, sticky="w", pady=(0, 8))

        ttk.Label(period_row, text="Month:", font=(
            "Arial", 10, "bold")).grid(row=0, column=0, padx=(0, 6))
        month_combo = ttk.Combobox(
            period_row, textvariable=self.month_var, values=MONTH_OPTIONS, width=13, state="readonly")
        month_combo.grid(row=0, column=1, padx=(0, 16))

        current_year = datetime.now().year
        year_options = [str(y)
                        for y in range(current_year - 5, current_year + 6)]
        ttk.Label(period_row, text="Year:", font=("Arial", 10, "bold")).grid(
            row=0, column=2, padx=(0, 6))
        year_combo = ttk.Combobox(
            period_row, textvariable=self.year_var, values=year_options, width=8, state="readonly")
        year_combo.grid(row=0, column=3)

        canvas_holder = ttk.Frame(top_frame)
        canvas_holder.grid(row=2, column=0, sticky="nsew")
        canvas_holder.grid_rowconfigure(0, weight=1)
        canvas_holder.grid_columnconfigure(0, weight=1)

        canvas = tk.Canvas(canvas_holder, highlightthickness=0, bg="#FFFFFF")
        scrollbar = ttk.Scrollbar(
            canvas_holder, orient="vertical", command=canvas.yview)
        self.form_frame = tk.Frame(canvas, bg="#FFFFFF")

        self.form_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )

        canvas.create_window((0, 0), window=self.form_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(
            int(-1 * (event.delta / 120)), "units"))

        self._build_form_rows()

        button_frame = ttk.Frame(self, padding=(14, 8))
        button_frame.grid(row=1, column=0, sticky="ew")

        ttk.Button(button_frame, text="Report in PDF",
                   command=self.report_pdf).grid(row=0, column=0, padx=6)
        ttk.Button(button_frame, text="Report in HTML",
                   command=self.report_html).grid(row=0, column=1, padx=6)
        ttk.Button(button_frame, text="Reset", command=self.reset_fields).grid(
            row=0, column=2, padx=6)
        ttk.Button(button_frame, text="Exit", command=self.quit).grid(
            row=0, column=3, padx=6)

    def _build_form_rows(self):
        header_bg = "#ECF2F8"
        headers = ["Category", "Section", "UOM", "Actual", "Target"]

        for col, header in enumerate(headers):
            label = tk.Label(
                self.form_frame,
                text=header,
                bg=header_bg,
                fg="#1B1B1B",
                font=("Arial", 10, "bold"),
                bd=1,
                relief="solid",
                padx=8,
                pady=6,
                anchor="w",
            )
            label.grid(row=0, column=col, sticky="ew")

        for col in range(5):
            self.form_frame.grid_columnconfigure(
                col, weight=1 if col == 1 else 0)

        current_row = 1
        active_category = None

        for row in KPI_ROWS:
            if row.category != active_category:
                cat_color = CATEGORY_COLORS.get(row.category, "#F5F5F5")
                cat_label = tk.Label(
                    self.form_frame,
                    text=row.category,
                    bg=cat_color,
                    fg="#0F2D40",
                    font=("Arial", 11, "bold"),
                    bd=1,
                    relief="solid",
                    padx=10,
                    pady=6,
                    anchor="w",
                )
                cat_label.grid(row=current_row, column=0,
                               columnspan=5, sticky="ew", pady=(8, 2))
                current_row += 1
                active_category = row.category

            uom_var = tk.StringVar(value=row.default_uom)
            actual_var = tk.StringVar()
            target_var = tk.StringVar()

            tk.Label(self.form_frame, text="", bg="#FFFFFF").grid(
                row=current_row, column=0, padx=2)
            tk.Label(self.form_frame, text=row.section, bg="#FFFFFF", fg="#1B1B1B", font=("Arial", 10), anchor="w").grid(
                row=current_row, column=1, padx=8, pady=4, sticky="w"
            )

            uom_combo = ttk.Combobox(
                self.form_frame,
                textvariable=uom_var,
                values=UOM_OPTIONS,
                width=18,
                state="readonly",
            )
            uom_combo.grid(row=current_row, column=2,
                           padx=8, pady=4, sticky="w")

            actual_entry = tk.Entry(self.form_frame, textvariable=actual_var,
                                    width=20, bg="#FFFFFF", fg="#1B1B1B", relief="solid", bd=1)
            target_entry = tk.Entry(self.form_frame, textvariable=target_var,
                                    width=20, bg="#FFFFFF", fg="#1B1B1B", relief="solid", bd=1)
            actual_entry.grid(row=current_row, column=3,
                              padx=8, pady=4, sticky="w")
            target_entry.grid(row=current_row, column=4,
                              padx=8, pady=4, sticky="w")

            self.row_fields[row.section] = {
                "category": row.category,
                "value_kind": row.value_kind,
                "default_uom": row.default_uom,
                "uom_var": uom_var,
                "actual_var": actual_var,
                "target_var": target_var,
            }
            current_row += 1

        note = (
            "Input rules: Responsible Purchase should be 0 to 100 (percentage). "
            "All other Actual/Target values should be numeric."
        )
        tk.Label(self.form_frame, text=note, bg="#FFFFFF", fg="#666", font=("Arial", 9)).grid(
            row=current_row,
            column=0,
            columnspan=5,
            padx=8,
            pady=(12, 6),
            sticky="w",
        )

    def reset_fields(self):
        self.month_var.set(MONTH_OPTIONS[datetime.now().month - 1])
        self.year_var.set(str(datetime.now().year))
        for meta in self.row_fields.values():
            meta["uom_var"].set(meta["default_uom"])
            meta["actual_var"].set("")
            meta["target_var"].set("")
        messagebox.showinfo("Reset Complete", "All values have been reset.")

    def report_pdf(self):
        try:
            collected, month, year = self._collect_inputs()
            report_path = self._generate_pdf_report(collected, month, year)
            self._open_file(report_path)
            messagebox.showinfo("Report Generated",
                                f"PDF report created at:\n{report_path}")
        except ValueError as exc:
            messagebox.showerror("Validation Error", str(exc))
        except ImportError:
            messagebox.showerror(
                "Missing Dependency",
                "matplotlib is required for report generation.\nInstall it with: pip install matplotlib",
            )
        except Exception as exc:
            messagebox.showerror(
                "Report Error", f"Could not generate PDF report:\n{exc}")

    def report_html(self):
        try:
            collected, month, year = self._collect_inputs()
            report_path = self._generate_html_report(collected, month, year)
            self._open_file(report_path)
            messagebox.showinfo("Report Generated",
                                f"HTML report created at:\n{report_path}")
        except ValueError as exc:
            messagebox.showerror("Validation Error", str(exc))
        except Exception as exc:
            messagebox.showerror(
                "Report Error", f"Could not generate HTML report:\n{exc}")

    def _collect_inputs(self):
        month = self.month_var.get().strip()
        year = self.year_var.get().strip()

        if not month:
            raise ValueError("Please select Month.")
        if not year:
            raise ValueError("Please select Year.")

        collected = {}
        errors = []

        for section, meta in self.row_fields.items():
            category = meta["category"]
            uom = meta["uom_var"].get().strip()
            actual_text = meta["actual_var"].get().strip().replace(",", "")
            target_text = meta["target_var"].get().strip().replace(",", "")

            if not actual_text or not target_text:
                errors.append(f"{section}: Actual and Target are required.")
                continue

            try:
                actual = float(actual_text)
                target = float(target_text)
            except ValueError:
                errors.append(f"{section}: Actual and Target must be numeric.")
                continue

            if meta["value_kind"] == "percent":
                if not (0 <= actual <= 100 and 0 <= target <= 100):
                    errors.append(
                        f"{section}: percentage values must be between 0 and 100.")

            if category not in collected:
                collected[category] = []

            collected[category].append(
                {
                    "section": section,
                    "uom": uom,
                    "actual": actual,
                    "target": target,
                    "value_kind": meta["value_kind"],
                    "location": _extract_location(section),
                }
            )

        if errors:
            raise ValueError("\n".join(errors))

        return collected, month, year

    @staticmethod
    def _format_value(value, value_kind):
        if value_kind == "percent":
            return f"{value:.2f}%"
        return f"{value:,.2f}"

    @staticmethod
    def _draw_speedometer(ax, actual_value):
        from matplotlib.patches import Circle, Wedge

        ax.set_aspect("equal")
        ax.axis("off")
        ax.set_xlim(-1.30, 1.30)
        ax.set_ylim(-0.36, 1.28)

        segments = [
            (0, 40, "#D64C4C"),
            (40, 75, "#E0B93A"),
            (75, 100, "#4CAF50"),
        ]
        for start, end, color in segments:
            theta1 = 180 * (1 - end / 100)
            theta2 = 180 * (1 - start / 100)
            wedge = Wedge((0, 0), 1.18, theta1, theta2, width=0.38,
                          facecolor=color, edgecolor="white")
            ax.add_patch(wedge)

        actual_value = max(0, min(100, actual_value))
        angle = math.radians(180 * (1 - actual_value / 100))
        x = 0.88 * math.cos(angle)
        y = 0.88 * math.sin(angle)
        ax.plot([0, x], [0, y], color="#1A1A1A", linewidth=2.5)
        ax.add_patch(Circle((0, 0), 0.055, color="#1A1A1A"))

        for tick in [0, 20, 40, 60, 75, 100]:
            tick_angle = math.radians(180 * (1 - tick / 100))
            tx = 1.30 * math.cos(tick_angle)
            ty = 1.30 * math.sin(tick_angle)
            ax.text(tx, ty, f"{tick}", ha="center", va="center", fontsize=9)

        ax.text(0, -0.16, f"Actual: {actual_value:.1f}%",
                ha="center", va="center", fontsize=10, fontweight="bold")
        ax.set_title("Responsible Purchase", fontsize=11.5,
                     fontweight="bold", pad=9)

    def _generate_pdf_report(self, data, month, year):
        import numpy as np
        import matplotlib.pyplot as plt
        from matplotlib.backends.backend_pdf import PdfPages
        from matplotlib.lines import Line2D
        from matplotlib.patches import Patch, Rectangle
        from matplotlib.ticker import FuncFormatter, MaxNLocator

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Purchase_KPI_Dashboard_Report_{timestamp}.pdf"
        output_path = os.path.abspath(filename)
        total_pages = 8

        cof_items = data["COF"]
        wc_items = data["Working Capital"]
        rp_item = data["Responsible Purchase"][0]
        rfx_items = data["RFX"]
        ea_items = data["E Auction"]
        prorating_items = data["Prorating"]
        pro_contract_items = data["Pro Contract"]
        tod_items = data["TOD"]

        def compact_tick(value, _pos=None):
            abs_value = abs(value)
            if abs_value >= 1_000_000_000:
                return f"{value / 1_000_000_000:.1f}B"
            if abs_value >= 1_000_000:
                return f"{value / 1_000_000:.1f}M"
            if abs_value >= 1_000:
                return f"{value / 1_000:.1f}K"
            if abs_value >= 100:
                return f"{value:.0f}"
            if abs_value >= 10:
                return f"{value:.1f}"
            return f"{value:.2f}"

        def add_page_border(fig):
            fig.patches.append(
                Rectangle(
                    (0.018, 0.018),
                    0.964,
                    0.964,
                    transform=fig.transFigure,
                    fill=False,
                    linewidth=1.6,
                    edgecolor="#8A9AA8",
                    zorder=1000,
                )
            )

        def add_page_header(fig, title, subtitle=None):
            title_text = fig.text(
                0.5,
                0.955,
                title,
                ha="center",
                va="center",
                fontsize=20,
                fontweight="bold",
                color="#11334A",
            )
            if subtitle:
                fig.text(
                    0.5,
                    0.922,
                    subtitle,
                    ha="center",
                    va="center",
                    fontsize=12.5,
                    fontweight="bold",
                    color="#1F4E6B",
                )
            fig.canvas.draw()
            title_bbox = title_text.get_window_extent(
                renderer=fig.canvas.get_renderer()).transformed(fig.transFigure.inverted())
            underline_y = title_bbox.y0 - 0.005
            fig.lines.append(
                Line2D(
                    [title_bbox.x0, title_bbox.x1],
                    [underline_y, underline_y],
                    transform=fig.transFigure,
                    color="#11334A",
                    linewidth=1.3,
                )
            )

        def add_page_footer(fig, page_number):
            footer_text = f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            fig.text(0.03, 0.026, footer_text, fontsize=8.4, color="#4F5B66")
            fig.text(0.97, 0.026, f"Page {page_number}/{total_pages}",
                     ha="right", fontsize=8.6, color="#4F5B66")

        def draw_summary_table(
            ax,
            title,
            items,
            rename_wgc=False,
            col_widths=None,
            bbox=None,
            font_size=8.4,
            row_scale=1.08,
            center_all=False,
            title_loc="left",
        ):
            ax.axis("off")
            ax.set_title(title, fontsize=11.2,
                         fontweight="bold", loc=title_loc, pad=8)

            rows = []
            for item in items:
                section = item["section"].replace(
                    "WGC", "Chennai") if rename_wgc else item["section"]
                rows.append(
                    [
                        section,
                        item["uom"],
                        f"{item['actual']:,.2f}",
                        f"{item['target']:,.2f}",
                    ]
                )

            if col_widths is None:
                col_widths = [0.56, 0.12, 0.16, 0.16]
            if bbox is None:
                bbox = [0.0, 0.03, 1.0, 0.90]

            table = ax.table(
                cellText=rows,
                colLabels=["Section", "UOM", "Actual", "Target"],
                loc="center",
                colLoc="center",
                cellLoc="left",
                colWidths=col_widths,
                bbox=bbox,
            )
            table.auto_set_font_size(False)
            table.set_fontsize(font_size)
            table.scale(1.0, row_scale)

            for (r, c), cell in table.get_celld().items():
                cell.set_edgecolor("#C4D0DB")
                cell.set_linewidth(1.05 if r != 0 else 1.20)
                cell.PAD = 0.085
                if r == 0:
                    cell.set_facecolor("#EFF4F8")
                    cell.set_text_props(
                        weight="bold", color="#1F2F3D", va="center")
                else:
                    cell.set_facecolor("white")
                    cell.set_text_props(va="center")
                    if c in (2, 3):
                        cell.get_text().set_fontweight("bold")
                if center_all or c in (1, 2, 3):
                    cell.get_text().set_ha("center")
                else:
                    cell.get_text().set_ha("left")
                cell.get_text().set_wrap(True)

        def draw_bar_chart(ax, items, title, legend_y=-0.22, title_size=10.8):
            locations = ["Chennai", "Bhiwadi", "Jhagadia"]
            by_loc = {item["location"]: item for item in items}
            actual = [by_loc.get(loc, {}).get("actual", 0.0)
                      for loc in locations]
            target = [by_loc.get(loc, {}).get("target", 0.0)
                      for loc in locations]

            x = np.arange(len(locations))
            width = 0.38
            bars_actual = ax.bar(x - width / 2, actual, width,
                                 color="#F2C94C", edgecolor="#A08D2E", linewidth=0.6)
            bars_target = ax.bar(x + width / 2, target, width,
                                 color="#2E8B57", edgecolor="#216A41", linewidth=0.6)

            ax.set_xticks(x)
            ax.set_xticklabels(locations, fontsize=9.5, fontweight="bold")
            ax.set_title(title, fontsize=title_size, fontweight="bold")
            ax.grid(axis="y", linestyle="--", alpha=0.25)
            ax.tick_params(axis="y", labelsize=9)

            max_val = max(1.0, max(actual + target))
            ax.set_ylim(0, max_val * 1.18)
            ax.yaxis.set_major_locator(MaxNLocator(nbins=5))
            ax.yaxis.set_major_formatter(FuncFormatter(compact_tick))
            ax.margins(x=0.12)

            for bar, value in zip(bars_actual, actual):
                ax.text(
                    bar.get_x() + bar.get_width() / 2,
                    bar.get_height() + (max_val * 0.03),
                    f"{value:.1f}",
                    ha="center",
                    va="bottom",
                    fontsize=8.3,
                    fontweight="bold",
                    color="#2B3B4B",
                )
            for bar, value in zip(bars_target, target):
                ax.text(
                    bar.get_x() + bar.get_width() / 2,
                    bar.get_height() + (max_val * 0.03),
                    f"{value:.1f}",
                    ha="center",
                    va="bottom",
                    fontsize=8.3,
                    fontweight="bold",
                    color="#2B3B4B",
                )

            actual_patch = Patch(facecolor="#F2C94C", edgecolor="#A08D2E")
            target_patch = Patch(facecolor="#2E8B57", edgecolor="#216A41")
            ax.legend(
                [actual_patch, target_patch],
                ["Actual", "Target"],
                fontsize=8.7,
                loc="upper center",
                bbox_to_anchor=(0.5, legend_y),
                ncol=2,
                frameon=False,
            )

        def draw_tod_pie(
            ax,
            items,
            title="TOD Target (Pie Chart)",
            radius=1.15,
            labeldistance=1.18,
            label_font=9.2,
            show_legend=False,
            inside_labels=False,
            inside_label_font=9.6,
        ):
            locations = ["Chennai", "Bhiwadi", "Jhagadia"]
            by_loc = {item["location"]: item for item in items}
            values = [float(by_loc.get(loc, {}).get("target", 0.0))
                      for loc in locations]
            colors = [LOCATION_COLORS[loc] for loc in locations]

            ax.set_title(title, fontsize=11.5, fontweight="bold", pad=8)
            total = sum(values)
            if total <= 0:
                ax.axis("off")
                ax.text(0.5, 0.5, "No target values to display",
                        ha="center", va="center", fontsize=10, color="#667788")
                return locations, values, total

            if inside_labels:
                wedges, _ = ax.pie(
                    values,
                    colors=colors,
                    startangle=90,
                    radius=radius,
                    labels=None,
                    wedgeprops={"linewidth": 1.0, "edgecolor": "white"},
                )
                for wedge, loc, value, color in zip(wedges, locations, values, colors):
                    pct = (value / total) * 100 if total else 0
                    angle = math.radians((wedge.theta1 + wedge.theta2) / 2.0)
                    label_radius = radius * 0.62
                    x = label_radius * math.cos(angle)
                    y = label_radius * math.sin(angle)
                    rgb = color.lstrip("#")
                    r = int(rgb[0:2], 16)
                    g = int(rgb[2:4], 16)
                    b = int(rgb[4:6], 16)
                    luminance = (0.299 * r + 0.587 * g + 0.114 * b)
                    text_color = "#14212E" if luminance > 165 else "white"
                    ax.text(
                        x,
                        y,
                        f"{loc}\n{pct:.1f}%\n{value:,.1f}",
                        ha="center",
                        va="center",
                        fontsize=inside_label_font,
                        fontweight="bold",
                        color=text_color,
                        linespacing=1.15,
                    )
            else:
                labels = []
                for loc, value in zip(locations, values):
                    pct = (value / total) * 100 if total else 0
                    labels.append(f"{loc}\n{value:.1f} ({pct:.1f}%)")

                wedges, _ = ax.pie(
                    values,
                    colors=colors,
                    startangle=90,
                    radius=radius,
                    labels=labels,
                    labeldistance=labeldistance,
                    textprops={"fontsize": label_font, "color": "#203040"},
                    wedgeprops={"linewidth": 1.0, "edgecolor": "white"},
                )
            ax.set_aspect("equal")
            if show_legend:
                legend_labels = []
                for loc, value in zip(locations, values):
                    pct = (value / total) * 100 if total else 0
                    legend_labels.append(f"{loc}: {value:.1f} ({pct:.1f}%)")
                ax.legend(
                    wedges,
                    legend_labels,
                    title="Locations",
                    fontsize=10,
                    title_fontsize=10.5,
                    loc="center left",
                    bbox_to_anchor=(1.02, 0.5),
                    frameon=False,
                )
            return locations, values, total

        def render_large_table_page(pdf, title, items, page_number, rename_wgc=False):
            fig = plt.figure(figsize=(16.5, 11), facecolor="white")
            add_page_header(fig, title, f"Month: {month} | Year: {year}")
            ax = fig.add_axes([0.06, 0.12, 0.88, 0.76])
            draw_summary_table(
                ax,
                title.replace("Big ", ""),
                items,
                rename_wgc=rename_wgc,
                col_widths=[0.50, 0.14, 0.18, 0.18],
                bbox=[0.0, 0.01, 1.0, 0.94],
                font_size=12.0,
                row_scale=1.58,
                center_all=True,
                title_loc="center",
            )
            add_page_footer(fig, page_number)
            add_page_border(fig)
            pdf.savefig(fig, bbox_inches="tight")
            plt.close(fig)

        def render_bar_page(pdf, page_title, items, page_number):
            fig = plt.figure(figsize=(16.5, 11), facecolor="white")
            add_page_header(fig, page_title, f"Month: {month} | Year: {year}")
            grid = fig.add_gridspec(
                2,
                1,
                height_ratios=[2.35, 1.25],
                left=0.07,
                right=0.93,
                top=0.84,
                bottom=0.10,
                hspace=0.52,
            )
            ax_chart = fig.add_subplot(grid[0, 0])
            draw_bar_chart(
                ax_chart, items, f"{page_title} (Actual vs Target)", legend_y=-0.05, title_size=13)

            ax_table = fig.add_subplot(grid[1, 0])
            draw_summary_table(
                ax_table,
                f"{page_title} Values",
                items,
                col_widths=[0.54, 0.12, 0.17, 0.17],
                bbox=[0.0, 0.02, 1.0, 0.92],
                font_size=10.5,
                row_scale=1.35,
                center_all=True,
                title_loc="center",
            )
            add_page_footer(fig, page_number)
            add_page_border(fig)
            pdf.savefig(fig, bbox_inches="tight")
            plt.close(fig)

        with PdfPages(output_path) as pdf:
            fig = plt.figure(figsize=(16.5, 11), facecolor="white")
            add_page_header(fig, "Purchase KPI Dashboard",
                            f"Month: {month} | Year: {year}")

            outer = fig.add_gridspec(
                2,
                1,
                height_ratios=[0.90, 3.10],
                left=0.045,
                right=0.955,
                top=0.855,
                bottom=0.085,
                hspace=0.22,
            )
            top = outer[0].subgridspec(1, 2, wspace=0.14)
            bottom = outer[1].subgridspec(
                3,
                2,
                height_ratios=[2.18, 1.02, 1.02],
                width_ratios=[1.08, 1.30],
                hspace=0.72,
                wspace=0.26,
            )

            overall_cof = sum(item["actual"] for item in cof_items)
            cof_uom = cof_items[0]["uom"] if cof_items else ""
            ax_cof = fig.add_subplot(top[0, 0])
            draw_summary_table(
                ax_cof,
                f"COF (Overall Actual: {overall_cof:,.2f} {cof_uom})",
                cof_items,
                col_widths=[0.56, 0.12, 0.16, 0.16],
                bbox=[0.0, 0.03, 1.0, 0.90],
                font_size=8.5,
                row_scale=1.03,
                center_all=False,
                title_loc="left",
            )

            ax_wc = fig.add_subplot(top[0, 1])
            draw_summary_table(
                ax_wc,
                "Working Capital",
                wc_items,
                rename_wgc=True,
                col_widths=[0.56, 0.12, 0.16, 0.16],
                bbox=[0.0, 0.03, 1.0, 0.90],
                font_size=8.5,
                row_scale=1.03,
                center_all=False,
                title_loc="left",
            )

            ax_resp = fig.add_subplot(bottom[0, 0])
            PurchaseKPIDashboard._draw_speedometer(ax_resp, rp_item["actual"])

            ax_tod = fig.add_subplot(bottom[0, 1])
            draw_tod_pie(
                ax_tod,
                tod_items,
                title="TOD Target",
                radius=1.34,
                labeldistance=1.35,
                label_font=10.2,
                show_legend=False,
            )

            ax_rfx = fig.add_subplot(bottom[1, 0])
            draw_bar_chart(ax_rfx, rfx_items, "RFX",
                           legend_y=-0.28, title_size=11)

            ax_ea = fig.add_subplot(bottom[1, 1])
            draw_bar_chart(ax_ea, ea_items, "E Auction",
                           legend_y=-0.28, title_size=11)

            ax_pc = fig.add_subplot(bottom[2, 0])
            draw_bar_chart(ax_pc, pro_contract_items,
                           "Pro Contract", legend_y=-0.28, title_size=11)

            ax_pr = fig.add_subplot(bottom[2, 1])
            draw_bar_chart(ax_pr, prorating_items, "Prorating",
                           legend_y=-0.28, title_size=11)

            add_page_footer(fig, 1)
            add_page_border(fig)
            pdf.savefig(fig, bbox_inches="tight")
            plt.close(fig)

            render_large_table_page(
                pdf, "Big COF Tabular", cof_items, page_number=2, rename_wgc=False)
            render_large_table_page(
                pdf, "Big Working Capital Tabular", wc_items, page_number=3, rename_wgc=True)

            fig = plt.figure(figsize=(16.5, 11), facecolor="white")
            add_page_header(fig, "TOD Target Pie Chart",
                            f"Month: {month} | Year: {year}")
            ax_pie = fig.add_axes([0.08, 0.26, 0.66, 0.52])
            locations, values, total = draw_tod_pie(
                ax_pie,
                tod_items,
                title="TOD Target (Large View)",
                radius=1.18,
                labeldistance=1.20,
                label_font=10.5,
                show_legend=False,
                inside_labels=True,
                inside_label_font=10.0,
            )

            pie_data_rows = []
            for loc, value in zip(locations, values):
                pct = (value / total) * 100 if total else 0.0
                pie_data_rows.append([loc, f"{value:,.2f}", f"{pct:.1f}%"])

            ax_pie_table = fig.add_axes([0.14, 0.10, 0.72, 0.12])
            ax_pie_table.axis("off")
            pie_table = ax_pie_table.table(
                cellText=pie_data_rows,
                colLabels=["Location", "Target", "Share"],
                loc="center",
                colLoc="center",
                cellLoc="center",
                colWidths=[0.45, 0.27, 0.28],
                bbox=[0.0, 0.0, 1.0, 1.0],
            )
            pie_table.auto_set_font_size(False)
            pie_table.set_fontsize(11)
            pie_table.scale(1.0, 1.45)
            for (r, c), cell in pie_table.get_celld().items():
                cell.set_edgecolor("#C4D0DB")
                cell.set_linewidth(1.05 if r != 0 else 1.20)
                if r == 0:
                    cell.set_facecolor("#EFF4F8")
                    cell.set_text_props(weight="bold", color="#1F2F3D")
                else:
                    cell.set_facecolor("white")
                    if c in (1, 2):
                        cell.get_text().set_fontweight("bold")

            add_page_footer(fig, 4)
            add_page_border(fig)
            pdf.savefig(fig, bbox_inches="tight")
            plt.close(fig)

            render_bar_page(pdf, "Pro Contract",
                            pro_contract_items, page_number=5)
            render_bar_page(pdf, "Prorating", prorating_items, page_number=6)
            render_bar_page(pdf, "E Auction", ea_items, page_number=7)
            render_bar_page(pdf, "RFX", rfx_items, page_number=8)

        return output_path

    def _generate_html_report(self, data, month, year):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Purchase_KPI_Dashboard_Report_{timestamp}.html"
        output_path = os.path.abspath(filename)

        def table_rows(items, rename_wgc=False):
            rows = []
            for item in items:
                loc = item["location"]
                dot_color = LOCATION_COLORS.get(loc, "#777")
                label = item["section"].replace(
                    "WGC", "Chennai") if rename_wgc else item["section"]
                rows.append(
                    f"<tr><td><span class='dot' style='background:{dot_color}'></span>{label}</td>"
                    f"<td>{item['uom']}</td><td>{item['actual']:.2f}</td><td>{item['target']:.2f}</td></tr>"
                )
            return "\n".join(rows)

        def category_payload(items):
            payload = {}
            for loc in ["Chennai", "Bhiwadi", "Jhagadia"]:
                hit = next(
                    (item for item in items if item["location"] == loc), None)
                payload[loc] = {
                    "actual": (hit["actual"] if hit else 0.0),
                    "target": (hit["target"] if hit else 0.0),
                }
            return payload

        cof_overall = sum(item["actual"] for item in data["COF"])
        rp_actual = data["Responsible Purchase"][0]["actual"]

        rfx_payload = category_payload(data["RFX"])
        e_auction_payload = category_payload(data["E Auction"])
        prorating_payload = category_payload(data["Prorating"])
        pro_contract_payload = category_payload(data["Pro Contract"])
        tod_target_payload = {
            loc: values["target"] for loc, values in category_payload(data["TOD"]).items()
        }

        location_colors_json = json.dumps(LOCATION_COLORS)
        rfx_json = json.dumps(rfx_payload)
        e_auction_json = json.dumps(e_auction_payload)
        prorating_json = json.dumps(prorating_payload)
        pro_contract_json = json.dumps(pro_contract_payload)
        tod_target_json = json.dumps(tod_target_payload)

        html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Purchase KPI Dashboard Report</title>
<style>
body {{ font-family: Arial, sans-serif; margin: 24px; background: #ffffff; color: #1f1f1f; }}
h1 {{ text-align: center; margin: 0; color: #11334A; text-decoration: underline; }}
h2 {{ margin: 8px 0 14px 0; text-align: center; color: #1F4E6B; }}
.summary-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 16px; }}
.viz-grid {{
  display: grid;
  grid-template-columns: 1fr 1fr;
  grid-template-areas:
    "resp tod"
    "rfx ea"
    "pc  pr";
  gap: 16px;
}}
.card {{ border: 1px solid #B8C3CD; border-radius: 8px; padding: 12px; background: #FFFFFF; }}
.card h3 {{ margin-top: 0; color: #0E2F44; }}
table {{ width: 100%; border-collapse: collapse; }}
th, td {{ border: 1px solid #D6DEE6; padding: 7px; font-size: 13px; }}
th {{ background: #EFF4F8; }}
.summary-table th:nth-child(2), .summary-table th:nth-child(3), .summary-table th:nth-child(4),
.summary-table td:nth-child(2), .summary-table td:nth-child(3), .summary-table td:nth-child(4) {{ text-align: center; }}
.dot {{ width: 10px; height: 10px; display: inline-block; border-radius: 50%; margin-right: 6px; }}
.chart {{ width: 100%; height: 260px; display: block; }}
.chart-tall {{ height: 300px; }}
.chart-pie {{ height: 380px; }}
.speedo {{ width: 100%; height: 280px; display: block; }}
.small {{ color: #607080; font-size: 12px; margin: 6px 0 0 0; }}
.speed-wrap {{ text-align: center; }}
.bar-legend {{ display: flex; gap: 10px; flex-wrap: wrap; margin-top: 8px; }}
.legend-chip {{ font-size: 12px; border: 1px solid #B8C3CD; border-radius: 12px; padding: 4px 10px; }}
.swatch {{ display: inline-block; width: 12px; height: 12px; border: 1px solid #888; margin-right: 6px; vertical-align: -1px; }}
.resp {{ grid-area: resp; }}
.rfx {{ grid-area: rfx; }}
.ea {{ grid-area: ea; }}
.pr {{ grid-area: pr; }}
.tod {{ grid-area: tod; }}
.pc {{ grid-area: pc; }}
</style>
</head>
<body>
  <h1>Purchase KPI Dashboard</h1>
  <h2>Month: {month} | Year: {year}</h2>

  <div class="summary-grid">
    <div class="card">
      <h3>COF (Overall Actual: {cof_overall:,.2f})</h3>
      <table class="summary-table">
        <tr><th>Section</th><th>UOM</th><th>Actual</th><th>Target</th></tr>
        {table_rows(data['COF'])}
      </table>
    </div>

    <div class="card">
      <h3>Working Capital</h3>
      <table class="summary-table">
        <tr><th>Section</th><th>UOM</th><th>Actual</th><th>Target</th></tr>
        {table_rows(data['Working Capital'], rename_wgc=True)}
      </table>
    </div>
  </div>

  <div class="viz-grid">
    <div class="card resp">
      <h3>Responsible Purchase</h3>
      <div class="speed-wrap">
        <canvas id="speedometer" class="speedo"></canvas>
        <div class="small">Actual Value from Input: <b>{rp_actual:.2f}%</b></div>
      </div>
    </div>

    <div class="card rfx">
      <h3>RFX</h3>
      <canvas id="rfxChart" class="chart chart-tall"></canvas>
      <div class="bar-legend">
        <span class="legend-chip"><span class="swatch" style="background:#F2C94C;"></span>Actual</span>
        <span class="legend-chip"><span class="swatch" style="background:#2E8B57;"></span>Target</span>
      </div>
    </div>

    <div class="card ea">
      <h3>E Auction</h3>
      <canvas id="eAuctionChart" class="chart"></canvas>
      <div class="bar-legend">
        <span class="legend-chip"><span class="swatch" style="background:#F2C94C;"></span>Actual</span>
        <span class="legend-chip"><span class="swatch" style="background:#2E8B57;"></span>Target</span>
      </div>
    </div>

    <div class="card pr">
      <h3>Prorating</h3>
      <canvas id="proratingChart" class="chart"></canvas>
      <div class="bar-legend">
        <span class="legend-chip"><span class="swatch" style="background:#F2C94C;"></span>Actual</span>
        <span class="legend-chip"><span class="swatch" style="background:#2E8B57;"></span>Target</span>
      </div>
    </div>

    <div class="card tod">
      <h3>TOD Target (Pie Chart)</h3>
      <canvas id="todPie" class="chart chart-pie"></canvas>
      <div class="bar-legend">
        <span class="legend-chip"><span class="swatch" style="background:#F2C94C;"></span>Chennai</span>
        <span class="legend-chip"><span class="swatch" style="background:#F2994A;"></span>Bhiwadi</span>
        <span class="legend-chip"><span class="swatch" style="background:#4A90E2;"></span>Jhagadia</span>
      </div>
    </div>

    <div class="card pc">
      <h3>Pro Contract</h3>
      <canvas id="proContractChart" class="chart"></canvas>
      <div class="bar-legend">
        <span class="legend-chip"><span class="swatch" style="background:#F2C94C;"></span>Actual</span>
        <span class="legend-chip"><span class="swatch" style="background:#2E8B57;"></span>Target</span>
      </div>
    </div>
  </div>

  <p class="small">Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
<script>
const LOCATION_COLORS = {location_colors_json};
const ACTUAL_YELLOW = "#F2C94C";
const TARGET_GREEN = "#2E8B57";
const RFX_DATA = {rfx_json};
const E_AUCTION_DATA = {e_auction_json};
const PRORATING_DATA = {prorating_json};
const PRO_CONTRACT_DATA = {pro_contract_json};
const TOD_TARGET_DATA = {tod_target_json};
const RESPONSIBLE_ACTUAL = {rp_actual};

function setupCanvas(canvas) {{
  const dpr = window.devicePixelRatio || 1;
  const rect = canvas.getBoundingClientRect();
  canvas.width = Math.max(1, Math.floor(rect.width * dpr));
  canvas.height = Math.max(1, Math.floor(rect.height * dpr));
  const ctx = canvas.getContext("2d");
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  return [ctx, rect.width, rect.height];
}}

function drawBarChart(canvasId, categoryData) {{
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const [ctx, w, h] = setupCanvas(canvas);
  ctx.clearRect(0, 0, w, h);

  const locations = ["Chennai", "Bhiwadi", "Jhagadia"];
  const points = locations.map(loc => categoryData[loc] || {{actual: 0, target: 0}});
  const maxVal = Math.max(1, ...points.map(p => Math.max(p.actual, p.target)));

  const left = 52, right = 14, top = 16, bottom = 42;
  const plotW = w - left - right;
  const plotH = h - top - bottom;

  ctx.strokeStyle = "#D1DAE4";
  ctx.lineWidth = 1;
  for (let i = 0; i <= 4; i++) {{
    const y = top + (plotH * i / 4);
    ctx.beginPath();
    ctx.moveTo(left, y);
    ctx.lineTo(left + plotW, y);
    ctx.stroke();
    const tickVal = (maxVal * (1 - i / 4)).toFixed(0);
    ctx.fillStyle = "#607080";
    ctx.font = "11px Arial";
    ctx.fillText(tickVal, 8, y + 4);
  }}

  ctx.strokeStyle = "#7A8A98";
  ctx.beginPath();
  ctx.moveTo(left, top);
  ctx.lineTo(left, top + plotH);
  ctx.lineTo(left + plotW, top + plotH);
  ctx.stroke();

  const groupW = plotW / 3;
  const barW = Math.min(28, groupW * 0.24);
  const innerGap = 8;

  locations.forEach((loc, i) => {{
    const item = points[i];
    const gx = left + groupW * i + (groupW / 2);
    const actualH = (item.actual / maxVal) * plotH;
    const targetH = (item.target / maxVal) * plotH;

    const ax = gx - barW - innerGap / 2;
    const ay = top + plotH - actualH;
    ctx.fillStyle = ACTUAL_YELLOW;
    ctx.fillRect(ax, ay, barW, actualH);

    const tx = gx + innerGap / 2;
    const ty = top + plotH - targetH;
    ctx.fillStyle = TARGET_GREEN;
    ctx.fillRect(tx, ty, barW, targetH);

    ctx.fillStyle = "#253748";
    ctx.font = "11px Arial";
    ctx.textAlign = "center";
    ctx.fillText(item.actual.toFixed(1), ax + barW / 2, ay - 4);
    ctx.fillText(item.target.toFixed(1), tx + barW / 2, ty - 4);

    ctx.fillStyle = "#1F2F3D";
    ctx.font = "12px Arial";
    ctx.textAlign = "center";
    ctx.fillText(loc, gx, top + plotH + 18);
  }});
}}

function drawTODPie(canvasId, targetData) {{
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const [ctx, w, h] = setupCanvas(canvas);
  ctx.clearRect(0, 0, w, h);

  const locations = ["Chennai", "Bhiwadi", "Jhagadia"];
  const values = locations.map(loc => Number(targetData[loc] || 0));
  const total = values.reduce((a, b) => a + b, 0);
  const cx = w * 0.5;
  const cy = h * 0.52;
  const r = Math.min(w, h) * 0.43;

  if (total <= 0) {{
    ctx.fillStyle = "#6B7C8A";
    ctx.font = "14px Arial";
    ctx.textAlign = "center";
    ctx.fillText("No target values to display", w / 2, h / 2);
    return;
  }}

  let start = -Math.PI / 2;
  locations.forEach((loc, i) => {{
    const val = values[i];
    const angle = (val / total) * Math.PI * 2;
    const end = start + angle;
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, r, start, end);
    ctx.closePath();
    ctx.fillStyle = LOCATION_COLORS[loc];
    ctx.fill();
    const mid = (start + end) / 2;
    const lx = cx + (r * 0.68) * Math.cos(mid);
    const ly = cy + (r * 0.68) * Math.sin(mid);
    const pct = ((val / total) * 100).toFixed(1);
    ctx.fillStyle = "#172B3A";
    ctx.font = "bold 12px Arial";
    ctx.textAlign = "center";
    ctx.fillText(`${{loc}}`, lx, ly - 7);
    ctx.font = "11px Arial";
    ctx.fillText(`${{val.toFixed(1)}} (${{pct}}%)`, lx, ly + 10);
    start = end;
  }});
}}

function drawSpeedometer(canvasId, value) {{
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const [ctx, w, h] = setupCanvas(canvas);
  ctx.clearRect(0, 0, w, h);

  const cx = w / 2;
  const cy = h * 0.80;
  const radius = Math.min(w, h) * 0.39;

  function drawArc(p1, p2, color) {{
    const a1 = Math.PI + (p1 / 100) * Math.PI;
    const a2 = Math.PI + (p2 / 100) * Math.PI;
    ctx.beginPath();
    ctx.arc(cx, cy, radius, a1, a2);
    ctx.strokeStyle = color;
    ctx.lineWidth = 24;
    ctx.lineCap = "round";
    ctx.stroke();
  }}

  drawArc(0, 40, "#D64C4C");
  drawArc(40, 75, "#E0B93A");
  drawArc(75, 100, "#4CAF50");

  const v = Math.max(0, Math.min(100, Number(value || 0)));
  const angle = Math.PI + (v / 100) * Math.PI;
  const nx = cx + (radius - 10) * Math.cos(angle);
  const ny = cy + (radius - 10) * Math.sin(angle);
  ctx.beginPath();
  ctx.moveTo(cx, cy);
  ctx.lineTo(nx, ny);
  ctx.strokeStyle = "#101820";
  ctx.lineWidth = 4;
  ctx.stroke();
  ctx.beginPath();
  ctx.arc(cx, cy, 6, 0, Math.PI * 2);
  ctx.fillStyle = "#101820";
  ctx.fill();

  [0, 20, 40, 60, 75, 100].forEach(tick => {{
    const tAngle = Math.PI + (tick / 100) * Math.PI;
    const tx = cx + (radius + 24) * Math.cos(tAngle);
    const ty = cy + (radius + 24) * Math.sin(tAngle);
    ctx.fillStyle = "#33495C";
    ctx.font = "11px Arial";
    ctx.textAlign = "center";
    ctx.fillText(String(tick), tx, ty + 4);
  }});

  ctx.fillStyle = "#12344B";
  ctx.font = "bold 14px Arial";
  ctx.textAlign = "center";
  ctx.fillText(`Actual: ${{v.toFixed(2)}}%`, cx, h - 16);
}}

function renderAllCharts() {{
  drawSpeedometer("speedometer", RESPONSIBLE_ACTUAL);
  drawBarChart("rfxChart", RFX_DATA);
  drawBarChart("eAuctionChart", E_AUCTION_DATA);
  drawBarChart("proratingChart", PRORATING_DATA);
  drawBarChart("proContractChart", PRO_CONTRACT_DATA);
  drawTODPie("todPie", TOD_TARGET_DATA);
}}

window.addEventListener("load", renderAllCharts);
window.addEventListener("resize", renderAllCharts);
</script>
</body>
</html>"""

        with open(output_path, "w", encoding="utf-8") as file:
            file.write(html)

        return output_path

    @staticmethod
    def _open_file(path):
        try:
            system = platform.system().lower()
            if system == "windows":
                os.startfile(path)  # type: ignore[attr-defined]
            elif system == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception:
            pass


if __name__ == "__main__":
    app = PurchaseKPIDashboard()
    app.mainloop()
