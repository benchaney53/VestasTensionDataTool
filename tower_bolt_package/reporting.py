# -*- coding: utf-8 -*-
"""
Created on Wed Jan 10 11:10:56 2024

@author: TOBHI
"""
from datetime import datetime as dt
import os
import shutil
import warnings

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages


def color_code(series: pd.Series) -> pd.Series:
    """
    Map Pass/Fail/Alert text to fill colors.
    """
    colors = pd.Series(data="w", index=series.index)
    colors.loc[series.str.contains("Pass", na=False)] = "#d2f090"
    colors.loc[series.str.contains("Fail", na=False)] = "#f59393"
    colors.loc[series.str.contains("Alert", na=False)] = "#ffe8b3"
    return colors


def findReportTemplate(path: str) -> str:
    """
    Find a single Excel template file in path. Name must contain 'template' and end with .xlsx.
    """
    template_name = None
    try:
        dir_list = os.listdir(path)
    except Exception as e:
        raise Exception(f"Unable to list directory: {path}\n{e}")

    for name in dir_list:
        low = name.lower()
        if "template" in low and low.endswith(".xlsx") and "~$" not in low:
            if template_name is not None:
                raise Exception("Multiple template .xlsx files found.")
            template_name = name

    if not template_name:
        raise Exception("No template .xlsx file found. Ensure name contains 'template'.")

    return os.path.abspath(os.path.join(path, template_name))


def write_to_excel(flange_obj, template_path: str, filename: str):
    """
    Write flange data into a copy of the Excel template.
    Returns the output path on success, 0 on failure.
    """
    if not getattr(flange_obj, "has_run", False):
        return 0

    headers = flange_obj.headers
    rd1 = flange_obj.xml_data["first"]["records"]
    rd2 = flange_obj.xml_data["second"]["records"]

    try:
        # Create the output report file by copying template
        shutil.copyfile(template_path, filename)

        # Write data into new excel file
        with pd.ExcelWriter(
            path=filename,
            mode="a",
            if_sheet_exists="replace",
            engine="openpyxl",
        ) as writer:
            warnings.filterwarnings(
                action="ignore",
                message="Conditional Formatting extension is not supported and will be removed",
                category=UserWarning,
                module="openpyxl",
            )
            headers.to_excel(writer, sheet_name="Headers")
            rd1.to_excel(writer, sheet_name="First Round")
            rd2.to_excel(writer, sheet_name="Second Round")

        return filename
    except Exception:
        return 0


# Brand colors
vestas_colors = {
    "Blue Sky 01":  "#005AFF",
    "White":        "#FFFFFF",
    "Night Sky":    "#1D3144",
    "Blue Sky 02":  "#4BA6F7",
    "Blue Sky 03":  "#96C8F0",
    "Earth Green":  "#19736E",
    "Earth Red":    "#772219",
    "Earth Orange": "#E17D28",
    "Light Grey":   "#E3E5E8",
    "Medium Grey":  "#A2A9B1",
    "Dark Grey":    "#606D7B",
    "Warm Black":   "#231F20",
    "Black":        "#000000",
}


def _inch_to_fig():
    """Return (inch_h, inch_v) converters for an 8.5x11 page (normalized coords)."""
    inch_h = 1 / 8.5
    inch_v = 1 / 11
    return inch_h, inch_v


def generate_pdf(flange_obj, filepath: str):
    """
    Build the multi-page PDF report from flange_obj data.
    """
    if not getattr(flange_obj, "has_run", False):
        return

    project = flange_obj.location["project"]
    tower = flange_obj.location["tower"]
    flange = flange_obj.location["flange"]
    headers = flange_obj.headers
    records = flange_obj.records.sort_values(by=["BoltNo"])
    required_rotation = flange_obj.required_rotation
    stats = flange_obj.stats

    initials = os.getlogin()
    date = str(dt.today())[:19]

    # Only show headers with actual Approval status
    print_headers = headers.loc[~headers["Approval"].str.contains("N/A", na=False), :]

    # Check for failed bolts
    failed_bolts = records[records["Approval"] == "Fail"]
    has_failures = len(failed_bolts) > 0

    inch_h, inch_v = _inch_to_fig()

    # ===========
    # Main Report Page
    # ===========
    fig, ((ax0, ax1), (ax2, ax3)) = plt.subplots(
        nrows=2, ncols=2,
        figsize=[8.5, 11],
        dpi=400,
        gridspec_kw={"hspace": 0.1, "wspace": 0.05},
    )

    # Layout
    ax0.set_position([0.5 * inch_h, 9.5 * inch_v, 1 - inch_h, 1.0 * inch_v])      # header (smaller)
    ax1.set_position([0.5 * inch_h, 6.3 * inch_v, 1 - inch_h, 2.7 * inch_v])      # header table (compressed)
    ax2.set_position([0.5 * inch_h, 3.5 * inch_v, 1 - inch_h, 2.5 * inch_v])      # stats tables (adjusted)
    ax3.set_position([1.2 * inch_h, 1.0 * inch_v, 0.8 - inch_h, 1.5 * inch_v])      # rotation chart (moved right 0.5")

    # Header (ax0)
    ax0.axis("off")
    ax0.text(
        0.5, 1.10, "Flange Bolt Report",
        ha="center", va="top", fontweight="extra bold", fontsize="xx-large",
        color=vestas_colors["Night Sky"],
    )
    ax0.axhline(0.85, color=vestas_colors["Night Sky"])
    header_text1 = "\n".join([
        f"Project: {project}",
        f"Tower: {tower}",
        f"Flange: {flange}",
        f"Report Date: {date}",
        f"Report Generated By: {initials}",
        "",
    ])
    header_text2 = "\n".join([
        "PDF report generated by flange report Python tool",
        "for Vestas AME Service Engineering.",
        "",
        "VAME flange bolt tension lead: GEOBE",
        "Software maintained by: TOBHI.",
        "",
    ])
    ax0.text(0, 0.75, header_text1, ha="left", va="top", fontweight="bold", fontsize="medium")
    ax0.text(1, 0, header_text2, ha="right", va="bottom", fontsize="small")

    # Header Table (ax1)
    ax1.set_title("Header Data from Xml Files", fontsize="large", fontweight="bold")
    ax1.axis("off")

    vals = np.append(
        np.transpose([print_headers.index.to_numpy()]),
        print_headers.to_numpy(),
        axis=1,
    )
    colors = np.full_like(vals, [vestas_colors["Light Grey"], "w", "w", "w"])
    colors[:, 3] = color_code(print_headers["Approval"])

    table = ax1.table(
        cellText=vals,
        cellColours=colors,
        colLabels=["Parameter", "First Round", "Second Round", "Pass/Fail"],
        colColours=[vestas_colors["Medium Grey"]] * np.shape(colors)[1],
        bbox=[0.05, 0, 0.9, 1],
        loc="center",
    )
    # Format header row
    for i in range(np.shape(colors)[1]):
        table.get_celld()[0, i].set_text_props(wrap=True, fontweight="bold")
    # Body font size
    for (r, c), cell in table.get_celld().items():
        cell.set_text_props(fontsize="small")
    table.auto_set_column_width(range(np.shape(colors)[1]))

    # Stats Tables (ax2)
    ax2.text(
        0.5, 1.0, "Bolt Rotation Data",
        ha="center", va="top", fontweight="extra bold", fontsize="x-large",
        color=vestas_colors["Night Sky"],
    )
    ax2.axhline(0.90, color=vestas_colors["Night Sky"])
    ax2.axis("off")

    # Total stats table (LEFT)
    table = ax2.table(
        cellText=stats["total"].to_numpy(),
        colLabels=["First\nRound", "Second\nRound", "Total\nRotation"],
        rowLabels=["Mean\nRotation", "Standard\nDeviation"],
        colColours=[vestas_colors["Medium Grey"]] * 3,
        rowColours=[vestas_colors["Light Grey"]] * 2,
        bbox=[0.10, 0.15, 0.28, 0.50],
    )
    # Format headers and row labels
    for i in range(3):
        table.get_celld()[0, i].set_text_props(fontweight="bold")
    for i in range(2):
        table.get_celld()[i+1, -1].set_text_props(fontweight="bold")
    # Auto-size ALL columns including row labels
    table.auto_set_column_width([-1, 0, 1, 2])

    # Counts table (RIGHT)
    table = ax2.table(
        cellText=stats["count"].to_numpy().transpose(),
        rowLabels=["First\nRound", "Second\nRound"],
        colLabels=["Cycle 1", "Cycle 2", "Cycle 3+"],
        colColours=[vestas_colors["Medium Grey"]] * 3,
        rowColours=[vestas_colors["Light Grey"]] * 2,
        bbox=[0.62, 0.15, 0.28, 0.50],
    )
    # Format headers and row labels
    for i in range(3):
        table.get_celld()[0, i].set_text_props(fontweight="bold")
    for i in range(2):
        table.get_celld()[i+1, -1].set_text_props(fontweight="bold")
    # Auto-size ALL columns including row labels
    table.auto_set_column_width([-1, 0, 1, 2])

    ax2.text(0.24, 0.83, "Bolt Rotation Angles", fontweight="bold", fontsize="medium", ha="center")
    ax2.text(0.76, 0.83, "Rotated Bolts Per Cycle", fontweight="bold", fontsize="medium", ha="center")

    # Rotation chart (ax3)
    boltnos = records[("BoltNo", "")].to_numpy()
    boltnos = np.append(min(boltnos) - 0.5, np.append(boltnos, max(boltnos) + 0.5))

    vals = records[
        [("First Round", "Cycle 1"),
         ("First Round", "Cycle 2"),
         ("First Round", "Cycle 3+"),
         ("Second Round", "Cycle 1"),
         ("Second Round", "Cycle 2"),
         ("Second Round", "Cycle 3+")]
    ].fillna(0).to_numpy()
    vals = np.vstack([vals[0, :], vals, vals[-1, :]])

    ax3.grid(True)
    ax3.set_axisbelow(True)
    ax3.stackplot(
        boltnos,
        vals.transpose().tolist(),
        colors=[
            vestas_colors["Blue Sky 01"],
            vestas_colors["Blue Sky 01"],
            vestas_colors["Blue Sky 01"],
            vestas_colors["Earth Orange"],
            vestas_colors["Earth Orange"],
            vestas_colors["Earth Orange"],
        ],
        linestyle="-",
        step="mid",
        linewidth=1.25,
        edgecolor=vestas_colors["Medium Grey"],
    )

    rd1_total = np.append(records[("First Round", "Round Total")].iloc[0], records[("First Round", "Round Total")])
    rd1_total = np.append(rd1_total, rd1_total[-1])
    total = np.append(records["Total Rotation"].iloc[0], records["Total Rotation"])
    total = np.append(total, total[-1])

    ax3.step(boltnos, rd1_total, color="w", where="mid", linewidth=1.5)
    ax3.step(boltnos, total, color="k", where="mid", linewidth=1.5)
    ax3.set_xlim(0, max(boltnos) + 1)
    ax3.axhline(required_rotation, color="k", linestyle="--")
    ax3.set_title(f"Bolt Rotation | Required Rotation = {required_rotation}", fontsize="medium", fontweight="bold")
    ax3.set_xlabel("Bolt #")
    ax3.set_ylabel("Rotation Degrees")
    ax3.set_xticks(np.arange(0, max(boltnos) + 1, 5))

    # Save both pages to the same PDF
    with PdfPages(filepath) as pdf:
        # Save main report page
        fig.supxlabel(
            f"Report Generated at {date} for {project}-{tower}-{flange} by user: {initials}",
            fontsize=8,
        )
        pdf.savefig()
        plt.close(fig)

        # Create failure analysis page (only if there are failures)
        if has_failures:
            # ===========
            # Failure Analysis Page - Simplified
            # ===========
            fig_fail = plt.figure(figsize=[8.5, 11], dpi=400)
            ax_fail = fig_fail.add_axes([0.1, 0.1, 0.8, 0.8])

            # Page header
            ax_fail.text(
                0.5, 1.0, f"FAILED BOLTS",
                ha="center", va="top", fontweight="extra bold", fontsize="xx-large",
                color=vestas_colors["Earth Red"],
            )
            ax_fail.axhline(0.95, color=vestas_colors["Earth Red"], linewidth=4)

            # Get failed bolt information
            failed_bolt_nos = sorted(failed_bolts[("BoltNo", "")].tolist())

            # Create the failure report text
            failure_text = f"Number of Failed Bolts: {len(failed_bolt_nos)}\n\n"
            failure_text += "Failed Bolt Numbers:\n"
            failure_text += abbreviate_numbers(failed_bolt_nos)
            failure_text += "\n\nFailure Reason: Insufficient Total Rotation"

            # Display the failure information
            ax_fail.text(
                0.05, 0.85, failure_text,
                ha="left", va="top", fontsize="large", fontweight="bold",
                color=vestas_colors["Night Sky"],
                linespacing=1.5,
            )

            ax_fail.axis("off")

            # Save failure analysis page
            pdf.savefig(fig_fail)
            plt.close(fig_fail)


def abbreviate_numbers(lst: list) -> str:
    """
    Turn a sorted list of ints into a compact string like '1-3, 7-8, 10'.
    """
    res = []
    i = 0
    while i < len(lst):
        j = i
        while j < len(lst) - 1 and lst[j] + 1 == lst[j + 1]:
            j += 1
        if j == i:
            res.append(str(lst[i]))
        else:
            res.append(f"{lst[i]}-{lst[j]}")
        i = j + 1
    return ", ".join(res)
