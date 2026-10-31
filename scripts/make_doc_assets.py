"""Generate the architecture diagram used in the DOCX submission report."""
from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "results" / "figures" / "doc_architecture.png"

NAVY = "#0B1F33"
TEAL = "#0E7C86"
AMBER = "#D9822B"
GRAY = "#6B7A8A"
LIGHT = "#EEF3F6"


def box(ax, x, y, w, h, text, edge=TEAL, face=LIGHT, size=9.8, color=NAVY):
    patch = FancyBboxPatch(
        (x, y), w, h,
        boxstyle="round,pad=0.04,rounding_size=0.10",
        linewidth=1.4, edgecolor=edge, facecolor=face,
    )
    ax.add_patch(patch)
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center",
            fontsize=size, color=color, fontweight="bold", linespacing=1.2)
    return patch


def arrow(ax, x1, y1, x2, y2, color=TEAL, style="-|>", lw=1.4, ls="-"):
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle=style,
                                 mutation_scale=13, color=color, linewidth=lw,
                                 linestyle=ls, shrinkA=0, shrinkB=0))


fig, ax = plt.subplots(figsize=(13.2, 4.8))
ax.set_xlim(0, 13.2)
ax.set_ylim(0, 4.8)
ax.axis("off")

# Main flow
y = 2.35
h = 0.72
w = 1.52
step = 1.74
xs = [0.25 + step * i for i in range(7)]
labels = [
    "Input\nhazy image",
    "Preprocess\n640 px + denoise",
    "Density\n3 statistics",
    "Selector\n3 haze classes",
    "DCP +\nguided filter",
    "Enhancement\nRetinex/CLAHE",
    "Output\nclear image",
]
for x, label in zip(xs, labels):
    box(ax, x, y, w, h, label)
for x in xs[:-1]:
    arrow(ax, x + w + 0.02, y + h / 2, x + step - 0.02, y + h / 2)

# Title
ax.text(0.25, 4.55, "FogSense: density-aware, multi-algorithm dehazing", fontsize=15,
        fontweight="bold", color=NAVY, ha="left")

# Baseline branch
box(ax, 1.99, 1.02, 2.1, 0.60, "Fixed DCP baseline\n$\\omega=0.95$, $t_0=0.10$",
    edge=GRAY, size=9.2)
box(ax, 4.55, 1.02, 1.95, 0.60, "CAP baseline\nZhu et al., 2015",
    edge=GRAY, size=9.2)
box(ax, 6.95, 1.02, 2.10, 0.60, "BCCR-style baseline\nboundary constraint",
    edge=GRAY, size=9.2)
box(ax, 9.55, 1.02, 1.45, 0.60, "Paired\nmetrics", edge=GRAY, size=9.2)
arrow(ax, 4.09, 1.32, 4.53, 1.32, color=GRAY)
arrow(ax, 6.50, 1.32, 6.93, 1.32, color=GRAY)
arrow(ax, 9.05, 1.32, 9.53, 1.32, color=GRAY)
ax.text(0.25, 1.32, "Comparison\nbranch", fontsize=9, color=GRAY,
        fontweight="bold", ha="left", va="center")

# v2 branch
box(ax, 7.30, 3.65, 1.75, 0.55, "v2 airlight\n+ sky", edge=AMBER, face="#FDF1E5", size=9.2)
box(ax, 9.20, 3.65, 1.80, 0.55, "local $\\omega(x)$, $t_0(x)$", edge=AMBER, face="#FDF1E5", size=9.2)
box(ax, 11.15, 3.65, 1.35, 0.55, "spatial\nblend", edge=AMBER, face="#FDF1E5", size=9.2)
arrow(ax, 9.07, 3.92, 9.18, 3.92, color=AMBER, ls="--")
arrow(ax, 11.02, 3.92, 11.13, 3.92, color=AMBER, ls="--")
ax.text(7.15, 3.92, "v2 opt-in\n(off by default)", fontsize=8.8, color=AMBER,
        fontweight="bold", ha="right", va="center")
# dashed feeds into DCP / enhancement
arrow(ax, 8.17, 3.63, 7.95, 3.12, color=AMBER, ls="--", lw=1.2)
arrow(ax, 10.10, 3.63, 9.55, 3.12, color=AMBER, ls="--", lw=1.2)

# Annotations
ax.text(0.25, 0.55, "Shared preprocessing and identical metrics for every method; "
        "all parameters fitted on the calibration split only.",
        fontsize=9.2, color=GRAY, ha="left")
ax.text(0.25, 0.22, "v1 is the recommended default. v2 is evaluated as an opt-in extension, "
        "not silently mixed into the headline result.",
        fontsize=9.2, color=GRAY, ha="left")

fig.tight_layout(pad=0.4)
fig.savefig(OUT, dpi=220, bbox_inches="tight", facecolor="white")
plt.close(fig)
print("wrote", OUT)
