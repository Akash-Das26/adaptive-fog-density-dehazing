"""Build the single-file DOCX submission report for FogSense.

Run with the base Python interpreter (python-docx is installed there):
    C:\\Users\\KIIT\\AppData\\Local\\Programs\\Python\\Python312\\python.exe scripts\\make_submission_docx.py
"""
from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
FIG = ROOT / "results" / "figures"
OUT = ROOT / "report" / "FogSense_Project_Report.docx"

NAVY = RGBColor(0x0B, 0x1F, 0x33)
TEAL = RGBColor(0x0E, 0x7C, 0x86)
AMBER = RGBColor(0xD9, 0x82, 0x2B)
GRAY = RGBColor(0x5B, 0x6B, 0x7A)
LIGHT = "EEF3F6"
LIGHTER = "F7F9FB"
HEADER_FILL = "0B1F33"
BORDER = "C9D2DB"

doc = Document()

# ---------------------------------------------------------------- page setup
sec = doc.sections[0]
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)
sec.top_margin = Cm(2.2)
sec.bottom_margin = Cm(2.0)
sec.left_margin = Cm(2.2)
sec.right_margin = Cm(2.2)
sec.header_distance = Cm(1.0)
sec.footer_distance = Cm(0.9)
sec.different_first_page_header_footer = True

# Normal style
style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(10.5)
style.font.color.rgb = RGBColor(0x1B, 0x27, 0x33)
style.paragraph_format.line_spacing = 1.15
style.paragraph_format.space_after = Pt(6)

# Core properties
cp = doc.core_properties
cp.title = "FogSense: An Intelligent Multi-Algorithm Framework for Adaptive Haze Removal"
cp.subject = "IPA course project — final submission report"
cp.author = "FogSense project group"
cp.keywords = "image dehazing; dark channel prior; fog density; adaptive enhancement; classical image processing"

# ---------------------------------------------------------------- helpers
fig_no = [0]
tab_no = [0]


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def set_cell_margins(cell, top=40, bottom=40, left=70, right=70):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for key, value in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        el = OxmlElement(f"w:{key}")
        el.set(qn("w:w"), str(value))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tcPr.append(mar)


def table_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), BORDER)
        borders.append(el)
    tblPr.append(borders)


def set_repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def cell_text(cell, text, bold=False, size=9.0, color=RGBColor(0x1B, 0x27, 0x33),
              align=WD_ALIGN_PARAGRAPH.LEFT, italic=False):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.space_before = Pt(1)
    r = p.add_run(str(text))
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    r.font.name = "Calibri"
    r.font.color.rgb = color


def heading(text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.space_before = Pt(16 if level == 1 else 11)
    p.paragraph_format.space_after = Pt(6 if level == 1 else 4)
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.bold = True
    if level == 1:
        r.font.size = Pt(16.5)
        r.font.color.rgb = NAVY
        # bottom rule
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "10")
        bottom.set(qn("w:space"), "3")
        bottom.set(qn("w:color"), "0E7C86")
        pBdr.append(bottom)
        pPr.append(pBdr)
    elif level == 2:
        r.font.size = Pt(12.8)
        r.font.color.rgb = TEAL
    else:
        r.font.size = Pt(11.2)
        r.font.color.rgb = NAVY
    return p


def body(text, bold_lead=None, size=10.5, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
         space_after=6):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    if bold_lead:
        r = p.add_run(bold_lead)
        r.bold = True
        r.font.size = Pt(size)
        r.font.color.rgb = NAVY
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.italic = italic
    r.font.color.rgb = RGBColor(0x1B, 0x27, 0x33)
    return p


def bullet(text, bold_lead=None, level=0):
    p = doc.add_paragraph(style="List Bullet" if level == 0 else "List Bullet 2")
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.1
    if bold_lead:
        r = p.add_run(bold_lead)
        r.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = NAVY
    r = p.add_run(text)
    r.font.size = Pt(10.5)
    r.font.color.rgb = RGBColor(0x1B, 0x27, 0x33)
    return p


def caption(text, kind="Figure"):
    if kind == "Figure":
        fig_no[0] += 1
        label = f"Figure {fig_no[0]}."
    else:
        tab_no[0] += 1
        label = f"Table {tab_no[0]}."
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run(f"{label} ")
    r.bold = True
    r.font.size = Pt(8.8)
    r.font.color.rgb = GRAY
    r2 = p.add_run(text)
    r2.italic = True
    r2.font.size = Pt(8.8)
    r2.font.color.rgb = GRAY
    return p


def figure(name, width_cm, text, center=True):
    path = FIG / name
    if not path.exists():
        body(f"[missing figure: {name}]", italic=True)
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(str(path), width=Cm(width_cm))
    caption(text)


def table(headers, rows, widths, text, font_size=8.6, right_cols=(), header_fill=HEADER_FILL,
          highlight_rows=(), shade_rows=(), first_col_bold=False):
    table_obj = doc.add_table(rows=1, cols=len(headers))
    table_obj.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_obj.autofit = False
    table_obj.allow_autofit = False
    table_borders(table_obj)
    hdr = table_obj.rows[0]
    set_repeat_header(hdr)
    for i, h in enumerate(headers):
        c = hdr.cells[i]
        c.width = Cm(widths[i])
        shade(c, header_fill)
        set_cell_margins(c)
        cell_text(c, h, bold=True, size=font_size, color=RGBColor(255, 255, 255),
                  align=WD_ALIGN_PARAGRAPH.RIGHT if i in right_cols else WD_ALIGN_PARAGRAPH.LEFT)
    for ri, row in enumerate(rows):
        cells = table_obj.add_row().cells
        for ci, value in enumerate(row):
            c = cells[ci]
            c.width = Cm(widths[ci])
            set_cell_margins(c)
            if ri in highlight_rows:
                shade(c, "DDEEF2")
            elif ri in shade_rows:
                shade(c, "FBEAE0")
            elif ri % 2 == 1:
                shade(c, LIGHTER)
            cell_text(c, value, bold=(first_col_bold and ci == 0), size=font_size,
                      align=WD_ALIGN_PARAGRAPH.RIGHT if ci in right_cols else WD_ALIGN_PARAGRAPH.LEFT)
    caption(text, kind="Table")
    return table_obj


def code_block(lines, size=8.8):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl.allow_autofit = False
    table_borders(tbl)
    c = tbl.rows[0].cells[0]
    c.width = Cm(16.2)
    shade(c, "F1F4F8")
    set_cell_margins(c, top=90, bottom=90, left=120, right=120)
    c.text = ""
    p = c.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run("\n".join(lines))
    r.font.name = "Consolas"
    r.font.size = Pt(size)
    r.font.color.rgb = RGBColor(0x1B, 0x27, 0x33)


def add_page_field(paragraph):
    run = paragraph.add_run()
    fld1 = OxmlElement("w:fldChar")
    fld1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "separate")
    txt = OxmlElement("w:t")
    txt.text = "1"
    fld3 = OxmlElement("w:fldChar")
    fld3.set(qn("w:fldCharType"), "end")
    run._r.append(fld1)
    run._r.append(instr)
    run._r.append(fld2)
    run._r.append(txt)
    run._r.append(fld3)


def add_numpages_field(paragraph):
    run = paragraph.add_run()
    fld1 = OxmlElement("w:fldChar")
    fld1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "NUMPAGES"
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "separate")
    txt = OxmlElement("w:t")
    txt.text = "1"
    fld3 = OxmlElement("w:fldChar")
    fld3.set(qn("w:fldCharType"), "end")
    run._r.append(fld1)
    run._r.append(instr)
    run._r.append(fld2)
    run._r.append(txt)
    run._r.append(fld3)


def header_footer():
    # header (not first page)
    hp = sec.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hp.paragraph_format.space_after = Pt(0)
    r = hp.add_run("FogSense  \u00b7  IPA Course Project")
    r.font.size = Pt(8.5)
    r.font.color.rgb = GRAY
    pPr = hp._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "4")
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), BORDER)
    pBdr.append(bottom)
    pPr.append(pBdr)

    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = fp.add_run("Page ")
    r1.font.size = Pt(8.5)
    r1.font.color.rgb = GRAY
    add_page_field(fp)
    r2 = fp.add_run(" of ")
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = GRAY
    add_numpages_field(fp)


header_footer()

# ================================================================ COVER
band = doc.add_table(rows=1, cols=1)
band.alignment = WD_TABLE_ALIGNMENT.CENTER
band.autofit = False
band.allow_autofit = False
bc = band.rows[0].cells[0]
bc.width = Cm(16.4)
shade(bc, HEADER_FILL)
set_cell_margins(bc, top=120, bottom=120, left=160, right=160)
p = bc.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
r = p.add_run("IMAGE PROCESSING & ANALYSIS (IPA)  \u00b7  COURSE PROJECT")
r.font.size = Pt(9.5)
r.font.bold = True
r.font.color.rgb = RGBColor(0xB7, 0xD9, 0xDD)
r.font.name = "Calibri"

for _ in range(3):
    doc.add_paragraph()

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(0)
r = p.add_run("FogSense")
r.font.size = Pt(42)
r.font.bold = True
r.font.color.rgb = NAVY
r.font.name = "Segoe UI"

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(10)
r = p.add_run("An Intelligent Multi-Algorithm Framework for Adaptive Haze Removal")
r.font.size = Pt(18)
r.font.color.rgb = TEAL
r.font.name = "Segoe UI"

line = doc.add_paragraph()
line.paragraph_format.space_after = Pt(10)
pPr = line._p.get_or_add_pPr()
pBdr = OxmlElement("w:pBdr")
bottom = OxmlElement("w:bottom")
bottom.set(qn("w:val"), "single")
bottom.set(qn("w:sz"), "12")
bottom.set(qn("w:space"), "1")
bottom.set(qn("w:color"), "0E7C86")
pBdr.append(bottom)
pPr.append(pBdr)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(4)
r = p.add_run("Project Submission Report")
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = NAVY

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(18)
r = p.add_run("Department of Computer Science & Engineering  \u00b7  November 2026")
r.font.size = Pt(10.5)
r.font.color.rgb = GRAY

# group table
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(4)
r = p.add_run("Group Members")
r.font.size = Pt(11)
r.font.bold = True
r.font.color.rgb = TEAL

members = doc.add_table(rows=1, cols=3)
members.alignment = WD_TABLE_ALIGNMENT.CENTER
members.autofit = False
members.allow_autofit = False
table_borders(members)
headers = ["Name", "Roll Number", "Primary contribution"]
widths = [6.2, 4.0, 6.2]
for i, h in enumerate(headers):
    c = members.rows[0].cells[i]
    c.width = Cm(widths[i])
    shade(c, HEADER_FILL)
    set_cell_margins(c)
    cell_text(c, h, bold=True, size=9.0, color=RGBColor(255, 255, 255))
for _ in range(4):
    cells = members.add_row().cells
    for i, w in enumerate(widths):
        cells[i].width = Cm(w)
        set_cell_margins(cells[i])
        cell_text(cells[i], "", size=9.0)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(12)
p.paragraph_format.space_after = Pt(2)
r = p.add_run("Submitted: 06 November 2026")
r.font.size = Pt(10.5)
r.font.bold = True
r.font.color.rgb = NAVY

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(2)
r = p.add_run("Repository: github.com/AtulTejaswi/adaptive-fog-density-dehazing")
r.font.size = Pt(9)
r.font.color.rgb = GRAY

p = doc.add_paragraph()
r = p.add_run("All numerical results in this report are generated from files under results/ "
              "and are reproducible from the committed scripts.")
r.font.size = Pt(8.5)
r.font.italic = True
r.font.color.rgb = GRAY

doc.add_page_break()

# ================================================================ CONTENTS
heading("Contents", 1)
contents = [
    "Abstract (Previously Submitted)",
    "1. Introduction",
    "2. Approach",
    "    2.1 System overview",
    "    2.2 Atmospheric scattering model",
    "    2.3 Preprocessing",
    "    2.4 Fog-density estimation",
    "    2.5 Adaptive strategy selection",
    "    2.6 DCP dehazing and guided filtering",
    "    2.7 Enhancement stack",
    "    2.8 Multi-algorithm baselines",
    "    2.9 v2 opt-in extensions",
    "    2.10 Algorithm summary",
    "    2.11 Evaluation protocol and reproducibility",
    "    2.12 Computational complexity",
    "3. Results",
    "    3.1 Experimental setup",
    "    3.2 Density validation",
    "    3.3 Main SOTS results",
    "    3.4 Per-class results",
    "    3.5 Qualitative comparison",
    "    3.6 Ablation studies",
    "    3.7 v2 evaluation and honest negative result",
    "    3.8 Real-world RTTS results",
    "    3.9 Runtime analysis",
    "    3.10 Failure cases and limitations",
    "    3.11 Baseline fairness audit",
    "    3.12 Discussion and threats to validity",
    "4. Conclusion",
    "References",
    "Appendix A. Reproducibility and repository structure",
    "Appendix B. Metric definitions",
]
for item in contents:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.left_indent = Cm(0.4 if not item.startswith("    ") else 1.0)
    r = p.add_run(item.strip())
    r.font.size = Pt(10.2)
    r.font.color.rgb = NAVY if not item.startswith("    ") else GRAY
    r.bold = not item.startswith("    ")

# ================================================================ ABSTRACT
doc.add_page_break()
heading("Abstract (Previously Submitted)", 1)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
r = p.add_run("The abstract below is the version submitted previously; it is reproduced here for completeness.")
r.italic = True
r.font.size = Pt(9)
r.font.color.rgb = GRAY

body(
    "Haze scattering reduces contrast and colour fidelity, and a fixed-parameter dehazer applies "
    "the same correction to light and dense fog. We present FogSense, a classical, density-aware, "
    "multi-algorithm framework that estimates a fog-density score from three image statistics, "
    "selects a light / moderate / heavy strategy, applies Dark Channel Prior (DCP) dehazing with "
    "class-specific parameters refined by a guided filter, and then optionally applies Retinex "
    "illumination correction, CLAHE and gray-world colour correction. We compare the adaptive "
    "pipeline against a fixed-parameter DCP baseline. On 243 held-out SOTS pairs the adaptive "
    "pipeline improves PSNR from 15.71 dB to 18.78 dB and SSIM from 0.8263 to 0.8896, with a "
    "96.3% PSNR win rate and improvements in every estimated haze class. We also report a "
    "no-reference real-world RTTS set, component/feature/selector ablations, runtime and failure "
    "cases. An opt-in v2 extension (v2 airlight/sky, local density maps, stronger CAP/BCCR "
    "baselines) is evaluated separately: it improves indoor SOTS and CIEDE2000 but does not beat "
    "v1 on aggregate SOTS PSNR/SSIM, and the stronger CAP baseline narrows the adaptive PSNR "
    "advantage."
)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(4)
r = p.add_run("Keywords: ")
r.bold = True
r.font.size = Pt(10)
r.font.color.rgb = NAVY
r2 = p.add_run("single-image dehazing; dark channel prior; fog-density estimation; adaptive image "
               "enhancement; classical image processing; reproducible evaluation")
r2.font.size = Pt(10)
r2.italic = True
r2.font.color.rgb = GRAY

# ================================================================ 1 INTRODUCTION
heading("1. Introduction", 1)
heading("1.1 Background and motivation", 2)
body(
    "Images acquired in fog, haze, smoke, or dust suffer from light scattering before it reaches "
    "the camera. The scattered light adds a veiling component and attenuates the direct scene "
    "radiance, so contrast falls, colours become washed out, and distant structures lose detail. "
    "This is not only a visual-quality problem: dehazing is a preprocessing step for surveillance, "
    "autonomous navigation, remote sensing, driver assistance, and any vision system that must "
    "operate outdoors."
)
body(
    "Mathematically, single-image dehazing is an ill-posed inverse problem. The observed image "
    "provides three colour channels, but the clear radiance, atmospheric light, and depth-dependent "
    "transmission are all unknown. A prior is therefore required. The Dark Channel Prior (DCP) is "
    "one of the most successful classical priors: in haze-free outdoor images, at least one colour "
    "channel is usually very small in a local patch, so the dark channel is a useful proxy for haze "
    "thickness. DCP is simple, explainable, and CPU-friendly, which makes it an excellent fit for a "
    "course project that must remain classical and reproducible."
)
body(
    "The weakness of a standard DCP implementation is that it is usually deployed with one fixed "
    "correction strength for every image. On light haze, a strong fixed setting over-corrects, "
    "producing dark and unnatural results with lost colour. On dense haze, a weak fixed setting "
    "under-corrects and leaves visible fog. The same issue applies to the enhancement chain: a "
    "setting that is appropriate for a lightly hazy indoor scene is often too aggressive for a "
    "dense outdoor scene. The central hypothesis of this project is that estimating fog density "
    "first, and then adapting both the DCP parameters and the enhancement strategy, gives more "
    "consistent results across haze levels."
)

heading("1.2 Problem statement", 2)
body(
    "Given a single hazy RGB image, the system must (i) estimate a fog-density score without a "
    "ground-truth clear image, (ii) map that score to an interpretable haze class, (iii) recover a "
    "clear image using a classical dehazing pipeline, and (iv) evaluate the result with both "
    "full-reference and no-reference metrics. The system must also support a fair comparison "
    "against fixed DCP and stronger classical baselines, and it must not tune any parameter on the "
    "test splits."
)
body(
    "A second requirement is engineering quality. The project must be reproducible: data splits, "
    "calibration artifacts, frozen parameters, evaluation scripts, figures, and report numbers must "
    "all be traceable. With every v2 flag disabled, the output must be byte-identical to v1.0. This "
    "discipline is important because otherwise an apparently improved result may simply be the "
    "product of test-set tuning or an inconsistent baseline."
)

figure("slide_problem.png", 15.5,
       "Fixed DCP on light and heavy haze. A single correction strength over-corrects the light "
       "scene and under-corrects the heavy scene; FogSense adapts the strength to the estimated density.")

heading("1.3 Contributions", 2)
bullet("a label-free fog-density score built from dark-channel mean, entropy and contrast, "
       "normalised on a calibration split and validated against known beta values;", "Density estimation: ")
bullet("a rule-based, per-class parameter selector with constrained monotonicity, plus continuous "
       "and decision-tree selector ablations;", "Adaptive control: ")
bullet("a complete classical pipeline and evaluation harness with fixed DCP, CAP and BCCR-style "
       "baselines, full-reference and no-reference metrics, runtime, ablations and failure cases;",
       "Multi-algorithm evaluation: ")
bullet("an opt-in v2 extension (v2 airlight, sky handling, local density maps, per-pixel parameters, "
       "spatially blended enhancement) evaluated honestly, including a negative aggregate result;",
       "Honest extension: ")
bullet("a reproducible repository with frozen tags, 83 tests, byte-identical v1 compatibility, "
       "a Streamlit demo, a presentation deck and a single-file submission report.",
       "Reproducibility: ")

heading("1.4 Scope and constraints", 2)
body(
    "The project is deliberately classical. No deep dehazing network, no pretrained dehazing model, "
    "and no learned image prior are used. This constraint makes the method explainable and keeps the "
    "runtime practical on a CPU, but it also limits the upper bound on synthetic benchmarks. The "
    "evaluation uses SOTS indoor/outdoor pairs from RESIDE for full-reference metrics and the "
    "real-world RTTS set for no-reference metrics. SOTS is synthetic and RTTS has no ground truth, "
    "so the report distinguishes between measured gains and qualitative trends."
)
body(
    "v1 is the recommended default. v2 is a separate, opt-in branch: all v2 flags default to off, "
    "and the v1 result is preserved. This design decision is intentional. The v2 experiments are "
    "reported as an scientific investigation rather than hidden or presented as an unconditional "
    "improvement."
)

# ================================================================ 2 APPROACH
heading("2. Approach", 1)
heading("2.1 System overview", 2)
body(
    "FogSense takes a single RGB image and produces a dehazed image through a sequence of "
    "explainable stages. The image is first converted to float32 RGB and resized so that its long "
    "side is 640 pixels; a light bilateral filter reduces sensor noise while preserving edges. A "
    "density estimator computes three global statistics and maps them to a score in [0, 1]. A "
    "three-bin selector maps the score to light, moderate, or heavy haze and retrieves a frozen "
    "parameter set. DCP estimates atmospheric light and transmission, the guided filter refines the "
    "transmission map, and the radiance is recovered. Finally, a class-specific enhancement chain "
    "applies Retinex, CLAHE and gray-world colour correction. The same interface can dispatch fixed "
    "DCP, CAP and BCCR-style baselines for a fair comparison."
)
figure("doc_architecture.png", 16.0,
       "FogSense architecture. The main adaptive path is shown in teal; the comparison baselines "
       "are shown in grey; the v2 extensions are shown in amber and are disabled by default.")

heading("2.2 Atmospheric scattering model", 2)
body(
    "The image formation model used throughout the project is the standard atmospheric scattering "
    "model:"
)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("I(x) = J(x) t(x) + A (1 \u2212 t(x)),     t(x) = exp(\u2212\u03b2 d(x))")
r.font.name = "Cambria Math"
r.font.size = Pt(12.5)
r.italic = True
r.font.color.rgb = NAVY

body(
    "Here I is the observed hazy image, J is the clear scene radiance, A is the global atmospheric "
    "light, t is the transmission, \u03b2 is the scattering coefficient, and d is scene depth. The "
    "term A(1 \u2212 t) is the veiling light that reaches the camera after scattering. Inverting the "
    "model requires estimates of A and t. DCP provides both: the dark channel estimates haze "
    "thickness, the brightest dark-channel pixels estimate A, and the transmission is recovered "
    "from the dark channel of the normalised image. A lower bound t\u2080 prevents division by very "
    "small transmission values, which would amplify noise."
)
body(
    "The model is a simplification: it assumes a single global atmospheric light, homogeneous "
    "scattering, and a single scattering event. Real fog can be spatially varying and can contain "
    "multiple scattering. These assumptions are part of the reason the project reports failure "
    "cases and does not claim a solved problem."
)

heading("2.3 Preprocessing", 2)
body(
    "All images are converted to RGB float32 in [0, 1]. The long side is resized to 640 pixels using "
    "area interpolation when downscaling, which avoids aliasing and keeps the evaluation runtime "
    "reasonable. A bilateral filter (diameter 5, sigma colour 20, sigma space 20) is applied as a "
    "light denoiser. The same preprocessing is used for the adaptive method, every baseline, and "
    "the ground truth, so the paired comparison is fair. The baseline fairness audit in Section "
    "3.11 verifies that the fixed DCP baseline is not accidentally handicapped by this choice."
)
body(
    "The native-resolution audit is deliberately included because resizing to 640 pixels costs "
    "roughly 1.6 dB on both methods. That cost is shared, so it does not change the relative "
    "conclusion; it is reported to avoid hiding a known limitation. The project prioritises "
    "reproducible, comparable measurements over a single flattering number."
)

heading("2.4 Fog-density estimation", 2)
body(
    "The density estimator is label-free. For each image it computes three statistics that are "
    "monotone with haze in practice:"
)
bullet("computed with a 15\u00d715 patch. Higher values indicate more haze because the dark channel "
       "is lifted by veiling light.", "Dark-channel mean: ")
bullet("Shannon entropy of the grayscale histogram. Haze compresses the histogram and reduces "
       "entropy, so lower entropy indicates more haze.", "Entropy: ")
bullet("the mean of RMS contrast and Michelson contrast. Haze reduces contrast, so lower contrast "
       "indicates more haze.", "Global contrast: ")
body(
    "Each feature is percentile-normalised using p2 and p98 statistics fitted on the calibration "
    "split only. The score is a weighted combination:"
)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("score = 0.45 \u00b7 n_dc + 0.30 \u00b7 (1 \u2212 n_entropy) + 0.25 \u00b7 (1 \u2212 n_contrast)")
r.font.name = "Cambria Math"
r.font.size = Pt(12.2)
r.italic = True
r.font.color.rgb = NAVY

body(
    "The weights were chosen to reflect the reliability of the features and are stored in "
    "config/params.yaml. The normalisation and the class thresholds live in config/calibration.json. "
    "The thresholds are the tertiles of the calibration score distribution, which gives three "
    "balanced classes on the calibration data without using any test images. The score is not a "
    "physical measurement of beta; it is an operational density index. That distinction is "
    "important for interpreting the validation results."
)
table(
    ["Feature", "Direction", "Weight", "Why it is used"],
    [
        ["Dark-channel mean", "higher = hazier", "0.45", "Direct DCP statistic; strong global haze cue."],
        ["Shannon entropy", "lower = hazier", "0.30", "Captures histogram compression; robust on SOTS."],
        ["Global contrast", "lower = hazier", "0.25", "Captures RMS and Michelson loss; complementary."],
    ],
    [4.0, 3.2, 2.0, 7.0],
    "Density features and weights. Normalisation is fitted on calib only.",
    right_cols=(2,),
)

heading("2.5 Adaptive strategy selection", 2)
body(
    "The score is mapped to three classes using calibration tertiles t\u2081 = 0.3937 and "
    "t\u2082 = 0.5885. Each class has its own omega, t\u2080, guided-filter radius/epsilon, CLAHE "
    "clip limit and Retinex strength. Parameters were fitted on the calibration split with a "
    "constrained two-stage search: first omega \u00d7 t\u2080, then CLAHE clip \u00d7 Retinex "
    "strength. The constraints encode a physical prior: correction strength should not decrease as "
    "estimated density increases."
)
table(
    ["Class", "omega", "t\u2080", "CLAHE clip", "Retinex strength"],
    [
        ["Light", "0.86", "0.35", "1.5", "0.00 (off)"],
        ["Moderate", "0.94", "0.30", "1.5", "0.35"],
        ["Heavy", "0.94", "0.20", "1.5", "0.00 (off)"],
    ],
    [3.4, 2.6, 2.6, 3.4, 4.0],
    "Frozen per-class parameters, fitted on the calibration split only.",
    right_cols=(1, 2, 3, 4),
)
body(
    "The three-bin rule is intentionally simple and interpretable. Two alternatives were evaluated "
    "as ablations: continuous interpolation of the parameters from the score, and a depth-3 "
    "decision tree trained on a separate 180-image synthetic selector set. Neither alternative "
    "improved SSIM: continuous interpolation reached 18.80 dB but only 0.8574 SSIM because it "
    "over-enhanced some images, and the classifier reached 18.40 dB / 0.8780 SSIM. The three-bin "
    "rule remains the primary method because it is the simplest, the most explainable, and the best "
    "on SSIM."
)

heading("2.6 DCP dehazing and guided filtering", 2)
body(
    "The DCP implementation follows the standard formulation. For an image I and a patch "
    "\u03a9(x), the dark channel is the minimum over the colour channels and the patch. The "
    "atmospheric light A is estimated from the brightest 0.1% of dark-channel pixels. The "
    "transmission is"
)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("t(x) = 1 \u2212 \u03c9 \u00b7 dark(I(x) / A)")
r.font.name = "Cambria Math"
r.font.size = Pt(12.2)
r.italic = True
r.font.color.rgb = NAVY

body(
    "Raw transmission maps are blocky and noisy because the dark channel is computed over patches. "
    "The guided filter refines the map using the hazy image as the guide. It fits a local linear "
    "model q_i = a_k I_i + b_k in each window and averages the overlapping estimates, which preserves "
    "strong edges while smoothing flat regions. The project uses the box-filter implementation, so "
    "OpenCV-contrib is not required. The guided-filter radius is 60 and epsilon is 1e-4 for the "
    "baseline; the per-class adaptive configuration uses radius 60 with class-specific epsilon. The "
    "radiance is recovered as"
)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("J(x) = (I(x) \u2212 A) / max(t(x), t\u2080) + A")
r.font.name = "Cambria Math"
r.font.size = Pt(12.2)
r.italic = True
r.font.color.rgb = NAVY

body(
    "The transmission floor t\u2080 limits noise amplification in dense haze. A low floor preserves "
    "more detail but also amplifies sensor noise; a high floor is cleaner but can leave haze. This "
    "trade-off is one reason the per-class parameters differ. The baseline fairness audit shows "
    "that removing the guided filter drops PSNR by about 0.85 dB, so the refinement is helping the "
    "baseline rather than handicapping it."
)
figure("transmission_maps.png", 12.0,
       "Transmission maps for light, moderate and heavy haze. The guided filter preserves edges while "
       "smoothing the patch artefacts of the raw dark channel.")

heading("2.7 Enhancement stack", 2)
body(
    "DCP alone often produces a dark, low-saturation image. The enhancement stack is therefore part "
    "of the method, but it is applied selectively:"
)
bullet("multi-scale Retinex on luminance. The illumination is estimated with Gaussian blurs at "
       "several scales, subtracted in log space, and the reflectance-like result is recombined with "
       "a colour gain map. Retinex is enabled only for the moderate class in the final configuration.",
       "Retinex: ")
bullet("contrast-limited adaptive histogram equalisation on the L channel in LAB. The clip limit is "
       "1.5 and the equalised luminance is blended with the dehazed image at strength 0.4. Working "
       "in LAB preserves hue and saturation.", "CLAHE: ")
bullet("per-channel gains that make the average scene colour approximately gray. This removes the "
       "residual colour cast left by DCP and Retinex and gives the lowest measured colour cast.",
       "Gray-world: ")
body(
    "The ablation in Section 3.6 shows the trade-off clearly. CLAHE alone raises PSNR from 17.77 to "
    "18.18 dB and SSIM from 0.8726 to 0.8850. Adding Retinex raises PSNR further to 19.30 dB but "
    "lowers SSIM to 0.8840, because Retinex can over-sharpen texture and amplify noise. The full "
    "pipeline keeps Retinex only for moderate haze and adds gray-world, which gives the best SSIM "
    "(0.8896) and the lowest colour cast (0.0187). The final selector therefore does not simply "
    "enable every enhancement for every class."
)

heading("2.8 Multi-algorithm baselines", 2)
body(
    "A contribution of the project is that the adaptive method is compared against more than one "
    "classical baseline. The fixed DCP baseline uses omega = 0.95, t\u2080 = 0.10, patch size 15, "
    "guided-filter radius 60 and epsilon 1e-4, with no enhancement. It is a strong, standard "
    "reference for the DCP family."
)
body(
    "CAP (Color Attenuation Prior, Zhu et al., 2015) models depth as a linear function of brightness "
    "V and saturation S: d = \u03b8\u2080 + \u03b8\u2081V + \u03b8\u2082S, with published coefficients. "
    "Transmission is t = exp(\u2212\u03b2d), and atmospheric light is taken from the brightest 0.1% of "
    "the depth map. The published theta coefficients are used; beta and t\u2080 are fixed config "
    "values, never fitted on the test splits. CAP is a much stronger PSNR baseline than fixed DCP, "
    "which is why it is included."
)
body(
    "The BCCR-style baseline follows the boundary-constraint idea of Meng et al. (2013). The "
    "dark-channel transmission is projected onto the physical constraint 0 \u2264 J \u2264 1 and "
    "refined with the guided filter. This is a simplified BCCR, not the full L1 contextual solver. "
    "It is included to test whether the adaptive advantage survives a second classical prior. The "
    "BCCR-style result is the most natural on the no-reference NIQE proxy, but it under-enhances "
    "detail and loses PSNR/SSIM."
)
body(
    "All methods share the same preprocessing, the same test images, and the same metric code. The "
    "baseline fairness audit re-runs the fixed DCP baseline with alternative atmospheric-light "
    "estimators, without the guided filter, and at native resolution. The audit confirms that the "
    "default estimator and the guided filter are helping the baseline, not hurting it."
)

heading("2.9 v2 opt-in extensions", 2)
body(
    "The v2 branch investigates three ideas. It is opt-in because the v1 pipeline is already a "
    "complete, defensible method, and because the v2 ideas were fitted on a small calibration set. "
    "All v2 knobs were frozen at tag v2.0-frozen-params before the first v2 test run."
)
body(
    "First, v2 atmospheric light takes the brightest 0.5% of dark-channel pixels, rejects small "
    "connected components (less than 0.05% of pixels) and high-gradient pixels (above the 60th "
    "percentile of Sobel magnitude), and averages the survivors. If too few pixels survive, it "
    "falls back to v1. A quad-tree sky search was implemented as an alternative but was rejected on "
    "calibration because it lost 0.88 dB. This is a good example of a negative result that is "
    "reported rather than hidden."
)
body(
    "Second, a conservative sky/white-object mask is built from brightness (at least 0.80), "
    "smoothness (gradient at most 0.04) and connected area (at least 2% of the image). Inside the "
    "mask, transmission is floored at 0.25 and the output is blended 25% toward the input. The "
    "intent is to avoid darkening bright sky regions that violate the dark-channel assumption. The "
    "mask is deliberately conservative; it can still false-positive on bright indoor surfaces."
)
body(
    "Third, a local density map is built from local dark-channel mean, rank entropy and local RMS "
    "contrast, normalised with calibration-only statistics and smoothed by a guided filter. The map "
    "drives per-pixel omega(x) and t\u2080(x), interpolated from the three class centroids with "
    "monotonicity constraints: omega is non-decreasing and t\u2080 is non-increasing in density. "
    "Enhancement is blended spatially as out = (1 \u2212 alpha) \u00b7 dehazed + alpha \u00b7 enhanced, "
    "with alpha proportional to local density and a floor of 0.0 selected on calibration."
)
figure("v2_density_maps.png", 15.0,
       "v2 local density, per-pixel omega and sky mask. The green mask illustrates the conservative "
       "sky/white-object region; bright indoor surfaces show why the mask is not universally safe.")

heading("2.10 Algorithm summary", 2)
body(
    "The inference procedure for the default v1 pipeline is summarised below. The v2 path differs "
    "only in the airlight/sky stage, the local map, and the spatially blended enhancement."
)
code_block([
    "Algorithm 1: FogSense inference (v1 default; v2 is an opt-in variant)",
    "Input: hazy RGB image I; frozen parameter set \u03b8",
    "1:  I \u2190 resize_long_side(I, 640); I \u2190 bilateral(I)",
    "2:  f \u2190 [dark_channel_mean(I), entropy(I), contrast(I)]",
    "3:  s \u2190 0.45\u00b7n_dc + 0.30\u00b7(1\u2212n_ent) + 0.25\u00b7(1\u2212n_con)",
    "4:  c \u2190 light if s < t1; moderate if t1 \u2264 s < t2; heavy otherwise",
    "5:  A \u2190 mean of brightest 0.1% dark-channel pixels",
    "6:  t \u2190 guided_filter(1 \u2212 \u03c9_c \u00b7 dark(I/A))",
    "7:  J \u2190 (I \u2212 A) / max(t, t0_c) + A",
    "8:  J \u2190 enhancement_chain(J, c)   // Retinex \u2192 CLAHE \u2192 gray-world",
    "Output: dehazed image J, density score s, class c, and maps for the demo",
], size=8.4)

heading("2.11 Evaluation protocol and reproducibility", 2)
body(
    "The evaluation is designed to answer three questions: does the density score behave sensibly, "
    "does the adaptive pipeline beat a fair fixed baseline, and which component is responsible for "
    "the gain. The data are split by unique clear scene, so the same scene cannot appear in both "
    "calibration and test. The calibration split is used for density normalisation, tertile "
    "thresholds, per-class parameter search and the NIQE-style model. The test splits are used only "
    "after the parameters are written and the repository is tagged."
)
table(
    ["Split", "Images", "Purpose"],
    [
        ["Calibration", "299 pairs", "Density normalisation, thresholds, per-class tuning, NIQE-style fit."],
        ["SOTS test", "243 pairs", "Full-reference PSNR/SSIM/CIEDE2000 evaluation."],
        ["RTTS", "72 images", "Real-world no-reference evaluation; no clear ground truth."],
        ["Density validation", "64 images", "8 scenes \u00d7 8 beta values for monotonicity checks."],
        ["Selector set", "180 images", "Synthetic data for the decision-tree selector ablation."],
    ],
    [3.7, 2.8, 9.5],
    "Data splits used in the project. Calibration and test are separated by unique clear scene.",
)
body(
    "Full-reference metrics are PSNR and SSIM. CIEDE2000 is the mean colour difference in Lab "
    "(lower is better). No-reference metrics are entropy gain, RMS contrast improvement, Sobel "
    "edge-preservation ratio, colour cast, NIQE-style and FADE-style. NIQE-style is a 20-statistic "
    "MSCN model fitted on the 299 calibration clear images. FADE-style uses the Choi et al. feature "
    "set with fixed documented weights. Both are proxies, not the official pretrained models, and "
    "are interpreted as relative trends."
)
body(
    "Paired comparisons use the Wilcoxon signed-rank test on per-image differences. The report "
    "states win rates and p-values rather than only means. Runtime is measured per image on CPU "
    "with identical preprocessing. The frozen tags are v0.1-frozen-params, v1.0, "
    "v2.0-frozen-params and v2.0. The v1 tags were never moved. With every v2 flag off, the output "
    "is byte-identical to v1.0, which is enforced by tests/test_v2_compat.py."
)

heading("2.12 Computational complexity", 2)
body(
    "All stages are linear in the number of pixels N for a fixed filter radius. Preprocessing, the "
    "dark channel, the guided filter and the enhancement steps are O(N). The density features add "
    "O(N + 256) for the histogram. The v2 local density map is the expensive part: rank entropy in "
    "a disk of radius 15 is O(N r\u00b2) with a small constant, and the local statistics are O(N). "
    "Memory is O(N) for the image and its float maps. On the SOTS test set, fixed DCP takes about "
    "59 ms, v1 about 192 ms, and v2 about 562 ms per 640-pixel image on the same CPU run. The v2 "
    "local density map accounts for roughly 405 ms, which is why v2 is about 2.9\u00d7 slower than "
    "v1."
)

# ================================================================ 3 RESULTS
doc.add_page_break()
heading("3. Results", 1)
heading("3.1 Experimental setup", 2)
body(
    "The main full-reference experiment uses 243 SOTS test pairs (22 indoor and 221 outdoor). Five "
    "methods are evaluated on every image with identical preprocessing: fixed DCP, CAP, BCCR-style, "
    "FogSense v1 and FogSense v2. The v1 and v2 parameters are frozen; the v2 run uses the "
    "configuration frozen at v2.0-frozen-params. The RTTS experiment uses 72 real-world images with "
    "no ground truth. All metrics are written to results/ and the report tables are generated from "
    "those files."
)

heading("3.2 Density validation", 2)
body(
    "The density score is validated on the 64-image synthetic sweep: 8 scenes at 8 beta values. The "
    "score increases monotonically with beta in all 8 scenes. The overall Spearman correlation is "
    "0.907, and the stricter within-scene mean is 0.992 with a minimum of 0.939. This is strong "
    "evidence that the score is measuring the intended quantity rather than an unrelated image "
    "statistic."
)
figure("score_vs_beta.png", 15.0,
       "Density score versus beta on the 64-image validation sweep. The score is monotone in beta for "
       "all eight scenes.")
table(
    ["Signal", "vs beta (validation)", "Within-scene", "vs FADE (RTTS)"],
    [
        ["Global density (shipped)", "0.907", "0.992", "0.933"],
        ["Local density mean", "0.948", "1.000", "0.817"],
        ["Dark channel only", "0.928", "\u2014", "0.762"],
        ["Dark channel + entropy", "0.917", "\u2014", "0.904"],
        ["Raw contrast std", "\u22120.780", "\u2014", "\u22120.611"],
        ["Ridge weights (experiment)", "0.925", "1.000", "0.969"],
    ],
    [5.2, 4.2, 3.2, 3.4],
    "Density-signal correlations. Local density is best on the beta sweep; global density tracks "
    "FADE better on real RTTS images. Ridge is an experiment, not the shipped estimator.",
    right_cols=(1, 2, 3),
)
body(
    "The ordering flips on real RTTS images: global density correlates better with FADE (0.933 vs "
    "0.817), while the calibration-fitted ridge weights are best (0.969). The ridge weights beat the "
    "global score on the beta sweep by +0.017 but not the local mean, so they are reported as an "
    "experiment rather than shipped. This is a useful lesson: a signal that is best on a controlled "
    "sweep is not automatically best on real images with mixed degradation."
)

heading("3.3 Main SOTS results", 2)
body(
    "Table 5 reports the main SOTS test result. FogSense v1 reaches 18.78 dB PSNR and 0.8896 SSIM, "
    "compared with 15.71 dB and 0.8263 for fixed DCP. The improvement is +3.07 dB PSNR and +0.0633 "
    "SSIM. The paired win rates are 96.3% for PSNR and 90.1% for SSIM. The gain is not confined to "
    "one haze class: every estimated class improves, as shown in Section 3.4."
)
table(
    ["Method", "PSNR (dB)", "SSIM", "CIEDE2000", "NIQE-style", "FADE-style", "Runtime (ms)"],
    [
        ["Fixed DCP", "15.71", "0.8263", "12.64", "1.311", "0.3735", "59"],
        ["CAP", "18.09", "0.7833", "9.84", "2.031", "0.3911", "38"],
        ["BCCR-style", "16.02", "0.8309", "12.36", "1.267", "0.3753", "83"],
        ["FogSense v1", "18.78", "0.8896", "9.50", "1.685", "0.4051", "192"],
        ["FogSense v2", "18.45", "0.8873", "9.38", "1.794", "0.4007", "562"],
    ],
    [3.1, 2.2, 2.0, 2.3, 2.3, 2.3, 2.4],
    "SOTS test (243 pairs), identical preprocessing. Lower is better for CIEDE2000, NIQE-style and "
    "FADE-style. v1 is highlighted as the default; v2 is the opt-in extension.",
    right_cols=(1, 2, 3, 4, 5, 6),
    highlight_rows=(3,),
    shade_rows=(4,),
)
body(
    "CAP is a much stronger PSNR baseline than fixed DCP: 18.09 dB vs 15.71 dB. This is an "
    "important result because it prevents an over-optimistic story. Against CAP, v1 still has a "
    "large SSIM advantage (0.8896 vs 0.7833) and a better CIEDE2000 (9.50 vs 9.84), but the PSNR "
    "advantage is smaller than against fixed DCP. The adaptive method is therefore best described "
    "as a structural and colour improvement, not a universal PSNR win over every classical prior."
)
body(
    "The BCCR-style baseline has the best NIQE-style score (1.267) on SOTS, which suggests it "
    "produces the most 'natural' images, but it under-enhances detail and loses PSNR/SSIM. The "
    "lesson is that no single metric tells the whole story. The report therefore keeps all three "
    "families of evidence: full-reference accuracy, colour difference, and no-reference naturalness."
)

heading("3.4 Per-class results", 2)
body(
    "The adaptive method improves every estimated haze class. The largest PSNR gain is in the "
    "moderate class (+4.11 dB), followed by light (+3.24 dB) and heavy (+1.37 dB). The heavy class "
    "is the hardest because the transmission is small, noise amplification is stronger, and the "
    "t\u2080 floor trades detail for stability. The SSIM gain is also positive in every class."
)
table(
    ["Estimated class", "Method", "PSNR (dB)", "SSIM", "\u0394 PSNR", "\u0394 SSIM"],
    [
        ["Light", "Fixed DCP", "16.58", "0.8399", "\u2014", "\u2014"],
        ["Light", "FogSense v1", "19.82", "0.9064", "+3.24", "+0.0665"],
        ["Moderate", "Fixed DCP", "15.32", "0.8165", "\u2014", "\u2014"],
        ["Moderate", "FogSense v1", "19.44", "0.8902", "+4.11", "+0.0737"],
        ["Heavy", "Fixed DCP", "15.14", "0.8225", "\u2014", "\u2014"],
        ["Heavy", "FogSense v1", "16.51", "0.8672", "+1.37", "+0.0446"],
    ],
    [3.2, 3.0, 2.5, 2.3, 2.3, 2.3],
    "Per-class SOTS results. The class labels are estimated from the density score; SOTS does not "
    "provide per-image beta labels.",
    right_cols=(2, 3, 4, 5),
)
body(
    "The per-class labels are an operational definition, not ground truth. They are useful for "
    "diagnosis but should not be interpreted as a physical beta measurement. This caveat is stated "
    "in the report and is one of the limitations of the evaluation."
)

heading("3.5 Qualitative comparison", 2)
body(
    "Figure 6 shows SOTS indoor and outdoor examples and two RTTS scenes, comparing the hazy input, "
    "fixed DCP, CAP, BCCR-style, v1 and v2. v1 consistently recovers contrast and detail without the "
    "heavy darkening of fixed DCP. CAP produces strong global contrast but can lose colour and "
    "structure, which is consistent with its lower SSIM. BCCR-style is visually natural but "
    "under-enhances. v2 helps some indoor bright regions but can over-correct outdoors, matching the "
    "quantitative outdoor loss."
)
figure("v2_qualitative.png", 16.0,
       "Qualitative comparison on SOTS indoor/outdoor and RTTS. Columns: hazy, fixed DCP, CAP, "
       "BCCR-style, FogSense v1, FogSense v2.")

heading("3.6 Ablation studies", 2)
body(
    "The v1 ablation isolates the contribution of the enhancement chain, the density features and "
    "the selector. Removing all enhancement lowers SSIM from 0.8896 to 0.8726 and PSNR from 18.78 to "
    "17.77 dB, so the enhancement chain is doing real work beyond DCP. CLAHE alone reaches 18.18 dB / "
    "0.8850. Adding Retinex raises PSNR to 19.30 dB but lowers SSIM to 0.8840, and adds colour cast. "
    "The full pipeline keeps Retinex only for moderate haze and adds gray-world, which gives the best "
    "SSIM and the lowest colour cast."
)
table(
    ["Variant", "Group", "PSNR (dB)", "SSIM", "Entropy gain", "Edge preservation", "Colour cast"],
    [
        ["Fixed DCP", "baseline", "15.71", "0.8263", "0.1929", "1.3360", "0.0376"],
        ["No enhancement", "component", "17.77", "0.8726", "0.2428", "1.3835", "0.0337"],
        ["CLAHE only", "component", "18.18", "0.8850", "0.3697", "1.7287", "0.0339"],
        ["CLAHE + Retinex", "component", "19.30", "0.8840", "0.5337", "2.0560", "0.0389"],
        ["Full v1", "component", "18.78", "0.8896", "0.4258", "1.8484", "0.0187"],
        ["Density: DC only", "density", "18.48", "0.8831", "0.4121", "1.8218", "0.0185"],
        ["Density: DC + entropy", "density", "18.84", "0.8901", "0.4276", "1.8496", "0.0183"],
        ["Selector: continuous", "selector", "18.80", "0.8574", "0.5385", "2.1860", "0.0170"],
        ["Selector: classifier", "selector", "18.40", "0.8780", "0.4464", "1.8592", "0.0201"],
    ],
    [3.5, 2.0, 2.1, 1.8, 2.2, 2.4, 2.0],
    "v1 ablation on the 243 SOTS test pairs. All variants use the same preprocessing and frozen "
    "calibration artifacts.",
    right_cols=(2, 3, 4, 5, 6),
    highlight_rows=(4,),
)
body(
    "The density ablation shows that dark-channel-only density reaches 18.48 dB / 0.8831 SSIM; "
    "adding entropy gives 18.84 dB / 0.8901; adding contrast gives 18.78 dB / 0.8896. Entropy "
    "therefore carries most of the density signal on SOTS, while contrast is close to neutral. The "
    "0.0005 SSIM difference between entropy-only and full density is within run-to-run and sampling "
    "noise, so the report does not claim a significant advantage for the third feature. All three "
    "features are kept because they are complementary on the beta sweep and on real low-contrast "
    "images."
)
body(
    "The selector ablation is equally informative. Continuous interpolation reaches 18.80 dB but "
    "only 0.8574 SSIM because it over-enhances some images. The depth-3 classifier reaches 18.40 dB / "
    "0.8780 SSIM. The three-bin selector is retained because it is the simplest, the most "
    "interpretable, and the best on SSIM. This is a design choice, not a claim that three bins are "
    "universally optimal."
)

heading("3.7 v2 evaluation and honest negative result", 2)
body(
    "The v2 extension is evaluated on the same 243 SOTS pairs with the frozen v2 configuration. The "
    "result is not an aggregate win. v2 reaches 18.45 dB / 0.8873 SSIM, which is 0.32 dB lower PSNR "
    "and 0.0023 lower SSIM than v1. The PSNR difference is statistically significant "
    "(p = 0.0002); the SSIM difference is not (p = 0.4416). CIEDE2000 improves by 0.120 (lower is "
    "better), NIQE-style worsens by 0.109, and FADE-style improves by 0.0044. The honest conclusion "
    "is that the three v2 upgrades did not generalise from calibration to the SOTS test average."
)
table(
    ["Group", "n", "v1 PSNR", "v2 PSNR", "\u0394 PSNR", "v1 SSIM", "v2 SSIM", "\u0394 SSIM"],
    [
        ["Overall", "243", "18.775", "18.450", "\u22120.325", "0.8896", "0.8873", "\u22120.0023"],
        ["Indoor", "22", "19.811", "20.558", "+0.747", "0.8410", "0.8528", "+0.0118"],
        ["Outdoor", "221", "18.672", "18.241", "\u22120.431", "0.8944", "0.8907", "\u22120.0037"],
    ],
    [2.0, 1.3, 2.2, 2.2, 2.0, 2.1, 2.1, 2.1],
    "v2 versus v1 on SOTS. The indoor gain is real but based on only 22 pairs; the outdoor loss "
    "dominates the aggregate.",
    right_cols=(1, 2, 3, 4, 5, 6, 7),
)
body(
    "The indoor subset improves by +0.75 dB PSNR and +0.0118 SSIM, while the outdoor subset loses "
    "0.43 dB PSNR and 0.0037 SSIM. The indoor subset is only 22 pairs, so the split is descriptive, "
    "not a powered test. The sky/local intervention is plausible for indoor scenes with large bright "
    "regions; outdoor scenes dominate the average and are where the sky mask and local map most often "
    "over- or under-correct."
)
figure("slide_v2_delta.png", 14.0,
       "v2 minus v1 on SOTS. The indoor gain is visible, but the outdoor loss and the overall PSNR "
       "loss dominate the aggregate. v1 remains the default.")
body(
    "The step-by-step v2 ablation explains the result. The airlight change is essentially neutral "
    "(+0.10 dB / \u22120.0001 SSIM). Adding the sky stage improves both metrics (+0.53 dB / +0.0038 "
    "SSIM), and airlight+sky alone reaches 19.41 dB / 0.8933, which is +0.64 dB / +0.0037 over v1. "
    "The local omega/t\u2080 maps as configured lose 0.87 dB and 0.0119 SSIM. The spatially blended "
    "enhancement recovers 0.0058 SSIM but not PSNR. Because dropping the local maps after seeing "
    "the test ablation would be test-set tuning, the frozen full-v2 configuration is kept and the "
    "v2-light variant is reported only as a direction for future work."
)
table(
    ["v2 variant", "PSNR (dB)", "SSIM", "CIEDE2000", "NIQE-style", "FADE-style", "Runtime (ms)"],
    [
        ["v1", "18.775", "0.8896", "9.501", "1.685", "0.4051", "228"],
        ["+ airlight", "18.879", "0.8895", "9.446", "1.723", "0.4045", "241"],
        ["+ airlight + sky", "19.410", "0.8933", "8.990", "1.804", "0.4067", "269"],
        ["+ local omega/t0", "18.544", "0.8815", "9.637", "1.897", "0.4061", "666"],
        ["+ local enhancement", "18.450", "0.8873", "9.381", "1.794", "0.4007", "680"],
        ["Full v2 (frozen)", "18.450", "0.8873", "9.381", "1.794", "0.4007", "678"],
    ],
    [3.4, 2.0, 1.8, 2.2, 2.2, 2.2, 2.2],
    "v2 ablation on SOTS. The sky stage is the only step that improves both PSNR and SSIM; the local "
    "maps cost PSNR.",
    right_cols=(1, 2, 3, 4, 5, 6),
    highlight_rows=(2,),
)
body(
    "The brightness/colour diagnostic is reported but not used as a headline result. Matching each "
    "output's per-channel mean and standard deviation to the ground truth before PSNR raises v1 by "
    "3.11 dB and v2 by 3.80 dB; after matching, v2 is +0.37 dB relative to v1. This indicates that "
    "part of v2's raw PSNR loss is a global brightness/colour shift rather than lost structure. It "
    "is a diagnostic, not a claim that v2 is better."
)

heading("3.8 Real-world RTTS results", 2)
body(
    "RTTS contains 72 real hazy photographs with no clear ground truth. Only no-reference metrics are "
    "reported. v1 leads on entropy gain (0.4165 vs 0.3063 for v2) and edge preservation (1.6733 vs "
    "1.4729). v2 has a slightly lower FADE (0.4586 vs 0.4677) but a higher NIQE (1.7760 vs 1.4982) "
    "and more colour cast (0.0174 vs 0.0129). BCCR-style has the best NIQE (1.2476), while CAP has "
    "the worst entropy gain (negative), showing that strong synthetic PSNR does not automatically "
    "translate to natural real-world appearance."
)
table(
    ["Method", "Entropy gain", "Contrast", "Edge preservation", "Colour cast", "NIQE-style", "FADE-style", "Runtime (ms)"],
    [
        ["Fixed DCP", "0.0912", "1.0648", "1.1967", "0.0263", "1.3766", "0.4451", "53"],
        ["CAP", "\u22120.0564", "1.1775", "1.0864", "0.0225", "1.8407", "0.4695", "37"],
        ["BCCR-style", "0.1113", "1.0772", "1.2025", "0.0259", "1.2476", "0.4448", "77"],
        ["FogSense v1", "0.4165", "1.2376", "1.6733", "0.0129", "1.4982", "0.4677", "132"],
        ["FogSense v2", "0.3063", "1.2196", "1.4729", "0.0174", "1.7760", "0.4586", "479"],
    ],
    [2.4, 1.9, 1.8, 2.2, 1.8, 2.0, 2.0, 2.0],
    "RTTS no-reference results (72 images). RTTS has no ground truth, so these are relative trends, "
    "not full-reference accuracy.",
    right_cols=(1, 2, 3, 4, 5, 6, 7),
)
body(
    "No single method wins every no-reference metric. v1 is strongest on detail/contrast/edge "
    "preservation, BCCR-style is strongest on NIQE naturalness, and v2 is slightly better on FADE. "
    "The report does not average these into a single score, because the metrics measure different "
    "properties and can conflict."
)

heading("3.9 Runtime analysis", 2)
body(
    "On the SOTS test set, fixed DCP takes about 59 ms per image, CAP 38 ms, BCCR-style 83 ms, v1 "
    "192 ms, and v2 562 ms on the same CPU run. v1 is about 3.2\u00d7 slower than fixed DCP, mainly "
    "because of the enhancement chain; v2 is about 2.9\u00d7 slower than v1, mainly because of the "
    "local density map. The local density map accounts for roughly 405 ms, while the airlight and "
    "sky stages are comparatively cheap. Fast Retinex was implemented and rejected: it lost 0.31 dB "
    "PSNR and 0.0042 SSIM for only about 40 ms, so it was not worth the quality cost."
)
figure("slide_runtime.png", 12.0,
       "Runtime per 640-pixel image on SOTS (same run, identical preprocessing). Lower is better.")
table(
    ["Method", "SOTS (ms)", "RTTS (ms)", "Relative to v1"],
    [
        ["Fixed DCP", "59.4", "53.2", "0.31\u00d7"],
        ["CAP", "38.2", "37.0", "0.20\u00d7"],
        ["BCCR-style", "83.1", "76.9", "0.43\u00d7"],
        ["FogSense v1", "192.3", "131.5", "1.00\u00d7"],
        ["FogSense v2", "561.7", "479.0", "2.92\u00d7"],
    ],
    [4.0, 3.2, 3.2, 4.0],
    "Runtime summary. The v1 runtime in the common v2 run (192 ms) differs from the v1-only run "
    "(256 ms) because of machine load; only relative conclusions are used.",
    right_cols=(1, 2, 3),
)
body(
    "The runtime trade-off is a real limitation. v1 is suitable for offline or moderate-throughput "
    "use, but not for high-frame-rate video on a CPU. v2 is an experiment, not a production choice. "
    "The project reports this openly rather than hiding it."
)

heading("3.10 Failure cases and limitations", 2)
body(
    "The failure cases are included honestly. The most important failure mode is sky and white "
    "objects, which violate the dark-channel assumption: transmission is underestimated, so these "
    "regions darken or shift colour. Night haze has low illumination and weak colour statistics, so "
    "the density score is less reliable. Dense haze amplifies noise because J divides by a small t; "
    "the t\u2080 floor trades detail for noise. CLAHE and Retinex can over-sharpen texture and amplify "
    "compression artefacts. The real RTTS set has no ground truth, so its no-reference metrics only "
    "indicate trends. The SOTS haze class is estimated from the density score, not from a per-image "
    "beta label."
)
figure("failure_cases.png", 14.0,
       "Lowest-SSIM adaptive results on SOTS, included honestly. These cases show the sky/white-object, "
       "noise-amplification and over-sharpening failure modes.")
body(
    "The v2-specific limitations are also important: the indoor gain is based on 22 pairs, the "
    "outdoor loss dominates the average, the sky mask can false-positive on bright indoor surfaces, "
    "and v2 is about 2.9\u00d7 slower than v1. The NIQE-style and FADE-style metrics are proxies, not "
    "official learned models. These are not minor caveats; they define the boundary of what the "
    "project can claim."
)

heading("3.11 Baseline fairness audit", 2)
body(
    "A fair comparison requires a baseline that is not accidentally handicapped. The audit re-runs "
    "fixed DCP on 20 evenly spaced SOTS test images with alternative atmospheric-light estimators, "
    "without the guided filter, and at native resolution. The default estimator and the guided filter "
    "are both helping the baseline: removing the guided filter drops PSNR from 15.44 to 14.59 dB. The "
    "native-resolution variant scores higher (17.02 dB) because the shared 640-pixel resize and "
    "bilateral denoise cost about 1.6 dB; since baseline, adaptive and ground truth all use the same "
    "preprocessing, the paired comparison remains fair."
)
table(
    ["Baseline variant", "PSNR (dB)", "SSIM", "Mean A"],
    [
        ["Current (top-0.1% mean, GF r=60)", "15.44", "0.8330", "0.941"],
        ["Brightest dark-channel pixel", "15.42", "0.8332", "0.943"],
        ["Top-1% mean atmospheric light", "15.59", "0.8342", "0.934"],
        ["No guided filter", "14.59", "0.7752", "0.941"],
        ["Native resolution (no resize/denoise)", "17.02", "0.8546", "0.939"],
    ],
    [7.0, 3.2, 3.0, 3.0],
    "Baseline fairness audit on 20 SOTS test images. The default baseline is not handicapped; "
    "native resolution is higher because the shared preprocessing costs about 1.6 dB.",
    right_cols=(1, 2, 3),
)
body(
    "This audit is a strength of the project. It shows that the headline comparison is not the "
    "result of a weak baseline, and it documents the cost of the shared preprocessing instead of "
    "hiding it."
)

heading("3.12 Discussion and threats to validity", 2)
body(
    "The main positive result is robust: density-aware adaptation improves fixed DCP by a large "
    "margin on SOTS, with high win rates and consistent per-class gains. The comparison with CAP "
    "makes the claim more nuanced: v1's PSNR advantage over CAP is smaller, but its SSIM and colour "
    "difference are clearly better. The adaptive method is therefore best interpreted as a "
    "structural and colour-quality improvement, not a universal PSNR record."
)
body(
    "The v2 result is a genuine negative result. The airlight change is neutral, the sky stage is "
    "useful, and the local omega/t\u2080 maps are harmful as configured. The fact that the sky stage "
    "alone reaches 19.41 dB / 0.8933 suggests a promising direction, but selecting it after seeing "
    "the test ablation would be test-set tuning. The frozen full-v2 configuration is therefore kept, "
    "and the v2-light variant is reported only as future work."
)
body(
    "Threats to validity include the synthetic nature of SOTS, the small indoor subset (22 pairs), "
    "the absence of ground truth on RTTS, the proxy nature of NIQE/FADE, the single global "
    "atmospheric light assumption, and the 640-pixel preprocessing. The paired design, frozen "
    "parameters, baseline audit and per-image CSVs mitigate these threats but do not eliminate them. "
    "The report states them explicitly so that the conclusions are interpreted at the right scope."
)

# ================================================================ 4 CONCLUSION
heading("4. Conclusion", 1)
body(
    "This project presented FogSense, a classical, density-aware, multi-algorithm framework for "
    "single-image haze removal. The method estimates fog density from three label-free statistics, "
    "selects a light / moderate / heavy strategy, applies DCP with class-specific parameters and a "
    "guided filter, and optionally applies Retinex, CLAHE and gray-world colour correction. The same "
    "interface supports fixed DCP, CAP and BCCR-style baselines, which makes the evaluation more "
    "credible than a single-baseline comparison."
)
body(
    "On 243 held-out SOTS pairs, FogSense v1 improves fixed DCP from 15.71 dB to 18.78 dB PSNR and "
    "from 0.8263 to 0.8896 SSIM, with a 96.3% PSNR win rate and a 90.1% SSIM win rate. The gains "
    "hold in every estimated haze class. CAP is a much stronger PSNR baseline than fixed DCP, but "
    "v1 retains a clear SSIM and colour advantage. The ablation shows that the enhancement chain and "
    "the density features both contribute, and that the three-bin selector is the best SSIM choice "
    "among the tested selector variants."
)
body(
    "The v2 extension is reported honestly as a partial/negative result. v2 improves indoor SOTS and "
    "CIEDE2000, but it loses 0.32 dB PSNR on the aggregate SOTS test because the local omega/t\u2080 "
    "maps do not generalise. The sky stage alone is promising (+0.64 dB / +0.0037 SSIM over v1), but "
    "it is not adopted post hoc because that would be test-set tuning. v1 therefore remains the "
    "recommended default."
)
body(
    "The project also contributes engineering reproducibility: frozen tags, 83 tests, byte-identical "
    "v1 compatibility with all v2 flags off, a complete evaluation harness, a Streamlit demo, a "
    "22-slide presentation deck with speaker notes, and this single-file report. Every number in the "
    "report is generated from files under results/ and can be regenerated from the committed "
    "scripts."
)
body("Future work should focus on four directions:", bold_lead="")
bullet("replace the hand-crafted airlight/sky rules with a better validated sky prior, and test on "
       "more outdoor sky types;", "Sky handling: ")
bullet("revisit the local density map with a less aggressive parameter schedule, and evaluate the "
       "v2-light variant on a separate validation set rather than the test set;", "Local adaptation: ")
bullet("profile and optimise the rank-entropy local features, or replace them with a faster "
       "approximation, to reduce the 2.9\u00d7 v2 overhead;", "Runtime: ")
bullet("validate on real paired data and on video, where temporal consistency and changing "
       "illumination matter.", "Deployment: ")

body(
    "In summary, the project meets its stated goal: it builds an explainable, classical, "
    "density-aware dehazing framework, evaluates it rigorously against stronger baselines, and "
    "reports both the positive and negative findings. The result is not a claim that haze removal is "
    "solved; it is a defensible, reproducible study of when density-aware adaptation helps and when "
    "it does not."
)

# ================================================================ REFERENCES
doc.add_page_break()
heading("References", 1)
refs = [
    "K. He, J. Sun, X. Tang. Single Image Haze Removal Using Dark Channel Prior. CVPR 2009.",
    "K. He, J. Sun, X. Tang. Guided Image Filtering. ECCV 2010.",
    "K. Zuiderveld. Contrast Limited Adaptive Histogram Equalization. Graphics Gems IV, 1994.",
    "D. J. Jobson, Z. Rahman, G. A. Woodell. A Multiscale Retinex for Bridging the Gap Between Color "
    "Images and Human Vision. IEEE TIP, 1997.",
    "B. Li et al. Benchmarking Single-Image Dehazing and Beyond. IEEE TIP, 2019.",
    "Q. Zhu, J. Mai, L. Shao. A Fast Single Image Haze Removal Algorithm Using Color Attenuation "
    "Prior. IEEE TIP, 2015.",
    "G. Meng, Y. Wang, J. Duan, S. Xiang, C. Pan. Efficient Image Dehazing with Boundary Constraint "
    "and Contextual Regularization. ICCV, 2013.",
    "G. Sharma, W. Wu, E. N. Dalal. The CIEDE2000 Color-Difference Formula: Implementation Notes, "
    "Supplementary Test Data, and Mathematical Observations. Color Research & Application, 2005.",
    "A. Mittal, R. Soundararajan, A. C. Bovik. Making a \"Completely Blind\" Image Quality Analyzer. "
    "IEEE Signal Processing Letters, 2013.",
    "L. K. Choi, J. You, A. C. Bovik. Referenceless Prediction of Perceptual Fog Density and "
    "Perceptual Image Defogging. IEEE TIP, 2015.",
]
for i, ref in enumerate(refs, 1):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.first_line_indent = Cm(-0.8)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(f"[{i}]  ")
    r.bold = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = NAVY
    r2 = p.add_run(ref)
    r2.font.size = Pt(9.5)

# ================================================================ APPENDIX A
doc.add_page_break()
heading("Appendix A. Reproducibility and repository structure", 1)
body(
    "The reference implementation is Python 3.12 with NumPy, OpenCV core, SciPy, scikit-image, "
    "pandas, matplotlib and Streamlit. The presentation and report deliverables use python-pptx "
    "and python-docx. No GPU or deep-learning runtime is required. The repository is self-contained "
    "apart from the image datasets, which are downloaded by the provided scripts. The main commands "
    "below are written for cmd.exe, matching the project's run environment:"
)
code_block([
    "cd /d A:\\IPA",
    "python -m pip install -r requirements.txt",
    "",
    "REM v1 pipeline and tests",
    "powershell -ExecutionPolicy Bypass -File run_all.ps1",
    "python -m pytest",
    "",
    "REM v2 pipeline (about 35-45 minutes; overwrites v2 artifacts)",
    "powershell -ExecutionPolicy Bypass -File run_all_v2.ps1 --no-download",
    "",
    "REM live demo",
    "python -m streamlit run app\\app.py",
], size=9.0)
body(
    "The repository contains the following key artifacts:"
)
bullet("src/ \u2014 preprocessing, density, selector, DCP, guided filter, enhancement, metrics and v2 modules.",
       "Implementation: ")
bullet("config/params.yaml, config/tuned_params.json, config/calibration.json, config/tuned_params_v2.json \u2014 frozen configuration.",
       "Configuration: ")
bullet("scripts/ \u2014 evaluation, ablations, density validation, runtime profiling, report generation and this DOCX builder.",
       "Scripts: ")
bullet("results/ \u2014 per-image CSVs, summary tables, figures, key findings and key numbers.",
       "Results: ")
bullet("tests/ \u2014 83 tests, including v2 compatibility and feature tests.", "Tests: ")
bullet("app/app.py \u2014 Streamlit demo with v1/v2 toggle and density/transmission maps.", "Demo: ")
body(
    "The frozen tags are v0.1-frozen-params, v1.0, v2.0-frozen-params and v2.0. The v1 tags were "
    "never moved. With every v2 flag disabled, the output is byte-identical to v1.0. The project "
    "has been pushed to GitHub at github.com/AtulTejaswi/adaptive-fog-density-dehazing."
)
figure("app_demo.png", 16.0,
       "Streamlit demo used for the presentation. The screenshot shows the FogSense title, the density "
       "score, the selected class, the v1/v2 comparison and the v2 toggles.")

# ================================================================ APPENDIX B
heading("Appendix B. Metric definitions", 1)
body(
    "The report uses several metrics because no single metric captures dehazing quality. The "
    "definitions below are included for completeness."
)
table(
    ["Metric", "Type", "Definition / interpretation"],
    [
        ["PSNR", "Full-reference", "Peak signal-to-noise ratio against the clear SOTS ground truth; higher is better."],
        ["SSIM", "Full-reference", "Structural similarity; higher is better, sensitive to structure and contrast."],
        ["CIEDE2000", "Full-reference", "Mean colour difference in Lab; lower is better."],
        ["NIQE-style", "No-reference", "MSCN statistics compared with a Gaussian fitted on 299 calib clear images; lower is better."],
        ["FADE-style", "No-reference", "Choi et al. feature set with fixed documented weights; lower = clearer."],
        ["Entropy gain", "No-reference", "Increase in grayscale entropy after dehazing; higher = more detail."],
        ["Contrast improvement", "No-reference", "RMS contrast ratio; higher = stronger contrast."],
        ["Edge preservation", "No-reference", "Sobel edge ratio; higher = more edge detail retained."],
        ["Colour cast", "No-reference", "Deviation of channel means from gray; lower = more neutral."],
    ],
    [3.8, 3.0, 9.2],
    "Metric definitions. NIQE-style and FADE-style are proxies, not the official pretrained models.",
)
body(
    "The report uses paired Wilcoxon signed-rank tests for PSNR and SSIM comparisons and reports "
    "win rates alongside means. This avoids the common mistake of treating a small mean difference "
    "as meaningful without checking its per-image consistency."
)

# ---------------------------------------------------------------- save
# Ask Word to refresh PAGE/NUMPAGES fields when the document is opened.
settings = doc.settings.element
if settings.find(qn("w:updateFields")) is None:
    update = OxmlElement("w:updateFields")
    update.set(qn("w:val"), "true")
    settings.append(update)

doc.save(OUT)
print("wrote", OUT)
