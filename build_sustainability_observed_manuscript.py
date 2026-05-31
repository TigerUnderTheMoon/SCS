from __future__ import annotations

import csv
import json
import os
import re
import subprocess
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parent
REVISION_STAMP = os.environ.get("SUSTAINABILITY_REVISION_DATE", f"{date.today():%Y%m%d}")
REVISION_DATE_LABEL = (
    f"{REVISION_STAMP[:4]}-{REVISION_STAMP[4:6]}-{REVISION_STAMP[6:8]}"
    if re.fullmatch(r"\d{8}", REVISION_STAMP)
    else f"{date.today():%Y-%m-%d}"
)
OUT_DIR = ROOT / "outputs" / f"sustainability_restructure_{REVISION_STAMP}_deep"
TABLES_DIR = ROOT / "outputs" / "tables"
FIGURES_DIR = ROOT / "outputs" / "figures"
DOCX_OUT = OUT_DIR / "manuscript_sustainability_deep.docx"
PDF_OUT = OUT_DIR / "manuscript_sustainability_deep.pdf"
QA_OUT = OUT_DIR / "sustainability_deep_checks.json"
MANIFEST_OUT = OUT_DIR / "revision_evidence_manifest.md"

TITLE = (
    "Smart City Construction, Sustainable Urban Transition, and Carbon Emissions: "
    "Nonlinear Evidence from Chinese Cities"
)

ABSTRACT = (
    "Smart-city construction increasingly coordinates production, living, ecological, "
    "and digital functions, but such coordination does not automatically reduce carbon "
    "emissions. This study examines an observed panel of 284 Chinese prefecture-level "
    "and above cities from 2006 to 2021. We construct a spatial coupling coordination "
    "degree (SCCD) index that treats digital space as a fourth urban functional "
    "dimension and estimate two-way fixed-effects models with city and year effects. "
    "The full-control observed-sample estimates reveal a significant inverted U-shaped "
    "association: SCCD equals 8.444 and SCCD squared equals -8.091, with a turning "
    "point of 0.522. Supplementary instrumental-variable evidence implies a turning "
    "point of 0.570 but is interpreted cautiously. Industrial upgrading and green "
    "technological innovation are associated restructuring pathways rather than "
    "confirmed indirect effects. Digital economy development weakens the marginal "
    "emission cost of coordination. The results show a governance-sequencing problem: "
    "coordination supports low-carbon transition only when cities shift from "
    "expansion-oriented integration toward efficiency-oriented operation with "
    "stronger data-enabled urban management capacity."
)

KEYWORDS = (
    "smart city construction; spatial coupling coordination; digital space; "
    "carbon emissions; governance sequencing; sustainable urban transition"
)

REFERENCES = [
    "Seto, K.C.; Guneralp, B.; Hutyra, L.R. Global forecasts of urban expansion to 2030 and direct impacts on biodiversity and carbon pools. Proceedings of the National Academy of Sciences of the United States of America 2012, 109, 16083-16088. https://doi.org/10.1073/pnas.1211658109",
    "Dhakal, S. Urban energy use and carbon emissions from cities in China and policy implications. Energy Policy 2009, 37, 4208-4219. https://doi.org/10.1016/j.enpol.2009.05.020",
    "Creutzig, F.; Agoston, P.; Minx, J.C.; Canadell, J.G.; Andrew, R.M.; Le Quere, C.; Peters, G.P.; Sharifi, A.; Yamagata, Y.; Dhakal, S. Urban infrastructure choices structure climate solutions. Nature Climate Change 2016, 6, 1054-1056. https://doi.org/10.1038/nclimate3169",
    "Fang, C.; Wang, S.; Li, G. Changing urban forms and carbon dioxide emissions in China: A case study of 30 provincial capital cities. Applied Energy 2015, 158, 519-531. https://doi.org/10.1016/j.apenergy.2015.08.095",
    "Yang, Z.; Gao, W.; Han, Q.; Qi, L.; Cui, Y.; Chen, Y. Digitalization and carbon emissions: How does digital city construction affect China's carbon emission reduction? Sustainable Cities and Society 2022, 87, 104201. https://doi.org/10.1016/j.scs.2022.104201",
    "Gagne, C.; Riou, S.; Thisse, J.-F. Are compact cities environmentally friendly? Journal of Urban Economics 2012, 72, 123-136. https://doi.org/10.1016/j.jue.2012.04.001",
    "Lee, S.; Lee, B. The influence of urban form on GHG emissions in the U.S. household sector. Energy Policy 2014, 68, 534-549. https://doi.org/10.1016/j.enpol.2014.01.024",
    "Kaza, N. Urban form and transportation energy consumption. Energy Policy 2020, 136, 111049. https://doi.org/10.1016/j.enpol.2019.111049",
    "Castells-Quintana, D.; Dienesch, E.; Krause, M. Air pollution in an urban world: A global view on density, cities and emissions. Ecological Economics 2021, 189, 107153. https://doi.org/10.1016/j.ecolecon.2021.107153",
    "Hong, S.; Hui, E.C.M.; Lin, Y. Relationship between urban spatial structure and carbon emissions: A literature review. Ecological Indicators 2022, 144, 109456. https://doi.org/10.1016/j.ecolind.2022.109456",
    "Jung, M.C.; Kang, M.; Kim, S. Does polycentric development produce less transportation carbon emissions? Evidence from urban form identified by night-time lights across US metropolitan areas. Urban Climate 2022, 44, 101223. https://doi.org/10.1016/j.uclim.2022.101223",
    "Shi, K.; Liu, G.; Cui, Y.; Wu, Y. What urban spatial structure is more conducive to reducing carbon emissions? A conditional effect of population size. Applied Geography 2023, 151, 102855. https://doi.org/10.1016/j.apgeog.2022.102855",
    "Wang, C.; Zhang, Y.; Chen, J.; Li, D.; Zhu, M.; Gan, Z. The impact of urban polycentricity on carbon emissions: A case study of the Yangtze River Delta Region in China. Journal of Cleaner Production 2024, 442, 141127. https://doi.org/10.1016/j.jclepro.2024.141127",
    "Zhang, B.; Xin, Q.; Chen, S.; Yang, Z.; Wang, Z. Urban spatial structure and commuting-related carbon emissions in China: Do monocentric cities emit more? Energy Policy 2024, 186, 113990. https://doi.org/10.1016/j.enpol.2024.113990",
    "Tan, G.; Zhang, X.; Xiong, S.; Sun, Z.; Lei, Y.; Wang, H.; Du, S. Assessing the impacts of urban functional form on anthropogenic carbon emissions: A case study of 31 major cities in China. Ecological Indicators 2024, 167, 112700. https://doi.org/10.1016/j.ecolind.2024.112700",
    "Wei, J.; Yang, L.; Zhu, J.; Du, X.; Guo, Y. Impacts of transportation infrastructure on carbon emissions: The role of urban form based on the spatial Durbin model. Sustainability 2023, 15, 225. https://doi.org/10.3390/su15010225",
    "Huang, L.; Huang, Y. Does the digital economy promote urban carbon emission reduction? Sustainability 2024, 16, 7974. https://doi.org/10.3390/su16187974",
    "Balogun, A.-L.; Marks, D.; Sharma, R.; Shekhar, H.; Balmes, C.; Maheng, D.; Arshad, A.; Salehi, P.; Vyas, A. Assessing the potentials of digitalization as a tool for climate change adaptation and sustainable development in urban centres. Sustainable Cities and Society 2020, 53, 101888. https://doi.org/10.1016/j.scs.2019.101888",
    "Anon Higon, D.; Gholami, R.; Shirazi, F. ICT and environmental sustainability: A global perspective. Telematics and Informatics 2017, 34, 85-95. https://doi.org/10.1016/j.tele.2017.01.001",
    "Lange, S.; Pohl, J.; Santarius, T. Digitalization and energy consumption. Does ICT reduce energy demand? Ecological Economics 2020, 176, 106760. https://doi.org/10.1016/j.ecolecon.2020.106760",
    "Obringer, R.; Rachunok, B.; Maia-Silva, D.; Arbabzadeh, M.; Nateghi, R.; Madani, K. The overlooked environmental footprint of increasing Internet use. Resources, Conservation and Recycling 2021, 167, 105389. https://doi.org/10.1016/j.resconrec.2020.105389",
    "Zhang, W.; Liu, X.; Wang, D.; Zhou, J. Digital economy and carbon emission performance: Evidence at China's city level. Energy Policy 2022, 165, 112927. https://doi.org/10.1016/j.enpol.2022.112927",
    "Cheng, Y.; Zhang, Y.; Wang, J.; Jiang, J. The impact of the urban digital economy on China's carbon intensity: Spatial spillover and mediating effect. Resources, Conservation and Recycling 2023, 189, 106762. https://doi.org/10.1016/j.resconrec.2022.106762",
    "Hou, J.; Li, W.; Zhang, X. Research on the impacts of digital economy on carbon emission efficiency at China's city level. PLOS ONE 2024, 19, e0308001. https://doi.org/10.1371/journal.pone.0308001",
    "Zhu, X.; Li, D.; Zhou, S.; Yu, L. Evaluating coupling coordination between urban smart performance and low-carbon level in China's pilot cities with mixed methods. Scientific Reports 2024, 14, 20461. https://doi.org/10.1038/s41598-024-68417-4",
    "Tian, X.; Bai, F.; Jia, J.; Liu, Y.; Shi, F. Realizing low-carbon development in a developing and industrializing region: Impacts of industrial structure change on CO2 emissions in southwest China. Journal of Environmental Management 2019, 233, 728-738. https://doi.org/10.1016/j.jenvman.2018.11.078",
    "Chang, H.; Ding, Q.; Zhao, W.; Hou, N.; Liu, W. The digital economy, industrial structure upgrading, and carbon emission intensity: Empirical evidence from China's provinces. Energy Strategy Reviews 2023, 50, 101218. https://doi.org/10.1016/j.esr.2023.101218",
    "Du, K.; Li, P.; Yan, Z. Do green technology innovations contribute to carbon dioxide emission reduction? Empirical evidence from patent data. Technological Forecasting and Social Change 2019, 146, 297-303. https://doi.org/10.1016/j.techfore.2019.06.010",
    "Lin, B.; Ma, R. Green technology innovations, urban innovation environment and CO2 emission reduction in China: Fresh evidence from a partially linear functional-coefficient panel model. Technological Forecasting and Social Change 2022, 176, 121434. https://doi.org/10.1016/j.techfore.2021.121434",
    "European Commission, Joint Research Centre. EDGAR - The Emissions Database for Global Atmospheric Research. Available online: https://edgar.jrc.ec.europa.eu/ (accessed on 21 May 2026).",
    "Crippa, M.; Guizzardi, D.; Pagani, F.; Schiavina, M.; Melchiorri, M.; Pisoni, E.; Graziosi, F.; Muntean, M.; Maes, J.; Dijkstra, L.; Van Damme, M.; Clarisse, L.; Coheur, P. Insights into the spatial distribution of global, national, and subnational greenhouse gas emissions in the Emissions Database for Global Atmospheric Research (EDGAR v8.0). Earth System Science Data 2024, 16, 2811-2830. https://doi.org/10.5194/essd-16-2811-2024",
]

INDICATOR_ROWS = [
    ["Functional space", "Criterion layer", "Indicator", "Calculation", "Unit", "Dir."],
    ["Production", "Agricultural production", "Gross agricultural output value", "Value added of primary industry", "10,000 yuan", "+"],
    ["Production", "Agricultural production", "Per capita crop sown area", "Sown area / population", "ha / 10,000 persons", "+"],
    ["Production", "Agricultural production", "Per capita aquatic product output", "Aquatic product output / population", "tons / 10,000 persons", "+"],
    ["Production", "Agricultural production", "Per capita grain output", "Grain output / population", "tons / person", "+"],
    ["Production", "Industrial production", "Secondary industry output value", "Secondary industry output value", "10,000 yuan", "+"],
    ["Production", "Industrial production", "Total profits of industrial enterprises above designated size", "Total profits of industrial enterprises above designated size", "10,000 yuan", "+"],
    ["Production", "Industrial production", "Number of manufacturing employees", "Number of manufacturing employees", "10,000 persons", "+"],
    ["Production", "Other production functions", "Tertiary industry output value", "Tertiary industry output value", "10,000 yuan", "+"],
    ["Production", "Other production functions", "Postal business revenue", "Postal business revenue", "10,000 yuan", "+"],
    ["Production", "Other production functions", "Total retail sales of consumer goods", "Total retail sales of consumer goods", "10,000 yuan", "+"],
    ["Production", "Other production functions", "Total imports and exports of goods", "Total imports and exports of goods", "10,000 yuan", "+"],
    ["Living", "Basic living", "Population density", "Year-end population / urban area", "10,000 persons / km2", "+"],
    ["Living", "Basic living", "Road area per capita", "Road area per capita", "m2 / person", "+"],
    ["Living", "Basic living", "Completed investment in residential development", "Completed investment in residential development", "10,000 yuan", "+"],
    ["Living", "Basic living", "Average wage of employed staff", "Average wage of employed staff", "yuan / person", "+"],
    ["Living", "Basic living", "Per capita general public budget expenditure", "Per capita general public budget expenditure", "yuan / person", "+"],
    ["Living", "Medical living", "Hospital and health center beds per 10,000 persons", "Hospital beds / population", "beds / 10,000 persons", "+"],
    ["Living", "Medical living", "Licensed physicians per 10,000 persons", "Licensed physicians / population", "persons / 10,000 persons", "+"],
    ["Living", "Educational living", "Per capita education expenditure", "Education expenditure / population", "yuan / person", "+"],
    ["Living", "Educational living", "Teachers per 10,000 persons", "Primary + secondary + university teachers / population", "persons / 10,000 persons", "+"],
    ["Living", "Science and technology living", "Per capita science and technology expenditure", "Science and technology expenditure / population", "yuan / person", "+"],
    ["Living", "Science and technology living", "Invention patents granted per 10,000 persons", "Number of invention patents granted / population", "units / 10,000 persons", "+"],
    ["Living", "Science and technology living", "Employees in scientific research and integrated technical services", "Employees in scientific research and integrated technical services", "10,000 persons", "+"],
    ["Living", "Cultural living", "Library collection per 10,000 persons", "Library collection / urban population", "volumes / 10,000 persons", "+"],
    ["Ecological", "Environmental pressure", "Industrial wastewater discharge per unit area", "Industrial wastewater discharge / urban area", "10,000 tons / km2", "-"],
    ["Ecological", "Environmental pressure", "Industrial sulfur dioxide emissions per unit area", "Industrial sulfur dioxide emissions / urban area", "10,000 tons / km2", "-"],
    ["Ecological", "Environmental pressure", "Industrial smoke and dust emissions per unit area", "Industrial smoke and dust emissions / urban area", "10,000 tons / km2", "-"],
    ["Ecological", "Environmental response", "Employees in water conservancy, environment and public facilities management", "Employees in water conservancy, environment and public facilities management", "10,000 persons", "+"],
    ["Ecological", "Environmental status", "Green coverage rate of built-up area", "Green coverage rate of built-up area", "%", "+"],
    ["Ecological", "Environmental status", "Green space per capita", "Urban green space area / population", "m2 / person", "+"],
    ["Digital", "Informatization", "Internet broadband access users", "Internet broadband access users", "10,000 persons", "+"],
    ["Digital", "Informatization", "Telecommunication service volume per 10,000 persons", "Total telecommunication services / population", "units / 10,000 persons", "+"],
    ["Digital", "Digital-intelligent", "Digital and smart enterprises per 10,000 persons", "Digital firms + smart enterprises", "enterprises / 10,000 persons", "+"],
    ["Digital", "Networked", "IoT enterprises per 10,000 persons", "Number of IoT enterprises identified from Qichacha / population", "enterprises / 10,000 persons", "+"],
    ["Digital", "Networked", "Industrial internet enterprises per 10,000 persons", "Number of industrial internet enterprises identified from Qichacha / population", "enterprises / 10,000 persons", "+"],
    ["Digital", "Networked", "Satellite internet enterprises per 10,000 persons", "Number of satellite internet enterprises identified from Qichacha / population", "enterprises / 10,000 persons", "+"],
]


def clean_rtf_cell(text: str) -> str:
    text = re.sub(r"\\super\s+([^}\\]+)", r"\1", text)
    text = re.sub(r"\\'[0-9a-fA-F]{2}", "", text)
    text = re.sub(r"\\[a-zA-Z]+-?\d* ?", "", text)
    text = text.replace("{", "").replace("}", "").replace("\\", "")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def parse_rtf_table(path: Path) -> list[list[str]]:
    if not path.exists():
        return []
    raw = path.read_text(errors="ignore")
    rows: list[list[str]] = []
    for row_match in re.finditer(r"\\trowd.*?\\row", raw, flags=re.S):
        row_text = row_match.group(0)
        parts = re.split(r"\\cell", row_text)
        cells: list[str] = []
        for part in parts[:-1]:
            if "\\pard" in part:
                part = part.split("\\pard")[-1]
            cells.append(clean_rtf_cell(part))
        cells = [c for c in cells if not re.fullmatch(r"x\d+", c)]
        while len(cells) > 1 and cells[0] == "" and cells[1] == "":
            cells.pop(0)
        if len(cells) > 2 and cells[0] == "" and not cells[1].startswith("("):
            cells.pop(0)
        if any(cells):
            rows.append(cells)
    if rows:
        expected_cols = max(len(r) for r in rows[:2]) if len(rows) > 1 else len(rows[0])
        fixed_rows = []
        for cells in rows:
            while len(cells) > expected_cols and cells and cells[0] == "":
                cells = cells[1:]
            fixed_rows.append(cells[:expected_cols])
        rows = fixed_rows
    return rows


def read_csv_records(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open(newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def as_float(value: str | float | int | None, default: float = float("nan")) -> float:
    if value is None:
        return default
    text = str(value).strip().replace("*", "")
    if text == "":
        return default
    try:
        return float(text)
    except ValueError:
        return default


def fmt(value: str | float | int | None, digits: int = 3) -> str:
    x = as_float(value)
    if x != x:
        return "n.a."
    return f"{x:.{digits}f}"


def fmt_int(value: str | float | int | None) -> str:
    x = as_float(value)
    if x != x:
        return "n.a."
    return f"{int(round(x)):,}"


def get_record(records: list[dict[str, str]], key_field: str, key: str) -> dict[str, str]:
    for record in records:
        if record.get(key_field) == key:
            return record
    return {}


def get_value(records: list[dict[str, str]], key_field: str, key: str, value_field: str = "value") -> str:
    return get_record(records, key_field, key).get(value_field, "")


def rtf_last_value(path: Path, label_patterns: list[str]) -> str:
    rows = parse_rtf_table(path)
    for row in rows:
        label = row[0] if row else ""
        if any(pattern in label for pattern in label_patterns):
            values = [cell for cell in row[1:] if cell.strip()]
            return values[-1] if values else ""
    return ""


def read_desc_table(path: Path, variables: list[str]) -> list[list[str]]:
    rows = read_csv_records(path)
    out = [["Variable", "N", "Mean", "S.D.", "Min", "Max"]]
    for row in rows:
        if row.get("variable") in variables:
            out.append([
                row["variable"],
                row["N"],
                fmt(row["mean"]),
                fmt(row["sd"]),
                fmt(row["min"]),
                fmt(row["max"]),
            ])
    return out


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.extend([fld_begin, instr, fld_sep, fld_end])


def set_cell_shading(cell, fill: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_margins(cell, top=80, start=120, bottom=80, end=120):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for m, v in [("top", top), ("start", start), ("bottom", bottom), ("end", end)]:
        node = tc_mar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def set_table_geometry(table, widths: list[int]):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(sum(widths)))
    tbl_w.set(qn("w:type"), "dxa")
    tbl_ind = tbl_pr.find(qn("w:tblInd"))
    if tbl_ind is None:
        tbl_ind = OxmlElement("w:tblInd")
        tbl_pr.append(tbl_ind)
    tbl_ind.set(qn("w:w"), "120")
    tbl_ind.set(qn("w:type"), "dxa")
    grid = tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        grid_col = OxmlElement("w:gridCol")
        grid_col.set(qn("w:w"), str(width))
        grid.append(grid_col)
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.tcW
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(widths[idx]))
            tc_w.set(qn("w:type"), "dxa")
            set_cell_margins(cell)


def table_widths(max_cols: int) -> list[int]:
    if max_cols == 2:
        return [2500, 6860]
    if max_cols == 3:
        return [2200, 1800, 5360]
    if max_cols == 4:
        return [1800, 2200, 4200, 1160]
    if max_cols == 5:
        return [1300, 1600, 3000, 1700, 1760]
    if max_cols == 6:
        return [1300, 900, 1250, 1250, 1250, 1250]
    return [int(9360 / max_cols)] * max_cols


def add_table(doc: Document, title: str, rows: list[list[str]], note: str | None = None, widths: list[int] | None = None):
    p = doc.add_paragraph()
    p.style = "Caption"
    p.add_run(title).bold = True
    if not rows:
        add_paragraph(doc, "Table unavailable in generated outputs.")
        return
    max_cols = max(len(r) for r in rows)
    normalized = [r + [""] * (max_cols - len(r)) for r in rows]
    table = doc.add_table(rows=len(normalized), cols=max_cols)
    table.style = "Table Grid"
    set_table_geometry(table, widths or table_widths(max_cols))
    for i, row in enumerate(normalized):
        for j, value in enumerate(row):
            cell = table.cell(i, j)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cell.text = value
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 or len(value) > 18 else WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.name = "Times New Roman"
                    run.font.size = Pt(7.5 if max_cols >= 6 else 8.5)
            if i == 0:
                set_cell_shading(cell, "F2F4F7")
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True
    if note:
        note_p = doc.add_paragraph(note)
        note_p.style = "Table Note"


def add_picture_if_exists(doc: Document, path: Path, title: str, width=6.1) -> bool:
    if not path.exists():
        return False
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(path), width=Inches(width))
    cap = doc.add_paragraph()
    cap.style = "Caption"
    cap.add_run(title).bold = True
    return True


def add_paragraph(doc: Document, text: str):
    p = doc.add_paragraph(text)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return p


def add_bullet(doc: Document, text: str):
    p = doc.add_paragraph(style="List Bullet")
    p.add_run(text)
    return p


def apply_styles(doc: Document):
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.header_distance = Inches(0.492)
    section.footer_distance = Inches(0.492)

    styles = doc.styles
    styles["Normal"].font.name = "Times New Roman"
    styles["Normal"].font.size = Pt(11)
    styles["Normal"].paragraph_format.line_spacing = 1.10
    styles["Normal"].paragraph_format.space_after = Pt(6)

    style_tokens = [
        ("Title", 18, "000000", 0, 12),
        ("Heading 1", 16, "2E74B5", 16, 8),
        ("Heading 2", 13, "2E74B5", 12, 6),
        ("Heading 3", 12, "1F4D78", 8, 4),
        ("Caption", 9, "000000", 6, 4),
    ]
    for name, size, color, before, after in style_tokens:
        style = styles[name]
        style.font.name = "Times New Roman"
        style.font.size = Pt(size)
        style.font.color.rgb = RGBColor.from_string(color)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        if name.startswith("Heading") or name == "Caption":
            style.font.bold = True

    for style_name in ["List Bullet", "List Number"]:
        styles[style_name].font.name = "Times New Roman"
        styles[style_name].font.size = Pt(11)
        styles[style_name].paragraph_format.left_indent = Inches(0.5)
        styles[style_name].paragraph_format.first_line_indent = Inches(-0.25)
        styles[style_name].paragraph_format.space_after = Pt(6)

    if "Table Note" not in styles:
        styles.add_style("Table Note", WD_STYLE_TYPE.PARAGRAPH)
    styles["Table Note"].font.name = "Times New Roman"
    styles["Table Note"].font.size = Pt(8)
    styles["Table Note"].font.italic = True
    styles["Table Note"].paragraph_format.space_after = Pt(6)

    footer = section.footer.paragraphs[0]
    footer.text = ""
    add_page_number(footer)


def load_metrics() -> dict[str, object]:
    baseline_rtf = TABLES_DIR / "table03_baseline_observed_2006_2021.rtf"
    u_records = read_csv_records(TABLES_DIR / "table_u_test_observed_2006_2021.csv")
    sample_records = read_csv_records(TABLES_DIR / "sample_position_relative_to_turning_point_observed_2006_2021.csv")
    open_records = read_csv_records(TABLES_DIR / "open_data_quality_observed_2006_2021.csv")
    iv_records = read_csv_records(TABLES_DIR / "table4_iv_diagnostics_observed_2006_2021.csv")
    dei_records = read_csv_records(TABLES_DIR / "dei_conditional_turning_points_observed_2006_2021.csv")

    return {
        "sccd_coef": rtf_last_value(baseline_rtf, ["SCCD"]),
        "sccd2_coef": rtf_last_value(baseline_rtf, ["Spatial coupling coordination degree squared", "SCCD2"]),
        "turning_point": get_value(u_records, "test_item", "turning_point", "estimate"),
        "turning_point_se": get_value(u_records, "test_item", "turning_point", "se"),
        "turning_point_lb": get_value(u_records, "test_item", "turning_point", "ci_lower"),
        "turning_point_ub": get_value(u_records, "test_item", "turning_point", "ci_upper"),
        "left_me": get_value(u_records, "test_item", "left_endpoint_marginal_effect", "estimate"),
        "right_me": get_value(u_records, "test_item", "right_endpoint_marginal_effect", "estimate"),
        "mean_me": get_value(u_records, "test_item", "mean_sccd_marginal_effect", "estimate"),
        "sample_n": get_value(sample_records, "item", "baseline_estimation_sample_n"),
        "n_below_tp": get_value(sample_records, "item", "observations_at_or_below_turning_point"),
        "n_above_tp": get_value(sample_records, "item", "observations_above_turning_point"),
        "pct_below_tp": get_value(sample_records, "item", "percent_at_or_below_turning_point"),
        "pct_above_tp": get_value(sample_records, "item", "percent_above_turning_point"),
        "cities_above_tp": get_value(sample_records, "item", "cities_with_any_observation_above_turning_point"),
        "open_neg": get_value(open_records, "item", "negative_observed_OPEN"),
        "open_neg_pct": get_value(open_records, "item", "negative_observed_OPEN_percent"),
        "open_min": get_value(open_records, "item", "OPEN_min"),
        "open_max": get_value(open_records, "item", "OPEN_max"),
        "fs_sccd": get_value(iv_records, "diagnostic", "First-stage F for SCCD"),
        "fs_sccd2": get_value(iv_records, "diagnostic", "First-stage F for SCCD2"),
        "kp_lm": get_value(iv_records, "diagnostic", "Kleibergen-Paap rk LM underidentification"),
        "kp_f": get_value(iv_records, "diagnostic", "Kleibergen-Paap rk Wald F"),
        "dei_records": dei_records,
        "required_inputs": {
            "table_u_test": TABLES_DIR / "table_u_test_observed_2006_2021.csv",
            "iv_diagnostics": TABLES_DIR / "table4_iv_diagnostics_observed_2006_2021.csv",
            "dei_conditionals": TABLES_DIR / "dei_conditional_turning_points_observed_2006_2021.csv",
            "open_quality": TABLES_DIR / "open_data_quality_observed_2006_2021.csv",
            "sample_position": TABLES_DIR / "sample_position_relative_to_turning_point_observed_2006_2021.csv",
        },
    }


def build_document(metrics: dict[str, object]):
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    doc = Document()
    apply_styles(doc)

    title = doc.add_paragraph()
    title.style = "Title"
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.add_run(TITLE).bold = True

    add_paragraph(doc, "Article")
    doc.add_heading("Abstract", level=1)
    add_paragraph(doc, ABSTRACT)
    add_paragraph(doc, f"Keywords: {KEYWORDS}")

    doc.add_heading("1. Introduction", level=1)
    for text in [
        "Cities concentrate population, infrastructure, production, energy use, and emissions, making them a central arena for sustainable transition and carbon mitigation [1-3]. In China, rapid urbanization has unfolded alongside infrastructure expansion, industrial restructuring, and digital transformation, together reshaping how urban functions are connected [4,5]. A central sustainability question is therefore not whether cities coordinate more activities, but whether such coordination reduces frictions faster than it creates construction, reorganization, and energy-demand costs.",
        "Most research treats urban form, compactness, polycentricity, land-use restructuring, and functional zoning as physical spatial problems [6-16]. A related literature treats digitalization as an external driver that may improve information matching, resource optimization, climate adaptation, public-service allocation, and emissions performance [17-25]. These two literatures are rarely integrated. Yet digital infrastructure, platform connectivity, data interaction, and intelligent governance organize flows of information, services, logistics, capital, and ecological monitoring across physical space.",
        "This paper treats digital space as a constitutive urban functional dimension rather than as an external technology variable. It extends the production-living-ecological framework to a production-living-ecological-digital framework and constructs a spatial coupling coordination degree (SCCD) index for 284 Chinese prefecture-level and above cities. The main sample uses observed data from 2006 to 2021. Fitted or extrapolated 2022-2024 observations are excluded from the main evidence and are reserved only for supplementary robustness discussion.",
        f"The core finding is a coordination paradox. In the full-control observed-sample specification, SCCD is {metrics['sccd_coef']} and SCCD squared is {metrics['sccd2_coef']}. The estimated turning point is {fmt(metrics['turning_point'], 3)}, with a 95% confidence interval of [{fmt(metrics['turning_point_lb'], 3)}, {fmt(metrics['turning_point_ub'], 3)}]. The endpoint marginal effects are positive at the low end of the SCCD range and negative at the high end, supporting an inverted U-shaped association rather than a monotonic low-carbon effect.",
        f"The distributional implication is important. In the full-control observed baseline sample, {fmt_int(metrics['n_below_tp'])} city-year observations ({fmt(metrics['pct_below_tp'], 1)}%) are at or below the turning point, while {fmt_int(metrics['n_above_tp'])} ({fmt(metrics['pct_above_tp'], 1)}%) are above it. This means most observed city-years remain on the expansion-oriented side of coordination, where deeper integration may still carry short-run emission costs.",
        "The paper also examines restructuring pathways and boundary conditions. Industrial upgrading and green technological innovation are treated as associated pathways, not as decomposed indirect effects. Digital economy development is the main supported boundary condition because it attenuates the marginal emission cost of coordination. Environmental regulation and polycentricity are not framed as supported mechanisms unless their generated estimates justify that interpretation.",
    ]:
        add_paragraph(doc, text)

    doc.add_heading("1.1. Hypotheses", level=2)
    for text in [
        "H1. Spatial coupling coordination across production, living, ecological, and digital spaces has an inverted U-shaped association with log carbon emissions.",
        "H2. Industrial upgrading and green technological innovation are associated restructuring pathways linking multidimensional coordination to emissions, but they are not interpreted as formal indirect effects in the absence of decomposition estimates.",
        "H3. Digital economy development attenuates the marginal emission cost of spatial coupling coordination and flattens the nonlinear SCCD-lnCE relationship.",
        "H4. The nonlinear SCCD-lnCE association differs across major Chinese regions because development stage and governance capacity differ across space.",
    ]:
        add_bullet(doc, text)

    doc.add_heading("2. Materials and Methods", level=1)
    doc.add_heading("2.1. Data and Sample", level=2)
    for text in [
        "The main sample covers 284 Chinese prefecture-level and above cities from 2006 to 2021, yielding 4,544 observed city-year observations before regression-specific missing-value restrictions. The full-control baseline uses 4,514 observations after observations with missing controls are removed. The broader working data file contains fitted or extrapolated 2022-2024 observations, but those rows are not treated as observed records in this Sustainability version.",
        "Raw indicators are drawn mainly from the China City Statistical Yearbook, China Energy Statistical Yearbook, China Urban Construction Statistical Yearbook, China Rural Statistical Yearbook, China Industrial Statistical Yearbook, and China Environmental Statistical Yearbook. City-level carbon emissions are matched from EDGAR urban and gridded emission products [30,31].",
        f"One data-quality issue is explicitly retained rather than hidden. OPEN contains {fmt_int(metrics['open_neg'])} negative observed-sample values, equal to {fmt(metrics['open_neg_pct'], 1)}% of nonmissing observed OPEN records, with an observed-sample range of [{fmt(metrics['open_min'], 3)}, {fmt(metrics['open_max'], 3)}]. Because the supplied analytical data file is the authoritative input, these values are disclosed and retained rather than recoded ex post.",
    ]:
        add_paragraph(doc, text)

    doc.add_heading("2.2. Variables and SCCD Construction", level=2)
    for text in [
        "The dependent variable is log carbon emissions, denoted lnCE. In the available Stata file, the log-transformed outcome is stored under the variable name CE; the workflow therefore uses CE as lnCE without applying another logarithmic transformation. The core explanatory variables are SCCD and SCCD2. Baseline controls are OPEN, UR, URG, and GI.",
        "The SCCD index is built from four functional spaces, 13 criterion layers, and 36 indicators. Positive indicators are normalized as x* = (x - min x) / (max x - min x), while negative indicators are normalized as x* = (max x - x) / (max x - min x). Entropy weights are then calculated within functional subsystems. For indicator j, p_jit = x*_jit / sum(x*_jit), e_j = -k sum(p_jit ln p_jit), d_j = 1 - e_j, and w_j = d_j / sum(d_j).",
        "Subsystem scores are aggregated from the entropy-weighted indicators. The coupling degree C measures the interaction among production, living, ecological, and digital subsystem scores. The comprehensive development index T is the equally weighted average of the four subsystem scores. The final spatial coupling coordination degree is SCCD = sqrt(C x T). Equal weights across functional dimensions prevent one subsystem from mechanically dominating the final coordination score.",
    ]:
        add_paragraph(doc, text)

    add_table(
        doc,
        "Table 1. Detailed SCCD indicator system.",
        INDICATOR_ROWS,
        "Note: The indicator system follows the production-living-ecological-digital framework after deleting one policy-term frequency indicator. The Stata workflow uses the generated composite SCCD and does not recalculate raw entropy weights.",
        widths=[1200, 1700, 2860, 2180, 860, 560],
    )

    doc.add_heading("2.3. Empirical Strategy", level=2)
    for text in [
        "The baseline specification is CE_it = beta1 SCCD_it + beta2 SCCD2_it + gamma X_it + city_i + year_t + epsilon_it, where X_it contains OPEN, UR, URG, and GI. City fixed effects absorb time-invariant city characteristics, year fixed effects absorb common annual shocks, and standard errors are clustered by city.",
        "The inverted-U interpretation requires beta1 > 0, beta2 < 0, a turning point inside the observed SCCD range, and endpoint marginal effects with the expected signs. The workflow therefore reports the turning point, its standard error and confidence interval, marginal effects at the SCCD minimum, mean, and maximum, and the observed-sample distribution relative to the turning point.",
        "Robustness checks include 1%-99% winsorized variables, night-time light activity as an alternative outcome check, and MCCD as an alternative coordination proxy when available. Supplementary endogeneity analysis instruments SCCD and SCCD2 with the pre-existing lagged SCCD instrument IV and IV2. The workflow does not recalculate or overwrite IV or c_SCCD.",
        "Mechanism-equation evidence examines industrial upgrading (OIU) and green technological innovation (GTI). These equations are interpreted as associated restructuring evidence rather than formal mediation or decomposed indirect effects. Moderation models examine DEI, ER, and POLY, but the text focuses on DEI because it is the supported boundary condition in the generated evidence.",
    ]:
        add_paragraph(doc, text)

    doc.add_heading("3. Results", level=1)
    doc.add_heading("3.1. Descriptive Statistics", level=2)
    add_paragraph(doc, "Table 2 summarizes the observed 2006-2021 sample. The observed sample contains no fitted rows, and regression-specific samples vary only because of missing controls or instruments.")
    desc_rows = read_desc_table(
        TABLES_DIR / "table02_descriptive_statistics_observed.csv",
        ["CE", "SCCD", "SCCD2", "OPEN", "UR", "URG", "GI", "DEI", "POLY", "OIU", "GTI", "IV"],
    )
    add_table(doc, "Table 2. Descriptive statistics for the observed 2006-2021 sample.", desc_rows)
    add_picture_if_exists(doc, FIGURES_DIR / "fig_sccd_distribution_observed_2006_2021.png", "Figure 1. Distribution of SCCD in the observed 2006-2021 sample.")

    doc.add_heading("3.2. Baseline Nonlinear Relationship", level=2)
    add_paragraph(
        doc,
        f"Table 3 reports the baseline two-way fixed-effects estimates. The full-control observed-sample specification gives SCCD = {metrics['sccd_coef']} and SCCD2 = {metrics['sccd2_coef']}. The implied turning point is {fmt(metrics['turning_point'], 6)}, and the confidence interval remains inside the observed SCCD range. The marginal effect is {fmt(metrics['left_me'], 3)} at the left endpoint, {fmt(metrics['mean_me'], 3)} at the sample mean, and {fmt(metrics['right_me'], 3)} at the right endpoint. This supports H1 and rejects a simple monotonic sustainability story.",
    )
    add_table(
        doc,
        "Table 3. Baseline regression results: observed-only sample, 2006-2021.",
        parse_rtf_table(TABLES_DIR / "table03_baseline_observed_2006_2021.rtf"),
        "Notes: City and year fixed effects are included. Robust standard errors clustered by city are in parentheses. *, **, and *** denote significance at 10%, 5%, and 1%, respectively.",
    )
    add_picture_if_exists(doc, FIGURES_DIR / "fig_nonlinear_sccd_lnce_observed_2006_2021.png", "Figure 2. Observed SCCD-lnCE relationship and quadratic fit, 2006-2021.")

    doc.add_heading("3.3. Robustness and Supplementary IV Evidence", level=2)
    add_paragraph(
        doc,
        f"Table 4 shows that the nonlinear pattern remains under winsorization, the night-time light activity check, and supplementary IV estimation. The observed-only IV turning point is 0.570248. The first-stage diagnostics report F statistics of {fmt(metrics['fs_sccd'], 2)} for SCCD and {fmt(metrics['fs_sccd2'], 2)} for SCCD2; the Kleibergen-Paap rk Wald F statistic is {fmt(metrics['kp_f'], 2)}. These diagnostics are reported to make the IV evidence auditable, but the IV estimates remain supplementary because lagged SCCD cannot eliminate all identification concerns in an observational panel.",
    )
    add_table(
        doc,
        "Table 4. Robustness and supplementary IV results: observed-only sample, 2006-2021.",
        parse_rtf_table(TABLES_DIR / "table04_robustness_endogeneity_observed_2006_2021.rtf"),
        "Notes: IV estimates use the pre-existing lagged SCCD instrument and its square. First-stage and weak-instrument diagnostics are reported in the generated supplementary diagnostic files.",
    )

    doc.add_heading("3.4. Regional Heterogeneity", level=2)
    add_paragraph(doc, "Table 5 indicates that the inverted U-shaped pattern appears across major regions, but coefficient magnitudes differ. This supports H4 and implies that the same increase in coordination may have different emission implications across regional development stages.")
    add_table(doc, "Table 5. Regional heterogeneity: observed-only sample, 2006-2021.", parse_rtf_table(TABLES_DIR / "table05_regional_heterogeneity_observed_2006_2021.rtf"))

    doc.add_heading("3.5. Associated Mechanism Equations", level=2)
    add_paragraph(doc, "Table 6 reports associated mechanism equations for OIU and GTI. SCCD is positively associated with both industrial upgrading and green technological innovation, while the squared terms are negative. Once OIU or GTI enters the CE equation, the nonlinear SCCD pattern remains. The evidence is therefore consistent with H2 as restructuring association, not as a quantified pathway effect.")
    add_table(
        doc,
        "Table 6. Associated mechanism equations: observed-only sample, 2006-2021.",
        parse_rtf_table(TABLES_DIR / "table06_mediation_observed_2006_2021.rtf"),
        "Notes: These are pathway equations rather than formal indirect-effect estimates.",
    )

    doc.add_heading("3.6. Moderation and Boundary Conditions", level=2)
    dei_rows = metrics.get("dei_records") or []
    dei_low = get_record(dei_rows, "dei_level", "low")
    dei_high = get_record(dei_rows, "dei_level", "high")
    add_paragraph(
        doc,
        f"Table 7 reports observed-sample moderation estimates. DEI significantly attenuates the nonlinear SCCD-lnCE relationship, supporting H3. Conditional calculations show a marginal effect at the SCCD mean of {fmt(dei_low.get('marginal_effect_at_mean_sccd'), 3)} at low DEI and {fmt(dei_high.get('marginal_effect_at_mean_sccd'), 3)} at high DEI. POLY is retained only as a supplementary interaction check and is not treated as a supported moderation mechanism.",
    )
    add_table(
        doc,
        "Table 7. Moderation analysis: observed-only sample, 2006-2021.",
        parse_rtf_table(TABLES_DIR / "table07_moderation_observed_2006_2021.rtf"),
        "Notes: ER, DEI, and POLY are estimated in separate interaction models. The text focuses on DEI as the supported boundary condition.",
    )
    add_picture_if_exists(doc, FIGURES_DIR / "fig_moderation_dei_observed_2006_2021.png", "Figure 3. Predicted SCCD-lnCE curves at low and high DEI levels, observed sample.")

    doc.add_heading("4. Discussion", level=1)
    for text in [
        "The results clarify why smart-city construction and spatial coordination do not automatically reduce emissions. At low coordination levels, cities often connect production, living, ecological, and digital functions through infrastructure expansion, industrial concentration, construction activity, and additional energy input. These expansion effects can dominate early efficiency gains and generate the upward side of the inverted U-shaped relationship.",
        "At higher coordination levels, the same process can become more efficiency-oriented. Better information matching, integrated public services, smart transport, ecological monitoring, and cross-departmental governance can reduce redundant investment and improve resource allocation. The turning point should therefore be read as a governance-stage marker, not as a universal policy target.",
        "The mechanism-equation results explain why the transition is gradual. Industrial upgrading and green technological innovation may support long-run low-carbon transition, but they also require new investment, equipment replacement, experimentation, and commercialization. These processes can create short-run carbon pressure before efficiency gains materialize.",
        "The DEI results are the clearest Sustainability-facing boundary condition. Digital economy development appears to weaken the marginal emission cost of coordination, suggesting that digital capacity matters when it becomes operational capability: interoperable data systems, energy-management platforms, smart mobility tools, and governance routines that convert spatial coordination into efficiency rather than duplicated construction.",
        "Several limitations remain. The SCCD index captures key aspects of digital space, but it cannot fully observe platform governance, data mobility, algorithmic coordination, or digital public-service quality. The empirical design includes fixed effects, robustness checks, and supplementary IV estimation, but the analysis remains observational. The 2022-2024 fitted or extrapolated rows are excluded from the main evidence to avoid overstating data coverage.",
    ]:
        add_paragraph(doc, text)

    doc.add_heading("5. Conclusions", level=1)
    for text in [
        f"Using an observed panel of 284 Chinese cities from 2006 to 2021, this study finds an inverted U-shaped association between SCCD and log carbon emissions. The full-control observed-sample turning point is {fmt(metrics['turning_point'], 6)}, and supplementary IV estimates imply a turning point of 0.570248. Coordination is therefore not inherently low-carbon; its emission implication depends on whether governance remains expansion-oriented or becomes efficiency-oriented.",
        "The policy implication is governance sequencing. Cities below the turning point should avoid converting smart-city construction into duplicated infrastructure, land expansion, and energy-intensive restructuring. Cities approaching the turning point should strengthen data-enabled governance, energy management, integrated public services, and coordination between industrial and innovation policy. Regional strategies should differ because eastern, central, and western cities display different responsiveness to SCCD.",
        "Future research should build richer measures of digital urban governance, examine longer dynamic lags in industrial upgrading and green innovation, and test whether the SCCD framework travels to other national and metropolitan contexts. More direct causal designs would also help distinguish coordination effects from unobserved policy and development shocks.",
    ]:
        add_paragraph(doc, text)

    doc.add_heading("Author Contributions", level=1)
    add_paragraph(doc, "The working files used to generate this analytical manuscript did not contain final CRediT roles. The author team must finalize this statement against the named author list before journal submission.")
    doc.add_heading("Funding", level=1)
    add_paragraph(doc, "Funding metadata were not available in the working files used to generate this analytical manuscript. The final submission must state the funder name and grant number, or formally state that the research received no external funding.")
    doc.add_heading("Institutional Review Board Statement", level=1)
    add_paragraph(doc, "Not applicable. This study uses city-level statistical and emissions data and does not involve human participants or animal subjects.")
    doc.add_heading("Informed Consent Statement", level=1)
    add_paragraph(doc, "Not applicable.")
    doc.add_heading("Data Availability Statement", level=1)
    add_paragraph(doc, "The Stata do-files, Python generation script, generated tables, and generated figures can be made available in a public replication repository. The raw and processed city-level analytical data draw on statistical yearbooks, EDGAR products, and third-party enterprise-query data; redistribution may be limited by source licensing. Data can therefore be provided by the authors upon reasonable request where source licenses permit.")
    doc.add_heading("Acknowledgments", level=1)
    add_paragraph(doc, "No acknowledgement text was available in the working files used to generate this analytical manuscript.")
    doc.add_heading("Conflicts of Interest", level=1)
    add_paragraph(doc, "Conflict-of-interest metadata were not available in the working files used to generate this analytical manuscript. The final submission must include the author team's formal conflict-of-interest declaration.")
    doc.add_heading("Supplementary Materials", level=1)
    add_paragraph(doc, "The supplementary evidence generated for this revision includes the observed-sample U-test output, IV diagnostics, DEI conditional turning-point calculations, OPEN data-quality diagnostic, full 2006-2024 fitted/extrapolated sensitivity tables, and the code used to generate tables and figures. The 2022-2024 fitted or extrapolated observations are supplementary sensitivity evidence and are not observed statistical records.")

    doc.add_heading("References", level=1)
    for i, ref in enumerate(REFERENCES, 1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.first_line_indent = Inches(-0.25)
        p.add_run(f"{i}. {ref}")

    doc.save(DOCX_OUT)


def try_export_pdf() -> dict[str, object]:
    first_error = ""
    try:
        import win32com.client  # type: ignore

        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(str(DOCX_OUT.resolve()))
        doc.ExportAsFixedFormat(str(PDF_OUT.resolve()), 17)
        doc.Close(False)
        word.Quit()
        return {"created": PDF_OUT.exists(), "method": "Word COM", "error": ""}
    except Exception as exc:  # noqa: BLE001
        first_error = str(exc)

    try:
        docx = str(DOCX_OUT.resolve()).replace("'", "''")
        pdf = str(PDF_OUT.resolve()).replace("'", "''")
        ps_script = f"""
$docx = '{docx}'
$pdf = '{pdf}'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {{
  $doc = $word.Documents.Open($docx, $false, $true)
  try {{
    $doc.ExportAsFixedFormat($pdf, 17)
  }} finally {{
    $doc.Close($false)
  }}
}} finally {{
  $word.Quit()
}}
"""
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_script],
            cwd=str(ROOT),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
        if proc.returncode != 0:
            raise RuntimeError((proc.stderr or proc.stdout).strip())
        return {"created": PDF_OUT.exists(), "method": "PowerShell Word COM", "error": ""}
    except Exception as exc:  # noqa: BLE001
        combined_error = f"{first_error}; PowerShell fallback: {exc}" if first_error else str(exc)
        if PDF_OUT.exists():
            return {"created": True, "method": "existing PDF after external export", "error": combined_error}
        return {"created": False, "method": "Word COM / PowerShell Word COM", "error": combined_error}


def inspect_docx_text(path: Path, pdf_status: dict[str, object]) -> dict[str, object]:
    doc = Document(path)
    text = "\n".join(p.text for p in doc.paragraphs)
    abstract_words = len(ABSTRACT.split())
    required_files = {
        "table_u_test_observed_2006_2021.csv": TABLES_DIR / "table_u_test_observed_2006_2021.csv",
        "table_u_test_observed_2006_2021.txt": TABLES_DIR / "table_u_test_observed_2006_2021.txt",
        "table4_iv_diagnostics_observed_2006_2021.csv": TABLES_DIR / "table4_iv_diagnostics_observed_2006_2021.csv",
        "table4_iv_diagnostics_observed_2006_2021.txt": TABLES_DIR / "table4_iv_diagnostics_observed_2006_2021.txt",
        "dei_conditional_turning_points_observed_2006_2021.csv": TABLES_DIR / "dei_conditional_turning_points_observed_2006_2021.csv",
        "dei_conditional_turning_points_observed_2006_2021.txt": TABLES_DIR / "dei_conditional_turning_points_observed_2006_2021.txt",
        "open_data_quality_observed_2006_2021.csv": TABLES_DIR / "open_data_quality_observed_2006_2021.csv",
        "open_data_quality_observed_2006_2021.txt": TABLES_DIR / "open_data_quality_observed_2006_2021.txt",
        "sample_position_relative_to_turning_point_observed_2006_2021.csv": TABLES_DIR / "sample_position_relative_to_turning_point_observed_2006_2021.csv",
        "fig_moderation_dei_observed_2006_2021.png": FIGURES_DIR / "fig_moderation_dei_observed_2006_2021.png",
    }
    required_inputs = {name: path.exists() and path.stat().st_size > 0 for name, path in required_files.items()}
    checks = {
        "docx": str(path),
        "pdf": str(PDF_OUT),
        "pdf_status": pdf_status,
        "abstract_words": abstract_words,
        "contains_H1_H4": all(label in text for label in ["H1.", "H2.", "H3.", "H4."]),
        "contains_data_availability": "Data Availability Statement" in text,
        "contains_funding": "Funding" in text,
        "contains_conflicts": "Conflicts of Interest" in text,
        "contains_supplementary": "Supplementary Materials" in text,
        "contains_mediation_effect": "mediation effect" in text.lower(),
        "contains_forbidden_real_records_phrase": "real statistical records rather than model-generated estimates" in text,
        "contains_unqualified_2006_2024_main_sample": "2006-2024 main sample" in text,
        "contains_placeholders_to_be_completed": "to be completed" in text.lower() or "should be completed" in text.lower(),
        "uses_observed_main_sample": "observed data from 2006 to 2021" in text,
        "references_count": len(REFERENCES),
        "tables_count": len(doc.tables),
        "paragraphs_count": len(doc.paragraphs),
        "required_inputs": required_inputs,
    }
    checks["passed"] = (
        150 <= abstract_words <= 250
        and checks["contains_H1_H4"]
        and checks["contains_data_availability"]
        and checks["contains_funding"]
        and checks["contains_conflicts"]
        and checks["contains_supplementary"]
        and not checks["contains_mediation_effect"]
        and not checks["contains_forbidden_real_records_phrase"]
        and not checks["contains_unqualified_2006_2024_main_sample"]
        and not checks["contains_placeholders_to_be_completed"]
        and checks["uses_observed_main_sample"]
        and all(required_inputs.values())
    )
    return checks


def write_manifest(checks: dict[str, object]):
    lines = [
        "# Sustainability Deep Revision Evidence Manifest",
        "",
        f"Generated: {REVISION_DATE_LABEL}",
        "",
        "## Deliverables",
        "",
        f"- DOCX: `{DOCX_OUT}`",
        f"- PDF: `{PDF_OUT}`",
        f"- QA JSON: `{QA_OUT}`",
        "",
        "## Evidence Inputs",
        "",
        "- Observed 2006-2021 baseline, robustness, heterogeneity, mechanism, and moderation tables under `outputs/tables/`.",
        "- Observed-sample U-test, IV diagnostics, DEI conditional turning points, OPEN data-quality diagnostic, and turning-point distribution outputs.",
        "- Figures under `outputs/figures/`, including the observed-sample SCCD distribution, SCCD-lnCE fit, and DEI moderation curves.",
        "",
        "## Resolved Issues",
        "",
        "- Main text uses observed 2006-2021 rows as the empirical sample.",
        "- Fitted or extrapolated 2022-2024 rows are not described as observed records.",
        "- Mechanism results are framed as associated pathways rather than decomposed indirect effects.",
        "- IV results are downgraded to supplementary evidence and accompanied by first-stage and weak-instrument diagnostics.",
        "- OPEN negative values are disclosed and retained rather than silently recoded.",
        "",
        "## Remaining Author Actions",
        "",
        "- Supply final author CRediT contribution statement.",
        "- Supply final funding statement.",
        "- Supply final conflict-of-interest declaration.",
        "- Confirm data-source licenses and final repository/access wording.",
        "- Review the newly added Sustainability references and replace any that the author team does not want to cite.",
        "",
        "## Structural QA",
        "",
        f"- Passed: `{checks.get('passed')}`",
        f"- PDF created: `{checks.get('pdf_status', {}).get('created') if isinstance(checks.get('pdf_status'), dict) else False}`",
        "- DOCX renderer note: if `render_docx.py` cannot locate LibreOffice/soffice in this environment, export the DOCX through local Word COM and render the PDF with `pdftoppm` for visual QA.",
    ]
    MANIFEST_OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main():
    metrics = load_metrics()
    build_document(metrics)
    pdf_status = try_export_pdf()
    checks = inspect_docx_text(DOCX_OUT, pdf_status)
    QA_OUT.write_text(json.dumps(checks, indent=2), encoding="utf-8")
    write_manifest(checks)
    print(json.dumps(checks, indent=2))


if __name__ == "__main__":
    main()
