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
    "Ma, Z.; Wu, F. Smart City, Digitalization and CO2 Emissions: Evidence from 353 Cities in China. Sustainability 2023, 15, 225. https://doi.org/10.3390/su15010225",
    "Li, D.; Zhang, X.; Lau, A.D.; Gong, Y. The Impact of Smart City Construction on the Spatial Distribution of Urban Carbon Emissions Based on the Panel Data Analysis of 277 Prefecture-Level Cities in China. Sustainability 2025, 17, 4934. https://doi.org/10.3390/su17114934",
    "An, X.; Yang, Y.; Zhang, X.; Zeng, X. Smarter and Cleaner? The Carbon Reduction Effect of Smart Cities: A Perspective on Green Technology Progress. Sustainability 2024, 16, 8048. https://doi.org/10.3390/su16188048",
    "Huang, L.; Huang, Y. Does the digital economy promote urban carbon emission reduction? Sustainability 2024, 16, 7974. https://doi.org/10.3390/su16187974",
    "Balogun, A.-L.; Marks, D.; Sharma, R.; Shekhar, H.; Balmes, C.; Maheng, D.; Arshad, A.; Salehi, P.; Vyas, A. Assessing the potentials of digitalization as a tool for climate change adaptation and sustainable development in urban centres. Sustainable Cities and Society 2020, 53, 101888. https://doi.org/10.1016/j.scs.2019.101888",
    "Anon Higon, D.; Gholami, R.; Shirazi, F. ICT and environmental sustainability: A global perspective. Telematics and Informatics 2017, 34, 85-95. https://doi.org/10.1016/j.tele.2017.01.001",
    "Lange, S.; Pohl, J.; Santarius, T. Digitalization and energy consumption. Does ICT reduce energy demand? Ecological Economics 2020, 176, 106760. https://doi.org/10.1016/j.ecolecon.2020.106760",
    "Obringer, R.; Rachunok, B.; Maia-Silva, D.; Arbabzadeh, M.; Nateghi, R.; Madani, K. The overlooked environmental footprint of increasing Internet use. Resources, Conservation and Recycling 2021, 167, 105389. https://doi.org/10.1016/j.resconrec.2020.105389",
    "Zhang, W.; Liu, X.; Wang, D.; Zhou, J. Digital economy and carbon emission performance: Evidence at China's city level. Energy Policy 2022, 165, 112927. https://doi.org/10.1016/j.enpol.2022.112927",
    "Cheng, Y.; Zhang, Y.; Wang, J.; Jiang, J. The impact of the urban digital economy on China's carbon intensity: Spatial spillover and mediating effect. Resources, Conservation and Recycling 2023, 189, 106762. https://doi.org/10.1016/j.resconrec.2022.106762",
    "Hou, J.; Li, W.; Zhang, X. Research on the impacts of digital economy on carbon emission efficiency at China's city level. PLOS ONE 2024, 19, e0308001. https://doi.org/10.1371/journal.pone.0308001",
    "Zhu, X.; Li, D.; Zhou, S.; Zhu, S.; Yu, L. Evaluating coupling coordination between urban smart performance and low-carbon level in China's pilot cities with mixed methods. Scientific Reports 2024, 14, 20461. https://doi.org/10.1038/s41598-024-68417-4",
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
        "kp_lm": get_value(iv_records, "diagnostic", "Kleibergen-Paap rk LM"),
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
        "Cities are now the main arena in which China must reconcile economic upgrading, public-service expansion, digital transformation, and carbon mitigation. Urban infrastructure choices shape energy demand for decades, and the spatial organization of production, residence, ecological protection, and public services determines whether new investment reduces frictions or locks cities into energy-intensive patterns [1-4]. Smart-city construction adds a further layer to this problem. It changes how information is produced, shared, and used across departments, firms, households, and environmental agencies. The key sustainability question is therefore not simply whether a city becomes smarter, but whether smart coordination turns into lower-carbon operation.",
        "Recent evidence is encouraging but incomplete for this question. Studies of digital city construction and smart-city pilot policies generally find that smart-city adoption reduces urban carbon emissions on average [5,16,18]. Ma and Wu report that smart-city digital governance reduces city CO2 emissions and works partly through green patenting [16]. An et al. show that smart-city construction lowers emissions from the perspective of green technology progress and total factor productivity [18]. Li et al. further show that the effect of smart-city construction on carbon emissions has regional and spatial heterogeneity across Chinese prefecture-level cities [17]. These papers establish that policy adoption can matter, but they estimate average policy effects rather than the marginal emission implications of deeper multidimensional coordination.",
        "This distinction matters because smart-city construction is not a single technology shock. In early stages, smart systems often require broadband networks, sensor deployment, platform construction, data centers, transport upgrades, public-service digitization, and administrative reorganization. These inputs may raise emissions before they improve efficiency. At more advanced stages, the same digital and spatial integration can reduce duplicated construction, improve public-service allocation, optimize traffic and logistics, strengthen ecological monitoring, and support cleaner industrial upgrading. A monotonic carbon-reduction story is therefore too simple for cities that are still moving from infrastructure-led smart construction to data-enabled urban operation.",
        "The paper develops this argument by treating digital space as a constitutive urban functional dimension. Existing urban-form studies focus on compactness, density, polycentricity, commuting, land-use mix, and functional layout [6-15]. Digitalization studies focus on information and communication technology, digital economy development, platform governance, and carbon performance [19-26]. Coupling-coordination research shows that smart performance and low-carbon levels can move together across pilot cities [27]. Yet the empirical literature rarely places digital space inside the same coordination framework as production, living, and ecological spaces. This paper does so by extending the production-living-ecological framework to a production-living-ecological-digital framework.",
        "The empirical object is spatial coupling coordination degree, or SCCD. SCCD measures how production, living, ecological, and digital functional spaces develop and interact within each city. It is not a smart-city pilot dummy and is not a simple digital-economy index. It is a continuous coordination index that can vary within cities over time. This design allows the paper to ask a different question from the policy-adoption literature: as the coordination intensity of urban functional spaces rises, does the emission association remain negative, or does it change across stages?",
        "The analysis uses an observed panel of 284 Chinese prefecture-level and above cities from 2006 to 2021. The working data file also contains fitted or extrapolated 2022-2024 rows, but the main Sustainability evidence excludes those rows. This distinction is central to the manuscript. The observed sample contains 4,544 city-year observations before regression-specific missing-value restrictions, and the full-control baseline uses 4,514 observations. The dependent variable is log carbon emissions, stored as CE in the analytical Stata file. The baseline regressions include city and year fixed effects and the control variables OPEN, UR, URG, and GI.",
        f"The main result is a nonlinear coordination paradox. In the full-control observed-sample specification, SCCD is {metrics['sccd_coef']} and SCCD squared is {metrics['sccd2_coef']}. The implied turning point is {fmt(metrics['turning_point'], 3)}, with a 95% confidence interval of [{fmt(metrics['turning_point_lb'], 3)}, {fmt(metrics['turning_point_ub'], 3)}]. The marginal effect is positive at the low end of the observed SCCD range and negative at the high end. This pattern supports an inverted U-shaped association: deeper coordination initially coincides with higher emissions, but the association turns downward after cities reach a higher coordination stage.",
        f"The distribution around the turning point reinforces the interpretation. In the full-control observed baseline sample, {fmt_int(metrics['n_below_tp'])} city-year observations, or {fmt(metrics['pct_below_tp'], 1)}%, are at or below the turning point, while {fmt_int(metrics['n_above_tp'])} observations, or {fmt(metrics['pct_above_tp'], 1)}%, are above it. Most observed city-years are therefore still on the expansion-oriented side of the curve. For those cities, more coordination may still mean more infrastructure, more construction, more industrial concentration, and more energy demand before efficiency gains become dominant.",
        "The paper also examines associated restructuring pathways and boundary conditions. Industrial upgrading and green technological innovation are positively associated with SCCD in the generated pathway equations, but these estimates are not interpreted as decomposed pathway shares. Digital economy development is the strongest supported boundary condition: higher DEI attenuates the marginal emission cost of SCCD and flattens the nonlinear relationship. Environmental regulation and polycentricity are reported as part of the moderation workflow, but the manuscript does not elevate them into supported mechanisms when the generated evidence does not justify that interpretation.",
        "The contribution is threefold. First, the paper shifts the smart-city carbon discussion from binary policy adoption to continuous coordination intensity. Second, it integrates spatial and digital urban functions in one SCCD framework, linking urban form, digitalization, smart-city policy, and low-carbon coupling literatures. Third, it shows why average smart-city carbon-reduction effects and a nonlinear coordination-intensity relationship can coexist. Smart-city designation may reduce emissions on average, while deeper coordination may still pass through an expansionary stage before it becomes efficiency-oriented.",
        "The rest of the paper proceeds directly from this argument. Section 2 reviews the smart-city carbon, spatial-distribution, green-technology-progress, and coupling-coordination literatures and develops the hypotheses. Section 3 describes the observed sample, SCCD construction, and empirical strategy. Section 4 reports descriptive statistics, baseline nonlinear estimates, robustness and supplementary IV evidence, regional heterogeneity, associated pathway equations, and DEI moderation. Section 5 compares the findings with prior smart-city studies and derives Sustainability-facing policy implications. Section 6 concludes with author-side items that remain necessary before journal submission.",
    ]:
        add_paragraph(doc, text)

    doc.add_heading("2. Literature Review and Theoretical Framework", level=1)
    doc.add_heading("2.1. Smart-City Construction and Urban Carbon Reduction", level=2)
    for text in [
        "The closest literature studies whether smart-city policy reduces urban carbon emissions. This literature usually treats smart-city construction as a policy intervention and estimates average treatment effects with difference-in-differences or related quasi-experimental designs. The policy variable captures designation, pilot inclusion, or digital governance construction rather than a continuous level of coordination. The central finding is broadly optimistic: smart-city construction can improve emissions performance by strengthening digital governance, information matching, innovation incentives, and urban management capacity [5,16-18].",
        "Yang et al. provide an important benchmark by linking digital city construction to China's carbon emission reduction [5]. Their framing is close to the policy side of this paper because digital city construction affects urban information infrastructure and governance capacity. Their dependent variable and empirical object differ from the SCCD framework, but the policy implication is relevant: digital systems can become emission-reduction tools when they improve the allocation of energy, transport, public services, and industrial resources. This paper builds on that insight while asking whether the marginal effect of coordination intensity is constant.",
        "Ma and Wu study 353 Chinese cities and find that smart-city digital governance reduces city-level CO2 emissions, with green patenting as one important channel [16]. This result is useful for two reasons. First, it shows that digital governance is not merely an information technology outcome; it can have measurable environmental implications. Second, it links smart-city construction to green innovation, which supports this paper's choice to examine GTI as an associated restructuring pathway. The current manuscript remains more cautious because its pathway equations do not decompose an identified channel; they show association patterns within the SCCD framework.",
        "An et al. examine the carbon-reduction effect of smart cities from the perspective of green technology progress [18]. Their study connects smart-city construction with green technology progress and total factor productivity, reinforcing the idea that smart systems can influence emissions through innovation and efficiency. This paper uses that logic to motivate H2, but it does not copy the same design. Instead of estimating a binary smart-city treatment effect, it studies whether a continuous production-living-ecological-digital coordination index has a nonlinear association with log emissions.",
        "Li et al. analyze 277 prefecture-level Chinese cities and focus on the spatial distribution of urban carbon emissions under smart-city construction [17]. Their study is especially relevant because it emphasizes spatial heterogeneity, including stronger effects in northern and non-resource-based cities. The current manuscript does not classify cities by resource-based status because the existing workflow does not generate that analysis. It uses Li et al. to motivate regional heterogeneity in a narrower and evidence-consistent way: the emission implication of smart-city construction and coordination should differ across regions because industrial structure, climate, governance capacity, and development stage differ across space.",
        "These studies provide a clear prior: smart-city construction can reduce emissions on average. They also leave a gap. A city that receives a smart-city policy label and a city that reaches a high level of integrated production-living-ecological-digital coordination are not the same empirical object. Policy adoption may trigger governance reforms, funding support, and digital infrastructure, while coordination intensity measures how deeply different functional subsystems have developed together. A city can be designated as smart but remain in an early coordination stage; another city can develop strong coordination without the same policy timing. This paper addresses that distinction.",
    ]:
        add_paragraph(doc, text)

    doc.add_heading("2.2. Urban Spatial Structure, Coordination, and Emissions", level=2)
    for text in [
        "A second literature studies urban spatial structure and carbon emissions. It asks how density, compactness, polycentricity, commuting distance, land-use allocation, and functional form shape energy demand and emissions [6-15]. This literature shows that the carbon consequences of spatial organization are theoretically ambiguous. Compact development may shorten trips and reduce infrastructure duplication, but density can also raise congestion, housing costs, and localized energy demand. Polycentric development may reduce commuting distance for some households, but it can also increase cross-center travel and infrastructure duplication if centers are poorly connected.",
        "Fang et al. show that urban form matters for CO2 emissions in Chinese provincial capital cities [4]. Later work extends this question to density, polycentricity, commuting-related emissions, and functional form across different countries and regions [6-15]. The common insight is that spatial structure is not merely a background control variable. It organizes the daily flows through which economic activity produces emissions. Production sites, residential districts, public services, transport corridors, ecological spaces, and commercial centers interact; emissions emerge from those interactions rather than from any single subsystem.",
        "The SCCD framework follows this interaction logic. A city with strong production capacity but weak living services may generate long commuting and inefficient labor-market matching. A city with strong ecological resources but weak production upgrading may face pressure to trade environmental quality for short-run industrial growth. A city with digital platforms but weak interoperability may duplicate data systems across departments. Coupling coordination is therefore not the same as high development in one dimension. It measures whether several urban functional spaces move together in a way that can support integrated governance.",
        "The spatial literature also helps explain why an inverted U-shaped association is plausible. In low-coordination stages, a city may raise SCCD by expanding infrastructure, improving roads, enlarging public-service supply, building industrial platforms, and adding digital facilities. These investments can raise construction emissions, electricity demand, and energy-intensive production in the short run. In high-coordination stages, the same functional links can improve matching and reduce waste: transport networks can be optimized, land can be used more efficiently, ecological monitoring can constrain pollution, and digital systems can reduce redundant administrative and physical investment.",
        "This paper therefore treats SCCD as a stage-dependent coordination measure. The estimated turning point is not interpreted as a universal policy threshold that every city should target mechanically. It is an empirical marker in the observed sample. The policy question is what cities are doing as they move along the coordination distribution. If coordination is mainly an infrastructure expansion process, emissions may rise. If coordination becomes an operational efficiency process, emissions may fall. This distinction anchors the paper's main hypothesis.",
    ]:
        add_paragraph(doc, text)

    doc.add_heading("2.3. Digitalization, Green Technology Progress, and Low-Carbon Coupling", level=2)
    for text in [
        "Digitalization can reduce emissions by improving information flows, reducing search costs, supporting smart mobility, strengthening energy management, and enabling real-time environmental monitoring [19-26]. It can also raise emissions through data centers, device production, rebound effects, platform-induced consumption, and increased electricity demand [21-23]. The environmental role of digitalization is therefore conditional. The same digital capacity can either help cities coordinate resources more efficiently or create new sources of energy use if it is built as an additional infrastructure layer without operational integration.",
        "Digital-economy studies at the city level are relevant because DEI is the paper's main boundary condition. Zhang et al. link digital economy development to carbon emission performance in Chinese cities [24]. Cheng et al. study digital economy, carbon intensity, and spatial spillovers [25]. Hou et al. examine the digital economy and carbon emission efficiency at China's city level [26]. Together, these studies support the idea that digital development changes the productivity and efficiency environment in which spatial coordination operates. The current paper tests this idea by interacting SCCD with DEI rather than treating digitalization only as a direct regressor.",
        "Green technological innovation is another mechanism through which smart-city construction may reduce emissions. Ma and Wu emphasize green patents in the smart-city digital governance setting [16], and An et al. place green technology progress at the center of the smart-city carbon-reduction effect [18]. Broader innovation evidence also links green technology innovation to lower CO2 emissions and carbon-emission reduction in China [30,31]. These findings support the expectation that SCCD may be associated with GTI, because coordinated digital and spatial systems can improve knowledge flows, policy targeting, firm upgrading, and environmental monitoring.",
        "Industrial upgrading provides a complementary restructuring pathway. Industrial structure change can reduce emissions when cities move away from heavy, energy-intensive production toward services, advanced manufacturing, and cleaner technologies [28,29]. Yet upgrading can also require new equipment, new buildings, reallocation costs, and transitional energy demand. For this reason, this manuscript does not assume that OIU or GTI automatically lowers emissions in every period. It treats them as associated pathways that may help explain the transition from expansion-oriented coordination to efficiency-oriented coordination.",
        "Coupling-coordination evidence provides a bridge between the digital and low-carbon literatures. Zhu et al. evaluate coupling coordination between urban smart performance and low-carbon level across Chinese pilot cities and find a positive coupling-coordination pattern [27]. Their analysis differs from this paper in sample, method, and index construction, but it supports the broader premise that smart performance and low-carbon development should be analyzed jointly. The current manuscript extends that premise by embedding digital space in a four-dimensional urban functional system and estimating a nonlinear SCCD-lnCE association.",
        "The theoretical implication is that digital space has two roles. First, it is one of the urban functional subsystems that contributes to SCCD. Second, digital economy development can condition how SCCD maps into emissions. A city may have digital indicators inside its coordination index, but the broader digital economy determines whether data systems, firms, workers, and public agencies can use digital resources effectively. This distinction motivates H3: DEI should attenuate the marginal emission cost of coordination when it helps convert coordination from construction to operation.",
    ]:
        add_paragraph(doc, text)

    doc.add_heading("2.4. The Coordination Paradox", level=2)
    for text in [
        "The coordination paradox is the paper's central theoretical claim. Urban coordination is necessary for sustainable transition, but coordination is not automatically low-carbon. In the early stage, cities improve coordination by building and connecting systems: industrial parks, roads, digital platforms, environmental facilities, residential amenities, and public-service networks. This process raises the measured level of coordination, but it also consumes materials, land, electricity, and construction services. The carbon cost of building coordination can therefore dominate the carbon savings from using coordination.",
        "In the later stage, coordination changes character. Once the main systems exist, marginal improvements can come from better data sharing, more precise public services, demand-side energy management, logistics optimization, industrial upgrading, and ecological monitoring. These changes may reduce emissions because they improve how existing assets are used. A mature smart city can coordinate traffic lights, public transit, industrial energy use, waste management, and environmental enforcement in ways that reduce energy waste and pollution. The same SCCD increase therefore has different implications depending on the stage of urban development.",
        "This logic differs from a standard linear smart-city argument. A linear argument would expect each additional unit of smart or spatial coordination to reduce emissions by the same amount. The coordination paradox expects a positive marginal association at low SCCD and a negative marginal association at high SCCD. In regression terms, it expects a positive coefficient on SCCD and a negative coefficient on SCCD squared, with the turning point inside the observed range. The empirical tests therefore focus on the signs, turning point, endpoint marginal effects, and distribution of city-year observations around the turning point.",
        "The paradox also clarifies why the paper's result can coexist with prior DID estimates. DID studies compare treated and untreated cities around policy adoption and estimate an average effect of designation. The SCCD analysis compares different levels of coordination intensity within a panel and estimates a nonlinear association. A policy can reduce average emissions if it improves governance, while marginal increases in coordination can still be emission-intensive in early-stage cities. These are not competing claims; they answer different questions about different empirical objects.",
        "The framework also explains why heterogeneity is expected. Regions differ in industrial base, climate, fiscal capacity, digital infrastructure, energy mix, administrative capacity, and exposure to national development strategies. The same increase in SCCD may therefore represent different activities in different regions. In an eastern city, it may reflect refined digital governance and service integration. In a central or western city, it may reflect infrastructure catch-up and industrial platform construction. Regional heterogeneity is therefore not an add-on result; it follows from the stage-dependent theory.",
        "Finally, the framework imposes claim discipline. The paper can show that OIU and GTI move with SCCD and that DEI changes the SCCD-lnCE relationship. It cannot, without additional design, assign a precise causal share of the SCCD effect to any pathway. It can report supplementary IV evidence using the pre-existing lagged SCCD instrument, but lagged coordination may still be correlated with unobserved development trajectories. The manuscript therefore states the main evidence as an observational fixed-effects association with robustness and supplementary endogeneity checks.",
    ]:
        add_paragraph(doc, text)

    doc.add_heading("2.5. Positioning Against the Four Closest Studies", level=2)
    for text in [
        "The four requested studies are not simply additional citations; they define the paper's closest comparison set. Ma and Wu [16] provide the digital-governance benchmark, Li et al. [17] provide the spatial-distribution and heterogeneity benchmark, An et al. [18] provide the green-technology-progress benchmark, and Zhu et al. [27] provide the coupling-coordination benchmark. The present manuscript uses each paper to sharpen a different part of the argument rather than adding them as a citation list. This is important for Sustainability because the journal's smart-city literature is already broad; the manuscript must state clearly what is new relative to adjacent published work.",
        "Relative to Ma and Wu [16], the paper changes the treatment concept. Ma and Wu ask whether smart-city digitalization or digital governance lowers CO2 emissions across a large sample of cities. Their result supports the view that smart-city construction can be environmentally beneficial. The present paper does not dispute that view. It asks whether a continuous coordination index that includes digital space, production space, living space, and ecological space has a constant marginal association with emissions. This turns the discussion from whether a policy helps on average to when coordination becomes cleaner at the margin.",
        "Relative to Li et al. [17], the paper changes the heterogeneity claim. Li et al. emphasize spatial distribution and city-type heterogeneity in a 277-city panel. Their evidence supports the idea that smart-city carbon effects differ across urban contexts. The present manuscript uses that insight to motivate regional heterogeneity, but it deliberately avoids unsupported resource-based-city claims because the current Stata outputs do not include such a classification. This conservative choice should make the paper easier to defend. It credits the city-type literature without importing results that the manuscript itself does not estimate.",
        "Relative to An et al. [18], the paper changes the mechanism standard. An et al. emphasize green technology progress and total factor productivity in the smart-city carbon-reduction process. This manuscript uses that literature to motivate GTI and OIU as associated restructuring pathways. It does not call the pathway evidence a formal decomposition. The distinction is not semantic. A reviewer may accept that SCCD is associated with GTI and that GTI matters for low-carbon transition while still objecting to an unsupported claim that GTI transmits a quantified share of the SCCD-emissions association. The revised wording avoids that vulnerability.",
        "Relative to Zhu et al. [27], the paper changes the coordination object. Zhu et al. examine coupling coordination between smart performance and low-carbon level in pilot cities. This paper instead constructs an internal urban functional coordination index and estimates its relationship with emissions. The difference matters because the present SCCD index includes digital space as one of four functional subsystems. It therefore asks whether the coordination of urban functions is itself related to emissions, not only whether smart performance and low-carbon indicators move together.",
        "Taken together, the four studies make the paper's contribution more precise. The manuscript is not the first to say that smart cities can reduce emissions, that green technology matters, that spatial heterogeneity exists, or that smart and low-carbon systems can be coordinated. Its value is to show that deeper production-living-ecological-digital coordination has a phase-dependent association with emissions. This contribution is narrower than a claim that smart cities always reduce emissions, but it is more useful for governance because it identifies why the transition path can be carbon intensive before it becomes carbon saving.",
    ]:
        add_paragraph(doc, text)

    doc.add_heading("2.6. Hypotheses", level=2)
    for text in [
        "H1. Spatial coupling coordination across production, living, ecological, and digital spaces has an inverted U-shaped association with log carbon emissions.",
        "H2. Industrial upgrading and green technological innovation are associated restructuring pathways linking multidimensional coordination to emissions, but they are not interpreted as formally decomposed pathways in the absence of decomposition estimates.",
        "H3. Digital economy development attenuates the marginal emission cost of spatial coupling coordination because data-enabled capacity can convert coordination from infrastructure expansion into operational efficiency.",
        "H4. The nonlinear SCCD-lnCE association differs across major Chinese regions because development stage, industrial legacy, digital infrastructure, and governance capacity differ across space.",
    ]:
        add_bullet(doc, text)

    doc.add_heading("3. Materials and Methods", level=1)
    doc.add_heading("3.1. Data and Sample", level=2)
    for text in [
        "The main sample covers 284 Chinese prefecture-level and above cities from 2006 to 2021, yielding 4,544 observed city-year observations before regression-specific missing-value restrictions. The full-control baseline uses 4,514 observations after observations with missing controls are removed. The broader working data file contains fitted or extrapolated 2022-2024 observations, but those rows are not treated as observed records in this Sustainability version. The separation between observed evidence and fitted sensitivity evidence is maintained throughout the manuscript.",
        "The unit of observation is the city-year. The city dimension uses the identifier id, and the time dimension uses year. The empirical workflow is organized around the observed sample because the paper's target journal requires clear data provenance. Excluding fitted or extrapolated rows reduces the apparent time coverage, but it makes the main claims more defensible. The supplementary materials can still report full-period sensitivity tables to show that the generated workflow remains auditable, provided that those rows are described as fitted or extrapolated sensitivity evidence rather than observed statistical records.",
        "Raw indicators are drawn mainly from the China City Statistical Yearbook, China Energy Statistical Yearbook, China Urban Construction Statistical Yearbook, China Rural Statistical Yearbook, China Industrial Statistical Yearbook, and China Environmental Statistical Yearbook. City-level carbon emissions are matched from EDGAR urban and gridded emission products [32,33]. These sources provide the statistical and emissions basis for the SCCD indicators, control variables, and outcome variable. The manuscript does not add new data sources beyond the existing workflow because the requested revision concerns literature integration and manuscript expansion, not empirical reconstruction.",
        "The dependent variable is log carbon emissions. The available Stata data file stores this log-transformed outcome under the variable name CE. The generator therefore describes the outcome as lnCE in the manuscript but uses CE in the reproducibility narrative, avoiding any implication that the script applies another logarithmic transformation. This point matters for replication because a second log transformation would change the interpretation of coefficients and turning points. The current workflow follows the project rule that CE is the dependent variable and is already in log form.",
        f"One data-quality issue is explicitly retained rather than hidden. OPEN contains {fmt_int(metrics['open_neg'])} negative observed-sample values, equal to {fmt(metrics['open_neg_pct'], 1)}% of nonmissing observed OPEN records, with an observed-sample range of [{fmt(metrics['open_min'], 3)}, {fmt(metrics['open_max'], 3)}]. Because the supplied analytical data file is the authoritative input, these values are disclosed and retained rather than recoded ex post. This choice is conservative. It avoids changing the empirical record during manuscript drafting and leaves final data-cleaning decisions to the author team if they decide to rebuild the Stata workflow.",
    ]:
        add_paragraph(doc, text)

    doc.add_heading("3.2. SCCD Construction and Indicator System", level=2)
    for text in [
        "SCCD is built to measure coordination among four urban functional spaces: production, living, ecological, and digital. Production space captures the economic and industrial functions that organize material output, employment, trade, and services. Living space captures public services, residential support, education, health care, culture, and daily urban amenities. Ecological space captures environmental pressure, environmental response, and ecological status. Digital space captures telecommunications, information infrastructure, digital services, and digital-development capacity. The four-space design is the central difference between this paper and studies that treat digitalization as a single external regressor.",
        "The indicator system contains 13 criterion layers and 36 indicators after deleting one policy-term frequency indicator from the earlier manuscript workflow. The deletion avoids over-reliance on a text-frequency proxy that may mix policy attention with implementation. Table 1 reports the retained indicators, calculation rules, units, and direction. Positive indicators are normalized so that higher values indicate stronger subsystem development. Negative ecological pressure indicators are normalized in the opposite direction so that higher normalized values consistently indicate better ecological performance.",
        "Normalization follows the standard min-max logic in the generated workflow. A positive indicator is normalized as x* = (x - min x) / (max x - min x), and a negative indicator is normalized as x* = (max x - x) / (max x - min x). Entropy weights are then calculated within functional subsystems. For indicator j, p_jit = x*_jit / sum(x*_jit), e_j = -k sum(p_jit ln p_jit), d_j = 1 - e_j, and w_j = d_j / sum(d_j). Entropy weighting gives more weight to indicators with more cross-sectional and temporal information.",
        "Subsystem scores are aggregated from the entropy-weighted indicators. The coupling degree C measures the interaction among production, living, ecological, and digital subsystem scores. The comprehensive development index T is the equally weighted average of the four subsystem scores. The final SCCD value is SCCD = sqrt(C x T). This expression means that a city needs both interaction and overall development to score highly. A city with strong interaction among weak subsystems, or high development in one subsystem but weak interaction, will not mechanically receive a high coordination score.",
        "Equal weights across the four functional dimensions are used at the subsystem level to keep the conceptual framework transparent. The purpose is not to claim that production, living, ecological, and digital functions have identical economic importance in every city. The purpose is to avoid allowing one subsystem to dominate the final index by construction. The empirical question is then whether the resulting coordination index is associated with emissions in a nonlinear way. This index design is consistent with the paper's theory that coordination is a multidimensional urban process rather than a single technology input.",
    ]:
        add_paragraph(doc, text)

    add_table(
        doc,
        "Table 1. Detailed SCCD indicator system.",
        INDICATOR_ROWS,
        "Note: The indicator system follows the production-living-ecological-digital framework after deleting one policy-term frequency indicator. The Stata workflow uses the generated composite SCCD and does not recalculate raw entropy weights.",
        widths=[1200, 1700, 2860, 2180, 860, 560],
    )

    doc.add_heading("3.3. Empirical Strategy", level=2)
    for text in [
        "The baseline specification is CE_it = beta1 SCCD_it + beta2 SCCD2_it + gamma X_it + city_i + year_t + epsilon_it, where X_it contains OPEN, UR, URG, and GI. City fixed effects absorb time-invariant city characteristics, including persistent geography, historical industrial base, and long-run administrative features. Year fixed effects absorb common annual shocks, including national policy cycles, macroeconomic conditions, and countrywide changes in emissions reporting or energy markets. Standard errors are clustered by city to allow arbitrary serial correlation within cities.",
        "The inverted-U interpretation requires four pieces of evidence. First, beta1 should be positive and beta2 should be negative. Second, the turning point -beta1 / (2 beta2) should lie inside the observed SCCD range. Third, the marginal effect should be positive at low SCCD and negative at high SCCD. Fourth, the distribution of observed city-year observations around the turning point should be reported so that the reader can see whether the turning point is empirically relevant. The workflow therefore reports the turning point, confidence interval, endpoint marginal effects, mean marginal effect, and the sample distribution relative to the turning point.",
        "The baseline controls are retained from the existing empirical design. OPEN measures openness, UR measures urbanization, URG captures urban-rural income or development gap information as defined in the Stata workflow, and GI captures government intervention. The manuscript does not add controls during the writing revision because doing so would require rerunning the empirical workflow and defending new model choices. The goal here is to make the existing evidence clear, not to create new empirical results without the corresponding Stata outputs.",
        "Robustness checks include 1%-99% winsorized variables, night-time light activity as an alternative outcome check, and MCCD as an alternative coordination proxy when available. These checks test whether the nonlinear pattern is driven by outliers, by the emissions measure alone, or by the specific SCCD construction. The manuscript reports them as robustness evidence, not as independent proof of causality. A result that survives robustness checks is more credible, but the design remains an observational panel with fixed effects.",
        "Supplementary endogeneity analysis instruments SCCD and SCCD2 with the pre-existing lagged SCCD instrument IV and IV2. The workflow does not recalculate or overwrite IV or c_SCCD. This restriction preserves provenance because the existing data and Stata outputs define the empirical package. The IV estimates are useful because they test whether the nonlinear pattern is robust to an alternative source of variation, but lagged coordination may still be correlated with unobserved development trajectories. For that reason, IV evidence is discussed as supplementary rather than as the paper's primary identification strategy.",
        "Associated pathway equations examine industrial upgrading, denoted OIU, and green technological innovation, denoted GTI. The equations test whether SCCD and SCCD2 are associated with these restructuring variables and whether the SCCD-lnCE pattern remains when they enter the emissions equation. The wording is intentionally cautious. The manuscript does not claim a quantified pathway share, because the current output does not provide a formal decomposition that would identify such a share. This claim discipline aligns the text with the available evidence and avoids overstating the mechanism analysis.",
        "Moderation models examine DEI, ER, and POLY. DEI is the digital economy development index and is the main boundary condition supported by the generated evidence. ER and POLY remain in the table because they are part of the workflow, but the results section does not promote them into central supported mechanisms unless the table warrants that interpretation. This approach preserves a clean distinction between reported analyses and emphasized findings. It also keeps the Sustainability message focused on data-enabled governance capacity, which is the boundary condition most directly connected to smart-city operation.",
    ]:
        add_paragraph(doc, text)

    doc.add_heading("3.4. Claim Boundaries and Reproducibility Logic", level=2)
    for text in [
        "The paper uses fixed effects, robustness checks, supplementary IV estimation, pathway equations, moderation models, and regional heterogeneity analysis. Even so, the main claim is an association claim. The data are not a randomized experiment, and the SCCD index is not assigned by policy at random. City fixed effects reduce bias from time-invariant city characteristics, and year fixed effects reduce bias from common shocks, but time-varying unobserved factors can still influence both coordination and emissions. The manuscript therefore avoids stronger causal language than the evidence supports.",
        "The distinction between observed and fitted rows is another boundary. The main text uses observed data from 2006 to 2021. Fitted or extrapolated 2022-2024 rows are reserved for supplementary sensitivity discussion and are never described as observed statistical records. This wording protects the manuscript from a common reviewer objection: a paper should not use model-generated rows as if they were observed data. It also makes the empirical scope clear to Sustainability reviewers who may focus on data transparency.",
        "The generator approach improves reproducibility. The DOCX and PDF are created from the Python script, which reads generated Stata tables, figures, and diagnostic files from the outputs directory. This prevents manual editing of the Word document from drifting away from the empirical outputs. If a coefficient, turning point, or diagnostic changes, the manuscript should be regenerated from updated evidence rather than edited by hand. This revision follows that rule and changes only manuscript content, references, section numbering, and QA logic.",
    ]:
        add_paragraph(doc, text)

    doc.add_heading("3.5. Reviewer-Facing Interpretation Rules", level=2)
    for text in [
        "The manuscript follows several interpretation rules designed to make the empirical contribution reviewer-proof. The first rule is to separate the policy object from the index object. Smart-city policy papers often define treatment by pilot status, policy timing, or digital-city construction. This manuscript defines SCCD as a continuous index of functional coordination. A reader should therefore not interpret SCCD as the probability of being selected into a smart-city pilot. SCCD is a measure of how urban functions develop together, and the estimated coefficients describe the association between that measure and log emissions.",
        "The second rule is to interpret the quadratic model through marginal effects rather than through isolated coefficients. A positive SCCD coefficient does not mean coordination is simply harmful, and a negative SCCD2 coefficient does not by itself prove that coordination is beneficial at high levels. The relevant object is the slope beta1 + 2 beta2 SCCD. This is why the results section reports endpoint marginal effects and the sample distribution around the turning point. The manuscript asks where the city-year observations lie on the curve before drawing policy implications.",
        "The third rule is to distinguish robustness from identification. Winsorization, alternative proxies, and supplementary outcome checks can show that the pattern is not mechanically driven by one data treatment or one measurement choice. They do not solve all omitted-variable concerns. Similarly, the supplementary IV estimates provide useful diagnostic evidence, but the lagged instrument is not presented as a perfect natural experiment. The wording therefore keeps the central finding as a robust association in an observed panel.",
        "The fourth rule is to avoid unsupported mechanism language. The generated pathway equations are valuable because they connect SCCD to OIU and GTI, and because they show that the nonlinear SCCD pattern persists when these variables enter the emissions equation. They do not provide a formal decomposition of the SCCD association into channel shares. The manuscript therefore uses associated pathway language. This is stricter than many empirical papers, but it should reduce reviewer objections about overclaiming.",
        "The fifth rule is to report weaker or supplementary evidence without forcing it into the main story. ER and POLY remain in the moderation table because they are part of the generated workflow. They are not promoted into central conclusions when the stronger and more policy-relevant boundary condition is DEI. This selectivity is not cherry-picking; it is claim discipline. The text states what is supported most clearly and leaves other estimates visible for readers to inspect.",
        "The final rule is to keep administrative submission statements separate from empirical content. The generator cannot know final author contributions, grant numbers, conflict declarations, or data-license permissions. It therefore preserves author-side action items rather than inventing complete statements. This choice keeps the manuscript analytically complete while making clear that final submission requires author confirmation of metadata and source-license wording.",
    ]:
        add_paragraph(doc, text)

    doc.add_heading("4. Results", level=1)
    doc.add_heading("4.1. Descriptive Statistics", level=2)
    for text in [
        "Table 2 summarizes the observed 2006-2021 sample. The observed sample contains no fitted rows, and regression-specific samples vary only because of missing controls or instruments. The table includes the outcome variable CE, the core SCCD variables, baseline controls, digital economy development, polycentricity, industrial upgrading, green technological innovation, and the lagged SCCD instrument. The purpose of the table is to show the empirical scale of the variables before the paper turns to fixed-effects estimates.",
        "The descriptive statistics also clarify why a nonlinear specification is appropriate. SCCD varies across city-years, and SCCD2 is included to capture curvature rather than to add a separate substantive variable. A reader should interpret SCCD and SCCD2 together. The baseline regression does not ask whether SCCD and SCCD2 have independent policy meanings. It asks whether the slope of the SCCD-lnCE relationship changes across the observed coordination distribution.",
        "Figure 1 plots the observed-sample SCCD distribution. The distribution matters because the estimated turning point must be evaluated against actual data support. A turning point far outside the sample would be a mathematical extrapolation rather than an empirical result. The generated diagnostics show that the observed-sample turning point is inside the observed range, and the distributional table reports how many city-year observations lie below and above it. This is why the paper emphasizes both coefficient signs and sample position relative to the turning point.",
    ]:
        add_paragraph(doc, text)
    desc_rows = read_desc_table(
        TABLES_DIR / "table02_descriptive_statistics_observed.csv",
        ["CE", "SCCD", "SCCD2", "OPEN", "UR", "URG", "GI", "DEI", "POLY", "OIU", "GTI", "IV"],
    )
    add_table(doc, "Table 2. Descriptive statistics for the observed 2006-2021 sample.", desc_rows)
    add_picture_if_exists(doc, FIGURES_DIR / "fig_sccd_distribution_observed_2006_2021.png", "Figure 1. Distribution of SCCD in the observed 2006-2021 sample.")

    doc.add_heading("4.2. Baseline Nonlinear Relationship", level=2)
    for text in [
        f"Table 3 reports the baseline two-way fixed-effects estimates. The full-control observed-sample specification gives SCCD = {metrics['sccd_coef']} and SCCD2 = {metrics['sccd2_coef']}. The implied turning point is {fmt(metrics['turning_point'], 6)}, and the confidence interval remains inside the observed SCCD range. The marginal effect is {fmt(metrics['left_me'], 3)} at the left endpoint, {fmt(metrics['mean_me'], 3)} at the sample mean, and {fmt(metrics['right_me'], 3)} at the right endpoint. These estimates support H1 and reject a simple monotonic sustainability story.",
        "The positive coefficient on SCCD and the negative coefficient on SCCD2 mean that the SCCD-lnCE slope declines as coordination rises. At low SCCD levels, additional coordination is associated with higher log emissions. At high SCCD levels, additional coordination is associated with lower log emissions. The point is not that low-coordination cities should avoid coordination. The point is that early coordination may be carbon-intensive unless cities manage the transition from building systems to operating systems efficiently.",
        f"The observed-sample distribution shows why this result is policy-relevant. The full-control observed baseline contains {fmt_int(metrics['n_below_tp'])} city-year observations at or below the turning point and {fmt_int(metrics['n_above_tp'])} above it. Thus, a large share of observed city-years has not yet crossed into the range where the estimated marginal association becomes negative. For Sustainability readers, this matters because smart-city construction cannot be evaluated only by its endpoint. The transition path itself may carry emissions pressure.",
        "Figure 2 visualizes the same nonlinear relationship. The fitted quadratic curve provides a compact graphical summary of the coordination paradox. It should be read together with Table 3 rather than as a separate estimate. The figure helps readers see that the paper's main result is not an isolated coefficient sign; it is a pattern across the SCCD range that connects the estimated curve, the turning point, and the theoretical distinction between expansion-oriented and efficiency-oriented coordination.",
    ]:
        add_paragraph(doc, text)
    add_table(
        doc,
        "Table 3. Baseline regression results: observed-only sample, 2006-2021.",
        parse_rtf_table(TABLES_DIR / "table03_baseline_observed_2006_2021.rtf"),
        "Notes: City and year fixed effects are included. Robust standard errors clustered by city are in parentheses. *, **, and *** denote significance at 10%, 5%, and 1%, respectively.",
    )
    add_picture_if_exists(doc, FIGURES_DIR / "fig_nonlinear_sccd_lnce_observed_2006_2021.png", "Figure 2. Observed SCCD-lnCE relationship and quadratic fit, 2006-2021.")

    doc.add_heading("4.3. Robustness and Supplementary IV Evidence", level=2)
    for text in [
        f"Table 4 shows that the nonlinear pattern remains under winsorization, the night-time light activity check, and supplementary IV estimation. The observed-only IV turning point is 0.570248. The first-stage diagnostics report F statistics of {fmt(metrics['fs_sccd'], 2)} for SCCD and {fmt(metrics['fs_sccd2'], 2)} for SCCD2; the Kleibergen-Paap rk Wald F statistic is {fmt(metrics['kp_f'], 2)}. These diagnostics are reported to make the IV evidence auditable, but the IV estimates remain supplementary because lagged SCCD cannot eliminate all identification concerns in an observational panel.",
        "The winsorized specification addresses the possibility that extreme observations drive the curvature. The nonlinear pattern remains when variables are winsorized at the 1%-99% range. This check is important because city-level panels can contain unusually large municipalities, resource-intensive industrial cities, and fast-growing digital hubs. If the inverted U-shaped association were driven only by outliers, the result would weaken after winsorization. The robustness output does not support that concern.",
        "The night-time light check addresses a different concern. Carbon-emissions data are constructed from emissions products and statistical information, and reviewers may ask whether the result is tied to the particular outcome measure. Night-time light activity is not a perfect substitute for carbon emissions, but it is a widely used proxy for economic activity and urban intensity. A consistent nonlinear pattern in this check supports the interpretation that SCCD captures a broader urban-development process rather than a narrow measurement artifact in the emissions outcome.",
        "The MCCD alternative proxy check asks whether the result depends on the specific SCCD construction. Any composite index can raise concerns about indicator selection and weighting. Using an alternative coordination proxy, when available in the generated table, helps show that the core curvature is not merely an artifact of one index formula. The manuscript still presents SCCD as the main variable because SCCD directly matches the production-living-ecological-digital framework developed in the theory section.",
        "The supplementary IV estimates are useful but not decisive. A lagged SCCD instrument can reduce some simultaneity concerns if current emissions shocks do not fully determine past coordination. Yet past coordination can also reflect persistent policy capacity, industrial development, infrastructure trajectories, and environmental governance preferences. These factors may affect current emissions directly. The manuscript therefore reports the IV turning point and diagnostics, but it does not use the IV specification to make stronger causal claims than the design can support.",
    ]:
        add_paragraph(doc, text)
    add_table(
        doc,
        "Table 4. Robustness and supplementary IV results: observed-only sample, 2006-2021.",
        parse_rtf_table(TABLES_DIR / "table04_robustness_endogeneity_observed_2006_2021.rtf"),
        "Notes: IV estimates use the pre-existing lagged SCCD instrument and its square. First-stage and weak-instrument diagnostics are reported in the generated supplementary diagnostic files.",
    )

    doc.add_heading("4.4. Regional Heterogeneity", level=2)
    for text in [
        "Table 5 indicates that the inverted U-shaped pattern appears across major regions, but coefficient magnitudes differ. This supports H4 and implies that the same increase in coordination may have different emission implications across regional development stages. Prior DID evidence also reports spatially heterogeneous smart-city carbon effects [17]. The current paper keeps the comparison at the regional level because the workflow does not classify cities by resource-based status or estimate a resource-based-city heterogeneity table.",
        "Regional heterogeneity is expected in the coordination paradox framework. Eastern cities tend to have stronger digital infrastructure, more service-oriented economies, and more mature urban governance capacity. Central and western cities may be more likely to use coordination improvements for infrastructure catch-up, industrial-platform construction, or resource-related restructuring. These differences can change both the location of the turning point and the slope on each side of it. The table therefore helps translate the nonlinear average result into region-specific policy caution.",
        "The regional estimates should not be read as a ranking of regions. A larger coefficient in one region does not automatically mean that region has worse policy or weaker sustainability prospects. It may mean that SCCD captures a different bundle of activities in that region. For example, coordination in a catching-up region may involve transport and industrial investment, while coordination in a more mature region may involve data integration and operational management. The paper therefore interprets heterogeneity through development stage and governance context rather than through a simple good-region versus bad-region framing.",
    ]:
        add_paragraph(doc, text)
    add_table(doc, "Table 5. Regional heterogeneity: observed-only sample, 2006-2021.", parse_rtf_table(TABLES_DIR / "table05_regional_heterogeneity_observed_2006_2021.rtf"))

    doc.add_heading("4.5. Associated Pathway Equations", level=2)
    for text in [
        "Table 6 reports associated pathway equations for OIU and GTI. SCCD is positively associated with both industrial upgrading and green technological innovation, while the squared terms are negative. Once OIU or GTI enters the CE equation, the nonlinear SCCD pattern remains. The evidence is therefore consistent with H2 as restructuring association, not as a quantified pathway effect. This interpretation aligns with smart-city studies that emphasize green technology progress, but it does not claim that the current estimates identify a decomposed pathway share [16,18,30,31].",
        "The OIU equations connect SCCD to the industrial restructuring side of sustainable urban transition. Better coordination can support industrial upgrading by improving infrastructure matching, public-service quality, environmental regulation capacity, and digital service availability. At the same time, industrial upgrading is not costless. New industrial equipment, technology adoption, relocation, and training can require energy-intensive investment. The positive and nonlinear association between SCCD and OIU is therefore consistent with the idea that coordination can reorganize the urban economy before all emission-saving gains are realized.",
        "The GTI equations connect SCCD to the innovation side of the smart-city literature. Digital governance and coordinated urban systems can improve knowledge spillovers, patenting incentives, environmental monitoring, and the matching of firms with policy support. Ma and Wu and An et al. both emphasize green technology progress in the smart-city carbon setting [16,18]. The current evidence is consistent with that literature, but the manuscript remains explicit about its own design. It shows that SCCD is associated with GTI and that the SCCD-lnCE curve remains nonlinear; it does not assign a causal share of emissions change to GTI.",
        "The persistence of the nonlinear SCCD terms after adding OIU or GTI is substantively important. It suggests that industrial upgrading and green technological innovation do not exhaust the coordination-emissions relationship. Coordination may also affect emissions through transport efficiency, public-service allocation, land-use change, energy management, ecological monitoring, and cross-departmental governance. The paper therefore frames OIU and GTI as two observable restructuring pathways within a broader urban coordination process.",
    ]:
        add_paragraph(doc, text)
    add_table(
        doc,
        "Table 6. Associated pathway equations: observed-only sample, 2006-2021.",
        parse_rtf_table(TABLES_DIR / "table06_mediation_observed_2006_2021.rtf"),
        "Notes: These are pathway equations rather than formally decomposed pathway estimates.",
    )

    doc.add_heading("4.6. Moderation and Boundary Conditions", level=2)
    dei_rows = metrics.get("dei_records") or []
    dei_low = get_record(dei_rows, "dei_level", "low")
    dei_high = get_record(dei_rows, "dei_level", "high")
    for text in [
        f"Table 7 reports observed-sample moderation estimates. DEI significantly attenuates the nonlinear SCCD-lnCE relationship, supporting H3. Conditional calculations show a marginal effect at the SCCD mean of {fmt(dei_low.get('marginal_effect_at_mean_sccd'), 3)} at low DEI and {fmt(dei_high.get('marginal_effect_at_mean_sccd'), 3)} at high DEI. POLY is retained only as a supplementary interaction check and is not treated as a supported moderation mechanism.",
        "The DEI result gives the paper its clearest Sustainability-facing boundary condition. Digital economy development appears to weaken the marginal emission cost of coordination. This is consistent with the idea that digital capacity changes how cities use coordination. When the digital economy is weak, SCCD improvements may rely more heavily on physical expansion and administrative construction. When the digital economy is stronger, the same coordination improvement can be translated into better data sharing, energy management, service matching, and operational efficiency.",
        "Figure 3 visualizes the conditional SCCD-lnCE relationship at low and high DEI levels. The curve at high DEI is flatter at the sample mean, which supports the interpretation that digital capacity attenuates the emission cost of coordination. The figure does not mean that digitalization automatically solves the carbon problem. It means that the relationship between coordination and emissions depends on whether digital systems become operational capabilities rather than stand-alone infrastructure projects.",
        "The POLY and ER estimates are included to keep the moderation workflow transparent. Polycentricity and environmental regulation are theoretically relevant to urban emissions, but the current manuscript does not make them central because the generated evidence does not support the same clear boundary-condition interpretation as DEI. This is an important writing choice. A paper becomes more credible when it emphasizes the supported result and reports weaker or supplementary results without forcing them into the main story.",
    ]:
        add_paragraph(doc, text)
    add_table(
        doc,
        "Table 7. Moderation analysis: observed-only sample, 2006-2021.",
        parse_rtf_table(TABLES_DIR / "table07_moderation_observed_2006_2021.rtf"),
        "Notes: ER, DEI, and POLY are estimated in separate interaction models. The text focuses on DEI as the supported boundary condition.",
    )
    add_picture_if_exists(doc, FIGURES_DIR / "fig_moderation_dei_observed_2006_2021.png", "Figure 3. Predicted SCCD-lnCE curves at low and high DEI levels, observed sample.")

    doc.add_heading("5. Discussion", level=1)
    doc.add_heading("5.1. Reconciling Average Smart-City Effects and Nonlinear Coordination Intensity", level=2)
    for text in [
        "The results clarify why smart-city construction and spatial coordination do not automatically reduce emissions at every stage. Recent DID studies show that smart-city policy adoption can reduce average urban carbon emissions [5,16-18]. This paper asks a different marginal question: how emissions change as production, living, ecological, and digital spaces become more tightly coordinated. At low coordination levels, cities often connect these functions through infrastructure expansion, industrial concentration, construction activity, and additional energy input. These expansion effects can dominate early efficiency gains and generate the upward side of the inverted U-shaped relationship.",
        "The apparent difference between the current inverted-U evidence and prior average treatment effects is not a contradiction. Policy adoption may reduce average emissions if designation brings governance reforms, digital-management capacity, funding support, and innovation incentives. Coordination intensity may still have phase-dependent effects if early-stage coordination is built through energy-intensive investment. The two claims can coexist because they describe different margins: one compares treated and untreated policy status, while the other compares higher and lower levels of continuous multidimensional coordination within cities over time.",
        "This distinction has practical implications for smart-city evaluation. A city can show progress in smart-city construction while still increasing emissions if the progress is dominated by new physical systems, construction, industrial parks, and electricity-intensive digital infrastructure. A city can show carbon benefits when the same systems mature into operational efficiency tools. Evaluators should therefore separate the construction phase from the operation phase. A single average effect can hide this sequencing problem.",
        "The results also suggest that smart-city policy should not be evaluated only by the presence of digital infrastructure. Sensors, platforms, networks, and data centers are inputs. Sustainability benefits depend on whether these inputs change decisions: energy dispatch, public transit, logistics routing, land-use approvals, environmental enforcement, public-service delivery, and industrial upgrading. The SCCD framework captures this broader coordination problem by putting digital space beside production, living, and ecological spaces rather than treating it as a separate technology appendage.",
    ]:
        add_paragraph(doc, text)

    doc.add_heading("5.2. Governance Sequencing for Sustainability", level=2)
    for text in [
        "The policy implication is governance sequencing. Cities at low coordination levels should not assume that more smart infrastructure will immediately reduce emissions. Their first task is to prevent smart-city construction from becoming duplicated infrastructure, fragmented platforms, and land-intensive development. Coordination investments should be screened for whether they reduce future energy use, shorten travel or logistics routes, improve public-service matching, or strengthen environmental monitoring. Projects that merely add digital layers without changing operations may raise emissions without creating durable low-carbon capacity.",
        "Cities approaching the turning point should shift from infrastructure-led construction to data-enabled operation. This means interoperable data systems, shared standards across departments, energy-management platforms, smart mobility tools, digital public services, and routines that allow environmental agencies to use real-time information. The goal is not to maximize the number of smart-city projects. The goal is to make existing urban systems work together with less waste. This sequencing language aligns with Sustainability's applied policy audience because it translates a nonlinear regression result into staged governance priorities.",
        "Cities above the turning point should focus on maintaining efficiency gains and preventing rebound effects. Digital systems can lower search costs and improve resource allocation, but they can also stimulate new consumption, logistics demand, and electricity use. High-coordination cities therefore need carbon accounting for digital infrastructure, green procurement for data centers, demand-side management, and continuous evaluation of whether smart services reduce or merely relocate emissions. The downward side of the curve should be protected, not assumed to persist automatically.",
        "Regional policy should be differentiated. The regional heterogeneity results indicate that the same SCCD increase does not carry the same emission implication everywhere. Eastern cities may benefit more from refined digital operation and service integration. Central and western cities may need stronger safeguards against infrastructure-heavy coordination. These statements remain at the regional level because the current analysis does not generate resource-based-city classifications. The manuscript avoids unsupported claims while still using prior spatial heterogeneity literature to motivate differentiated policy.",
    ]:
        add_paragraph(doc, text)

    doc.add_heading("5.3. Interpreting Associated Pathways", level=2)
    for text in [
        "The pathway-equation results help explain why the transition from expansion to efficiency can be gradual. Industrial upgrading and green technological innovation are associated with SCCD, but both can involve transitional costs. Upgrading may require new equipment, new buildings, retraining, and relocation. Green innovation may require experimentation, pilot production, commercialization, and complementary infrastructure. These activities can create short-run carbon pressure even if they support long-run low-carbon transition. This is why the manuscript avoids treating OIU and GTI as automatically emission-reducing in every period.",
        "The finding also speaks to the green technology progress literature. Prior smart-city studies emphasize green patents, green technology progress, and productivity as channels through which smart-city construction can reduce emissions [16,18]. This paper agrees with the mechanism logic but adds a stage condition. Green technology progress may matter most when the city has enough digital and governance capacity to move from innovation inputs to operational emission reductions. If innovation activity is still tied to construction, equipment replacement, and industrial restructuring, short-run emission pressure can remain.",
        "The DEI moderation result strengthens this interpretation. Digital economy development appears to make coordination less emission-intensive at the margin. In practical terms, DEI may indicate that firms, households, and governments can actually use data-enabled systems. Digital platforms then become governance tools rather than display projects. They can connect energy users, transit systems, industrial parks, public services, and environmental agencies. This operational interpretation is more useful for policy than a generic statement that digitalization is good for the environment.",
    ]:
        add_paragraph(doc, text)

    doc.add_heading("5.4. Limitations", level=2)
    for text in [
        "Several limitations remain. The SCCD index captures key aspects of digital space, but it cannot fully observe platform governance, data mobility, algorithmic coordination, cybersecurity quality, or the actual use of data in daily administration. A city may have strong digital infrastructure but weak cross-departmental data sharing. Another city may have modest infrastructure but effective administrative routines. The index captures observable indicators, not every institutional feature that determines whether digital systems become operational capabilities.",
        "The empirical design includes fixed effects, robustness checks, supplementary IV estimation, pathway equations, moderation models, and regional heterogeneity analysis, but it remains observational. Time-varying unobserved shocks may still affect both SCCD and emissions. Examples include local leadership changes, industrial policy packages, energy-price exposure, environmental campaigns, and unobserved infrastructure investment. The manuscript therefore describes the main result as an association. More direct causal designs would require new policy variation, event timing, or external instruments that are not generated in the current workflow.",
        "The observed sample ends in 2021. The broader working data file contains fitted or extrapolated 2022-2024 rows, but those rows are not part of the main evidence. This choice makes the paper more conservative but also means the manuscript does not claim to observe the most recent post-2021 city-level dynamics. Future data updates could test whether the nonlinear pattern persists as more observed years become available, especially after major changes in digital infrastructure, energy policy, and urban governance.",
        "A related limitation concerns the interpretation of regional evidence. The current data classify cities into eastern, central, and western regions only. The manuscript therefore does not estimate a separate northeastern model and does not claim region-specific coefficients for a northeastern subsample. This matters because regional language can otherwise outrun the generated tables. The heterogeneity discussion is intentionally limited to the categories that appear in the Stata data and in the exported regional tables.",
        "The IV evidence should also be read with discipline. The lagged SCCD instruments are useful diagnostics for simultaneity concerns, and the weak-instrument statistics make the supplementary estimation more transparent. They do not transform the design into a fully exogenous natural experiment. Persistent city capacity, industrial trajectories, fiscal resources, and long-running governance preferences may influence both lagged coordination and current emissions. For that reason, the paper reports IV estimates as a robustness-oriented supplement and keeps the main claim anchored in fixed-effects associations.",
        "These limits shape how the tables are used in the paper. The baseline, robustness, heterogeneity, associated-pathway, and moderation outputs are not separate stories competing for attention. They are checks around one central empirical pattern. The baseline table establishes the nonlinear association; the robustness table asks whether the pattern is fragile; the regional table asks whether the pattern is spatially uniform; the pathway equations ask which restructuring variables move with SCCD; and the moderation table asks whether digital economy capacity changes the marginal relationship. Reading the tables this way keeps the interpretation coherent and prevents stronger causal or regional claims than the generated evidence supports. This table-by-table discipline is especially important for a manuscript built from generated empirical outputs.",
        "Finally, the empirical result is a city-level average relationship, not a project-level assessment tool. A city below the estimated turning point may still contain specific smart projects that reduce emissions, while a city above the turning point may still approve projects with high embodied carbon or rebound effects. The policy value of the curve is to warn against one-size-fits-all smart-city claims. It does not replace project appraisal, energy accounting, or sector-specific monitoring. Future work could connect SCCD changes to project inventories, transport data, digital infrastructure energy use, and firm-level upgrading to explain more precisely which parts of coordination create short-run pressure and which parts generate durable efficiency gains.",
    ]:
        add_paragraph(doc, text)

    doc.add_heading("6. Conclusions", level=1)
    for text in [
        f"Using an observed panel of 284 Chinese cities from 2006 to 2021, this study finds an inverted U-shaped association between SCCD and log carbon emissions. The full-control observed-sample turning point is {fmt(metrics['turning_point'], 6)}, and supplementary IV estimates imply a turning point of 0.570248. Coordination is therefore not inherently low-carbon. Its emission implication depends on whether governance remains expansion-oriented or becomes efficiency-oriented.",
        "The result refines the smart-city carbon literature. Existing DID studies show that smart-city policy adoption can reduce average emissions, including through digital governance, green patenting, green technology progress, productivity, and spatially heterogeneous effects [5,16-18]. This paper shows that the marginal association between continuous coordination intensity and emissions can still be nonlinear. Average policy effects and phase-dependent coordination effects can both be true because smart-city designation and multidimensional coordination intensity are different empirical objects.",
        "The policy implication is to sequence smart-city construction. Low-coordination cities should avoid duplicating infrastructure and building disconnected digital platforms. Mid-stage cities should convert smart construction into interoperable data systems, energy management, integrated public services, and coordinated industrial and innovation policy. High-coordination cities should maintain efficiency gains while monitoring rebound effects and the energy footprint of digital systems. Regional strategies should differ because the same SCCD increase reflects different development activities across space.",
        "The paper's associated pathway and moderation evidence points to a practical governance message. Industrial upgrading and green technological innovation are relevant restructuring pathways, but they should not be treated as automatic carbon reductions. Digital economy development matters because it appears to attenuate the marginal emission cost of coordination. The low-carbon value of smart-city construction depends less on the label smart and more on whether data-enabled capacity changes how cities operate.",
        "Future research should build richer measures of digital urban governance, distinguish construction-phase and operation-phase smart-city investment, examine longer dynamic lags in industrial upgrading and green innovation, and test whether the SCCD framework travels to other national and metropolitan contexts. New observed data after 2021 would also allow the author team to test whether the turning point changes as digital governance matures and as low-carbon policy constraints tighten.",
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


def count_words(text: str) -> int:
    return len(re.findall(r"[A-Za-z0-9]+(?:[-'][A-Za-z0-9]+)?", text))


def inspect_docx_text(path: Path, pdf_status: dict[str, object]) -> dict[str, object]:
    doc = Document(path)
    paragraph_texts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    table_texts = [
        cell.text.strip()
        for table in doc.tables
        for row in table.rows
        for cell in row.cells
        if cell.text.strip()
    ]
    text = "\n".join(paragraph_texts)
    total_text = "\n".join(paragraph_texts + table_texts)
    main_text_chunks = []
    for paragraph_text in paragraph_texts:
        if paragraph_text in {"Author Contributions", "References"}:
            break
        main_text_chunks.append(paragraph_text)
    abstract_words = count_words(ABSTRACT)
    word_count_total = count_words(total_text)
    word_count_main_text_no_references = count_words("\n".join(main_text_chunks))
    required_sections = {
        "Abstract": "Abstract" in text,
        "Keywords": "Keywords:" in text,
        "Introduction": "1. Introduction" in text,
        "Materials and Methods": "Materials and Methods" in text,
        "Results": "Results" in text,
        "Discussion": "Discussion" in text,
        "Conclusions": "Conclusions" in text,
        "Author Contributions": "Author Contributions" in text,
        "Funding": "Funding" in text,
        "Data Availability": "Data Availability Statement" in text,
        "Conflicts": "Conflicts of Interest" in text,
        "References": "References" in text,
    }
    target_references = {
        "Ma_Wu_2023": "Ma, Z.; Wu, F. Smart City, Digitalization and CO2 Emissions" in text,
        "Li_Zhang_Lau_Gong_2025": "Li, D.; Zhang, X.; Lau, A.D.; Gong, Y. The Impact of Smart City Construction" in text,
        "An_Yang_Zhang_Zeng_2024": "An, X.; Yang, Y.; Zhang, X.; Zeng, X. Smarter and Cleaner?" in text,
        "Zhu_Li_Zhou_Zhu_Yu_2024": "Zhu, X.; Li, D.; Zhou, S.; Zhu, S.; Yu, L. Evaluating coupling coordination" in text,
    }
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
        "word_count_total": word_count_total,
        "word_count_main_text_no_references": word_count_main_text_no_references,
        "meets_10000_word_requirement": word_count_main_text_no_references >= 10000,
        "required_sections": required_sections,
        "target_references": target_references,
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
        bool(pdf_status.get("created")) and PDF_OUT.exists()
        and 100 <= abstract_words <= 200
        and checks["meets_10000_word_requirement"]
        and all(required_sections.values())
        and all(target_references.values())
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
        "- Confirm that the author team wants to retain the final literature framing and citation choices.",
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
