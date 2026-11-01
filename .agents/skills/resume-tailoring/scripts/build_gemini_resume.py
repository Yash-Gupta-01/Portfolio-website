#!/usr/bin/env python3
"""
build_gemini_resume.py - Autonomous Headless Résumé Generator for Yash Gupta
Designed for Gemini Code Execution Sandbox (Linux VM) and Cross-Platform Execution.

Features:
- Pure Word OpenXML formatting with customizable font engine (Noto Serif / Caladea / Georgia).
- CLI-driven: decouples data from code via a lightweight resume_data.json file.
- Markdown bold parsing: writes bullets naturally as "Engineered an enterprise **GPA platform**..."
- Headless PDF conversion: native LibreOffice CLI execution (with Windows COM fallback).
- Automated Quality & Visual QA:
    1. Banned buzzword scan & action verb deduplication pass.
    2. Strict 1-page boundary validation via pypdf.
    3. Multimodal image render via pypdfium2 (PNG) for visual self-correction.
"""

import os
import sys
import json
import argparse
import subprocess
import collections
from typing import List, Tuple, Dict, Any, Union

import docx.opc.constants
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# ---------------------------------------------------------------------------
# Default Configuration
# ---------------------------------------------------------------------------
DEFAULT_STYLING = {
    "font_family": "Noto Serif",       # Resolves cleanly in Gemini Linux VM; Georgia or Caladea also supported
    "body_font_size": 9.2,             # Tunable: 9.0 - 9.2pt
    "title_font_size": 21.0,           # Candidate name
    "section_heading_size": 11.5,      # Section headers (Experience, Education, etc.)
    "role_title_size": 9.8,            # Role titles & Degree names
    "sub_heading_size": 9.2,           # Company names & Institutions
    "contact_font_size": 8.8,          # Header contact rows
    "margins": {
        "top": 0.36,
        "bottom": 0.36,
        "left": 0.45,
        "right": 0.45
    },
    "tab_stop_right": 7.6,             # Right tab stop for dates/locations (8.5 - 0.45 - 0.45 = 7.6)
    "bullet_spacing": {
        "space_before": 0.0,
        "space_after": 1.8,            # Tunable: 1.6 - 2.2pt
        "left_indent": 0.18,
        "first_line_indent": -0.14
    },
    "section_spacing": {
        "space_before": 5.5,
        "space_after": 2.2
    },
    "role_spacing": {
        "space_before": 3.5,
        "space_after_title": 1.0,
        "space_after_sub": 2.2
    }
}

BANNED_BUZZWORDS = [
    'problem solving', 'dynamic', 'team player', 'hardworking',
    'passionate', 'detail-oriented', 'results-driven', 'self-starter',
    'quick learner', 'fast learner', 'enthusiastic', 'motivated',
    'synergy', 'go-getter', 'strategic thinker', 'think outside the box',
    'rockstar', 'ninja', 'various', 'proven track record'
]


# ---------------------------------------------------------------------------
# Markdown Bold Parser
# ---------------------------------------------------------------------------
def parse_markdown_spans(text_or_spans: Union[str, List[Any]]) -> List[Tuple[str, bool]]:
    """
    Parses bullet input into a list of (text, is_bold) tuples.
    Supports either:
      1. Plain string with markdown **bold** tags:
         "Engineered an enterprise **GPA platform** with **NestJS**."
      2. Structured lists of dicts: [{"text": "...", "bold": true}, ...]
      3. Structured lists of pairs: [["...", true], ...]
    """
    if isinstance(text_or_spans, list):
        parsed = []
        for item in text_or_spans:
            if isinstance(item, dict):
                parsed.append((item.get("text", ""), bool(item.get("bold", False))))
            elif isinstance(item, (list, tuple)) and len(item) >= 2:
                parsed.append((str(item[0]), bool(item[1])))
            elif isinstance(item, str):
                parsed.extend(parse_markdown_spans(item))
        return parsed

    if not isinstance(text_or_spans, str):
        return [(str(text_or_spans), False)]

    # Split on markdown **bold** delimiters
    parts = text_or_spans.split("**")
    spans = []
    for i, part in enumerate(parts):
        if part:
            is_bold = (i % 2 == 1)
            spans.append((part, is_bold))
    return spans


# ---------------------------------------------------------------------------
# Low-Level OpenXML Helpers
# ---------------------------------------------------------------------------
def add_hyperlink(paragraph, url: str, text: str, font_family: str, font_size: float = 8.8,
                  bold: bool = False, italic: bool = False, color: str = "111111", underline: bool = False):
    """Inserts a clickable hyperlink into a python-docx paragraph using native OpenXML."""
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = parse_xml(f'<w:hyperlink {nsdecls("w", "r")} r:id="{r_id}"/>')
    new_run = parse_xml(f'<w:r {nsdecls("w")}/>')
    rPr = parse_xml(f'<w:rPr {nsdecls("w")}/>')

    rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="{font_family}" w:hAnsi="{font_family}"/>')
    rPr.append(rFonts)

    sz = int(font_size * 2)
    rPr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="{sz}"/>'))

    if bold:
        rPr.append(parse_xml(f'<w:b {nsdecls("w")}/>'))
    if italic:
        rPr.append(parse_xml(f'<w:i {nsdecls("w")}/>'))
    if color:
        rPr.append(parse_xml(f'<w:color {nsdecls("w")} w:val="{color}"/>'))
    if underline:
        rPr.append(parse_xml(f'<w:u {nsdecls("w")} w:val="single"/>'))
    else:
        rPr.append(parse_xml(f'<w:u {nsdecls("w")} w:val="none"/>'))

    new_run.append(rPr)
    t = parse_xml(f'<w:t {nsdecls("w")} xml:space="preserve">{text}</w:t>')
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


# ---------------------------------------------------------------------------
# Document Construction Engine
# ---------------------------------------------------------------------------
def build_resume_from_data(data: Dict[str, Any], output_docx_path: str):
    """Renders the resume Document according to executive 1-page specifications."""
    styling = {**DEFAULT_STYLING, **data.get("styling", {})}
    margins = {**DEFAULT_STYLING["margins"], **styling.get("margins", {})}
    bullet_cfg = {**DEFAULT_STYLING["bullet_spacing"], **styling.get("bullet_spacing", {})}
    section_cfg = {**DEFAULT_STYLING["section_spacing"], **styling.get("section_spacing", {})}
    role_cfg = {**DEFAULT_STYLING["role_spacing"], **styling.get("role_spacing", {})}

    font_family = styling.get("font_family", "Noto Serif")
    body_font_size = float(styling.get("body_font_size", 9.2))
    right_tab = float(styling.get("tab_stop_right", 7.6))

    doc = Document()

    # Configure Page Geometry (Letterpaper)
    for section in doc.sections:
        section.top_margin = Inches(margins["top"])
        section.bottom_margin = Inches(margins["bottom"])
        section.left_margin = Inches(margins["left"])
        section.right_margin = Inches(margins["right"])
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)

    # Base typography styling
    normal_style = doc.styles['Normal']
    font = normal_style.font
    font.name = font_family
    font.size = Pt(body_font_size)
    font.color.rgb = RGBColor(0x11, 0x11, 0x11)
    normal_style.paragraph_format.line_spacing = 1.0

    # 1. Header Section
    contact = data.get("contact", {})
    candidate_name = contact.get("candidate_name", "Yash Gupta")

    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(2.0)
    p_name.paragraph_format.line_spacing = 1.0
    r_name = p_name.add_run(candidate_name)
    r_name.font.size = Pt(styling.get("title_font_size", 21.0))
    r_name.bold = True
    r_name.font.name = font_family

    # Contact Line 1: Location | Phone | Email
    p_c1 = doc.add_paragraph()
    p_c1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_c1.paragraph_format.space_before = Pt(0)
    p_c1.paragraph_format.space_after = Pt(1.5)
    p_c1.paragraph_format.line_spacing = 1.0

    loc_str = contact.get("location", "New Delhi, India")
    phone_str = contact.get("phone", "+91 8287552253")
    email_str = contact.get("email", "guptayash112003@gmail.com")

    r_c1 = p_c1.add_run(f"{loc_str}  |  {phone_str}  |  ")
    r_c1.font.size = Pt(styling.get("contact_font_size", 8.8))
    r_c1.font.name = font_family
    add_hyperlink(p_c1, f"mailto:{email_str}", email_str, font_family=font_family,
                  font_size=styling.get("contact_font_size", 8.8))

    # Contact Line 2: LinkedIn | GitHub
    p_c2 = doc.add_paragraph()
    p_c2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_c2.paragraph_format.space_before = Pt(0)
    p_c2.paragraph_format.space_after = Pt(4.5)
    p_c2.paragraph_format.line_spacing = 1.0

    linkedin_display = contact.get("linkedin_display", "linkedin.com/in/yash-gupta-engineer")
    linkedin_url = contact.get("linkedin_url", "https://www.linkedin.com/in/yash-gupta-engineer/")
    github_display = contact.get("github_display", "github.com/Yash-Gupta-01")
    github_url = contact.get("github_url", "https://github.com/Yash-Gupta-01")

    add_hyperlink(p_c2, linkedin_url, linkedin_display, font_family=font_family,
                  font_size=styling.get("contact_font_size", 8.8))
    r_sep = p_c2.add_run("  |  ")
    r_sep.font.size = Pt(styling.get("contact_font_size", 8.8))
    r_sep.font.name = font_family
    add_hyperlink(p_c2, github_url, github_display, font_family=font_family,
                  font_size=styling.get("contact_font_size", 8.8))

    # Helper: Add Section Header with solid bottom border
    def add_section_header(title: str):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(section_cfg["space_before"])
        p.paragraph_format.space_after = Pt(section_cfg["space_after"])
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.keep_with_next = True
        run = p.add_run(title)
        run.font.size = Pt(styling.get("section_heading_size", 11.5))
        run.bold = True
        run.font.name = font_family
        pPr = p._p.get_or_add_pPr()
        pBdr = parse_xml(r'<w:pBdr %s><w:bottom w:val="single" w:sz="6" w:space="2" w:color="000000"/></w:pBdr>' % nsdecls('w'))
        pPr.append(pBdr)

    # Helper: Add Bullet
    def add_bullet(spans_or_text: Union[str, List[Any]], space_after: float = None):
        if space_after is None:
            space_after = bullet_cfg["space_after"]
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(bullet_cfg["space_before"])
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.left_indent = Inches(bullet_cfg["left_indent"])
        p.paragraph_format.first_line_indent = Inches(bullet_cfg["first_line_indent"])
        
        r_bullet = p.add_run("•  ")
        r_bullet.font.size = Pt(body_font_size)
        r_bullet.font.name = font_family

        spans = parse_markdown_spans(spans_or_text)
        for text, bold in spans:
            r = p.add_run(text)
            r.bold = bold
            r.font.size = Pt(body_font_size)
            r.font.name = font_family

    # 2. Iterate through sections
    sections = data.get("sections", [])
    for sec in sections:
        sec_type = sec.get("type", "").lower()
        sec_title = sec.get("title", sec_type.capitalize())

        add_section_header(sec_title)

        if sec_type == "experience":
            roles = sec.get("roles", [])
            for r_idx, role in enumerate(roles):
                p1 = doc.add_paragraph()
                p1.paragraph_format.space_before = Pt(role_cfg["space_before"] if r_idx > 0 else 2.5)
                p1.paragraph_format.space_after = Pt(role_cfg["space_after_title"])
                p1.paragraph_format.line_spacing = 1.0
                p1.paragraph_format.keep_with_next = True
                p1.paragraph_format.tab_stops.add_tab_stop(Inches(right_tab), WD_TAB_ALIGNMENT.RIGHT)

                r_title = p1.add_run(role.get("title", ""))
                r_title.bold = True
                r_title.font.size = Pt(styling.get("role_title_size", 9.8))
                r_title.font.name = font_family
                p1.add_run("\t")
                r_dates = p1.add_run(role.get("dates", ""))
                r_dates.bold = True
                r_dates.font.size = Pt(styling.get("role_title_size", 9.8))
                r_dates.font.name = font_family

                p2 = doc.add_paragraph()
                p2.paragraph_format.space_before = Pt(0)
                p2.paragraph_format.space_after = Pt(role_cfg["space_after_sub"])
                p2.paragraph_format.line_spacing = 1.0
                p2.paragraph_format.keep_with_next = True
                p2.paragraph_format.tab_stops.add_tab_stop(Inches(right_tab), WD_TAB_ALIGNMENT.RIGHT)

                r_comp = p2.add_run(role.get("company", ""))
                r_comp.italic = True
                r_comp.font.size = Pt(styling.get("sub_heading_size", 9.2))
                r_comp.font.name = font_family
                p2.add_run("\t")
                r_loc = p2.add_run(role.get("location", ""))
                r_loc.italic = True
                r_loc.font.size = Pt(styling.get("sub_heading_size", 9.2))
                r_loc.font.name = font_family

                bullets = role.get("bullets", [])
                for b in bullets:
                    add_bullet(b)

        elif sec_type == "education":
            items = sec.get("items", [])
            for item in items:
                p1 = doc.add_paragraph()
                p1.paragraph_format.space_before = Pt(role_cfg["space_before"])
                p1.paragraph_format.space_after = Pt(role_cfg["space_after_title"])
                p1.paragraph_format.line_spacing = 1.0
                p1.paragraph_format.keep_with_next = True
                p1.paragraph_format.tab_stops.add_tab_stop(Inches(right_tab), WD_TAB_ALIGNMENT.RIGHT)

                r_deg = p1.add_run(item.get("degree", ""))
                r_deg.bold = True
                r_deg.font.size = Pt(styling.get("role_title_size", 9.8))
                r_deg.font.name = font_family
                p1.add_run("\t")
                r_dates = p1.add_run(item.get("dates", ""))
                r_dates.bold = True
                r_dates.font.size = Pt(styling.get("role_title_size", 9.8))
                r_dates.font.name = font_family

                p2 = doc.add_paragraph()
                p2.paragraph_format.space_before = Pt(0)
                p2.paragraph_format.space_after = Pt(role_cfg["space_after_sub"])
                p2.paragraph_format.line_spacing = 1.0
                p2.paragraph_format.keep_with_next = True
                p2.paragraph_format.tab_stops.add_tab_stop(Inches(right_tab), WD_TAB_ALIGNMENT.RIGHT)

                r_inst = p2.add_run(item.get("institution", ""))
                r_inst.italic = True
                r_inst.font.size = Pt(styling.get("sub_heading_size", 9.2))
                r_inst.font.name = font_family
                p2.add_run("\t")
                r_loc = p2.add_run(item.get("location", ""))
                r_loc.italic = True
                r_loc.font.size = Pt(styling.get("sub_heading_size", 9.2))
                r_loc.font.name = font_family

                coursework = item.get("coursework")
                if coursework:
                    p3 = doc.add_paragraph()
                    p3.paragraph_format.space_before = Pt(0)
                    p3.paragraph_format.space_after = Pt(bullet_cfg["space_after"])
                    p3.paragraph_format.line_spacing = 1.0
                    r_cw_label = p3.add_run("Relevant Coursework: ")
                    r_cw_label.bold = True
                    r_cw_label.font.size = Pt(body_font_size)
                    r_cw_label.font.name = font_family
                    r_cw_val = p3.add_run(coursework)
                    r_cw_val.font.size = Pt(body_font_size)
                    r_cw_val.font.name = font_family

        elif sec_type == "projects":
            projects = sec.get("projects", [])
            for proj in projects:
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(role_cfg["space_before"])
                p.paragraph_format.space_after = Pt(role_cfg["space_after_sub"])
                p.paragraph_format.line_spacing = 1.0
                p.paragraph_format.keep_with_next = True
                p.paragraph_format.tab_stops.add_tab_stop(Inches(right_tab), WD_TAB_ALIGNMENT.RIGHT)

                r_name = p.add_run(proj.get("name", ""))
                r_name.bold = True
                r_name.font.size = Pt(styling.get("role_title_size", 9.8))
                r_name.font.name = font_family

                stack = proj.get("stack")
                if stack:
                    r_bar = p.add_run("  |  ")
                    r_bar.font.size = Pt(body_font_size)
                    r_bar.font.name = font_family
                    r_stack = p.add_run(stack)
                    r_stack.italic = True
                    r_stack.font.size = Pt(styling.get("sub_heading_size", 9.0))
                    r_stack.font.name = font_family

                link_text = proj.get("link_text")
                link_url = proj.get("link_url", link_text)
                if link_text:
                    p.add_run("\t")
                    target_url = link_url if link_url.startswith("http") else f"https://{link_url}"
                    add_hyperlink(p, target_url, link_text, font_family=font_family,
                                  font_size=styling.get("sub_heading_size", 9.0))

                bullets = proj.get("bullets", [])
                for b in bullets:
                    add_bullet(b)

        elif sec_type in ("technical_skills", "skills"):
            categories = sec.get("categories", [])
            for cat in categories:
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(2.0)
                p.paragraph_format.line_spacing = 1.0
                r_label = p.add_run(cat.get("label", "") + ": ")
                r_label.bold = True
                r_label.font.size = Pt(body_font_size)
                r_label.font.name = font_family
                r_items = p.add_run(cat.get("items", ""))
                r_items.font.size = Pt(body_font_size)
                r_items.font.name = font_family

        elif sec_type == "achievements":
            bullets = sec.get("bullets", [])
            for b in bullets:
                add_bullet(b)

    doc.save(output_docx_path)
    print(f"-> Successfully generated DOCX: {output_docx_path}")


# ---------------------------------------------------------------------------
# Quality Audit & Vocabulary Verification
# ---------------------------------------------------------------------------
def audit_resume_vocabulary(docx_path: str):
    """Enforces strict anti-buzzword constraints and duplicate opening action verb checks."""
    doc = Document(docx_path)
    bullet_verbs = []

    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue

        # Check for banned buzzwords
        t_lower = t.lower()
        for buzz in BANNED_BUZZWORDS:
            if buzz in t_lower:
                # Allow technical terms like 'dynamic programming'
                if buzz == 'dynamic' and any(tech in t_lower for tech in ['dynamic programming', 'dynamically typed']):
                    continue
                raise ValueError(f"Quality QA Failed: Banned buzzword '{buzz}' detected in: '{t}'")

        # Check opening action verbs on bullets
        if t.startswith("•"):
            is_skill_or_course = any(k in t_lower for k in [
                'languages:', 'backend &', 'messaging &', 'databases &',
                'devops,', 'devops &', 'methodology &', 'architecture &', 'security &',
                'coursework:', 'relevant coursework:', 'frontend &', 'ai & developer tools:',
                'systems &', 'distributed &', 'cloud &', 'languages &',
                'data analytics &', 'ai /', 'devops, testing &', 'testing &',
                'observability &', 'ai &', 'agentic &', 'computer vision &'
            ])
            if not is_skill_or_course:
                clean_text = t.lstrip("•").strip()
                words = clean_text.split()
                if words:
                    first_word = words[0].rstrip(":,.;")
                    bullet_verbs.append(first_word)

    # Check for duplicates
    verb_counts = collections.Counter(bullet_verbs)
    duplicates = {v: c for v, c in verb_counts.items() if c > 1}
    if duplicates:
        raise ValueError(f"Quality QA Failed: Duplicate opening action verbs detected: {duplicates}")

    print(f"-> Vocabulary Audit Passed: Zero banned buzzwords, {len(bullet_verbs)} distinct opening verbs.")


# ---------------------------------------------------------------------------
# Headless PDF Export Engine (LibreOffice CLI + Win32 Fallback)
# ---------------------------------------------------------------------------
def export_to_pdf(docx_path: str, pdf_path: str):
    """Converts DOCX to PDF using headless LibreOffice (Linux VM native) with Windows fallback."""
    docx_path = os.path.abspath(docx_path)
    pdf_path = os.path.abspath(pdf_path)
    out_dir = os.path.dirname(pdf_path)

    # 1. Primary: Headless LibreOffice (standard on Gemini Linux VM)
    try:
        cmd = [
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", out_dir,
            docx_path
        ]
        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, check=True)
        stem_pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
        if stem_pdf != pdf_path and os.path.exists(stem_pdf):
            os.replace(stem_pdf, pdf_path)
        print(f"-> Successfully exported PDF via Headless LibreOffice: {pdf_path}")
        return
    except (subprocess.SubprocessError, FileNotFoundError):
        pass

    # 2. Secondary fallback for local Windows dev machines with Word installed
    try:
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            doc_word = word.Documents.Open(docx_path)
            doc_word.SaveAs(pdf_path, FileFormat=17) # 17 = wdFormatPDF
            doc_word.Close()
            print(f"-> Successfully exported PDF via MS Word COM fallback: {pdf_path}")
            return
        finally:
            word.Quit()
    except Exception as e:
        raise RuntimeError(f"PDF export failed: LibreOffice was not found and Word COM fallback failed: {e}")


# ---------------------------------------------------------------------------
# Visual QA & Page Count Validation (pypdf + pypdfium2)
# ---------------------------------------------------------------------------
def verify_and_render_qa(pdf_path: str, png_path: str):
    """Verifies strict 1-page boundary and renders image for multimodal visual inspection."""
    import pypdf
    import pypdfium2 as pdfium

    # 1. Page Count Validation
    reader = pypdf.PdfReader(pdf_path)
    page_count = len(reader.pages)
    print(f"-> PDF Page Count: {page_count}")
    if page_count != 1:
        raise ValueError(f"Visual QA Failed: Document is {page_count} pages; must strictly be exactly 1 page.")

    # 2. Render Page 1 to PNG for Gemini Multimodal Visual Inspection
    doc = pdfium.PdfDocument(pdf_path)
    img = doc[0].render(scale=2).to_pil()
    img.save(png_path)
    print(f"-> Visual QA Passed: Rendered single-page PNG saved to: {png_path}")


# ---------------------------------------------------------------------------
# CLI Entrypoint
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Autonomous Headless Résumé Generator for Yash Gupta")
    parser.add_argument("data_json", type=str, help="Path to resume_data.json configuration file")
    parser.add_argument("--docx", type=str, default=None, help="Override output .docx file path")
    parser.add_argument("--pdf", type=str, default=None, help="Override output .pdf file path")
    parser.add_argument("--png", type=str, default=None, help="Override output .png visual QA render path")
    parser.add_argument("--font", type=str, default=None, help="Override font family (e.g. 'Noto Serif', 'Caladea', 'Georgia')")
    parser.add_argument("--body-size", type=float, default=None, help="Override body font size in pt (e.g. 9.0, 9.2)")
    parser.add_argument("--skip-pdf", action="store_true", help="Skip PDF generation and visual QA")

    args = parser.parse_args()

    if not os.path.exists(args.data_json):
        print(f"Error: Specified data JSON file '{args.data_json}' does not exist.")
        sys.exit(1)

    with open(args.data_json, "r", encoding="utf-8") as f:
        data = json.load(f)

    # CLI Overrides
    if args.font:
        data.setdefault("styling", {})["font_family"] = args.font
    if args.body_size:
        data.setdefault("styling", {})["body_font_size"] = args.body_size

    meta = data.get("meta", {})
    company = meta.get("target_company", "Company")
    role = meta.get("target_role", "Software_Engineer")
    candidate_name = data.get("contact", {}).get("candidate_name", "Yash Gupta").replace(" ", "_")

    default_stem = f"{candidate_name}_{company}_{role}"
    out_docx = args.docx or meta.get("output_docx") or f"{default_stem}.docx"
    out_pdf = args.pdf or meta.get("output_pdf") or f"{default_stem}.pdf"
    out_png = args.png or meta.get("output_png") or f"{default_stem}_render.png"

    print(f"============================================================")
    print(f"Generating Résumé for {candidate_name.replace('_', ' ')}: {company} ({role})")
    print(f"Font Family: {data.get('styling', {}).get('font_family', DEFAULT_STYLING['font_family'])} | Body Size: {data.get('styling', {}).get('body_font_size', DEFAULT_STYLING['body_font_size'])}pt")
    print(f"============================================================")

    # 1. Build DOCX
    build_resume_from_data(data, out_docx)

    # 2. Audit Vocabulary & Action Verbs
    audit_resume_vocabulary(out_docx)

    # 3. Export PDF & Visual QA
    if not args.skip_pdf:
        export_to_pdf(out_docx, out_pdf)
        verify_and_render_qa(out_pdf, out_png)

    print(f"============================================================")
    print(f"Build Complete Successfully!")
    print(f"  • DOCX: {os.path.abspath(out_docx)}")
    if not args.skip_pdf:
        print(f"  • PDF:  {os.path.abspath(out_pdf)}")
        print(f"  • PNG:  {os.path.abspath(out_png)}")
    print(f"============================================================")


if __name__ == "__main__":
    main()
