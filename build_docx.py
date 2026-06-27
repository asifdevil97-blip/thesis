#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Convert rTKR_Thesis_Protocol.md to a Word .docx using only the stdlib.

Renders:
  * Headings / paragraphs with bold/italic/inline-code
  * Markdown tables  -> native Word tables (bordered, fixed layout)
  * Markdown lists / blockquotes / horizontal rules / formula code blocks
  * The patient-flow code block (Section 5.10) -> a real boxed flowchart (Figure 6)
  * A self-made two-stage PJI revision algorithm flowchart (Figure 5)
  * Image-placeholder frames + captions for the external figures (1-4, 7)
"""
import re, os, zipfile, datetime

SRC = "/projects/sandbox/rTKR_Thesis_Protocol.md"
OUT = "/projects/sandbox/rTKR_Thesis_Protocol.docx"

PAGE_W = 9026  # usable twips (A4, 1-inch margins)

_FORCE_ALIGN = None  # when set (e.g. "center"), para() centres text — used for the cover page

def esc(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
             .replace('"', "&quot;"))

# ---------- inline run parsing (**bold**, *italic*, `code`) ----------
INLINE = re.compile(r"(\*\*.+?\*\*|`[^`]+`|\*[^*]+?\*)")

def runs(text, base=None):
    base = base or {}
    out = []
    pos = 0
    for m in INLINE.finditer(text):
        if m.start() > pos:
            out.append((text[pos:m.start()], dict(base)))
        tok = m.group(0)
        st = dict(base)
        if tok.startswith("**"):
            st["b"] = True; out.append((tok[2:-2], st))
        elif tok.startswith("`"):
            st["mono"] = True; out.append((tok[1:-1], st))
        else:
            st["i"] = True; out.append((tok[1:-1], st))
        pos = m.end()
    if pos < len(text):
        out.append((text[pos:], dict(base)))
    return out or [("", dict(base))]

def run_xml(text, st):
    rpr = []
    if st.get("b"): rpr.append("<w:b/>")
    if st.get("i"): rpr.append("<w:i/>")
    if st.get("mono"):
        rpr.append('<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="18"/>')
    if st.get("color"): rpr.append('<w:color w:val="%s"/>' % st["color"])
    if st.get("sz"): rpr.append('<w:sz w:val="%d"/>' % st["sz"])
    rprx = "<w:rPr>%s</w:rPr>" % "".join(rpr) if rpr else ""
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rprx, esc(text))

def runs_xml(text, base=None):
    return "".join(run_xml(t, s) for t, s in runs(text, base))

def para(text="", style=None, ppr_extra="", base=None):
    ppr = []
    if style: ppr.append('<w:pStyle w:val="%s"/>' % style)
    if _FORCE_ALIGN and "w:jc" not in ppr_extra:
        ppr.append('<w:jc w:val="%s"/>' % _FORCE_ALIGN)
    ppr.append(ppr_extra)
    pprx = "<w:pPr>%s</w:pPr>" % "".join(ppr) if any(ppr) else ""
    return "<w:p>%s%s</w:p>" % (pprx, runs_xml(text, base) if text else "")

def raw_para(runs_inner, ppr_extra=""):
    pprx = "<w:pPr>%s</w:pPr>" % ppr_extra if ppr_extra else ""
    return "<w:p>%s%s</w:p>" % (pprx, runs_inner)

# ---------- tables ----------
def cell(text, w, *, bold=False, shade=None, borders=True, align=None, valign="center"):
    tcpr = ['<w:tcW w:w="%d" w:type="dxa"/>' % w]
    if shade: tcpr.append('<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % shade)
    if borders is True:
        tcpr.append('<w:tcBorders>' + "".join(
            '<w:%s w:val="single" w:sz="6" w:space="0" w:color="999999"/>' % s
            for s in ("top","left","bottom","right")) + '</w:tcBorders>')
    elif borders == "none":
        tcpr.append('<w:tcBorders>' + "".join(
            '<w:%s w:val="nil"/>' % s for s in ("top","left","bottom","right")) + '</w:tcBorders>')
    tcpr.append('<w:vAlign w:val="%s"/>' % valign)
    jc = '<w:jc w:val="%s"/>' % align if align else ""
    base = {"b": True} if bold else None
    p = '<w:p><w:pPr>%s</w:pPr>%s</w:p>' % (jc, runs_xml(text, base))
    return '<w:tc><w:tcPr>%s</w:tcPr>%s</w:tc>' % ("".join(tcpr), p)

def md_table(header, rows):
    n = len(header)
    cw = PAGE_W // n
    grid = "".join('<w:gridCol w:w="%d"/>' % cw for _ in range(n))
    out = ['<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/>' % PAGE_W,
           '<w:tblLayout w:type="fixed"/>',
           '<w:tblBorders>' + "".join(
               '<w:%s w:val="single" w:sz="6" w:space="0" w:color="999999"/>' % s
               for s in ("top","left","bottom","right","insideH","insideV")) + '</w:tblBorders>',
           '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>',
           '</w:tblPr><w:tblGrid>%s</w:tblGrid>' % grid]
    # header
    out.append("<w:tr>" + "".join(
        cell(c, cw, bold=True, shade="D9E2F3") for c in header) + "</w:tr>")
    for r in rows:
        r = (r + [""] * n)[:n]
        out.append("<w:tr>" + "".join(cell(c, cw) for c in r) + "</w:tr>")
    out.append("</w:tbl>")
    out.append(para())  # required trailing paragraph
    return "".join(out)

# ---------- flowchart ----------
def flow_box(text, w, *, shade="EAF1FB", bold=True):
    return ('<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:jc w:val="center"/>'
            '<w:tblLayout w:type="fixed"/>'
            '<w:tblBorders>%s</w:tblBorders></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid>'
            '<w:tr>%s</w:tr></w:tbl>') % (
        w,
        "".join('<w:%s w:val="single" w:sz="10" w:space="0" w:color="2E5496"/>' % s
                for s in ("top","left","bottom","right")),
        w,
        cell(text, w, bold=bold, shade=shade, borders=False, align="center"))

def flow_arrow():
    return raw_para(run_xml("\u2193", {"sz": 28, "color": "2E5496", "b": True}),
                    '<w:jc w:val="center"/><w:spacing w:before="20" w:after="20"/>')

def flowchart(nodes):
    """nodes: list of dicts {text, kind: box|decision|note}."""
    out = []
    w = int(PAGE_W * 0.82)
    first = True
    for nd in nodes:
        if not first and nd.get("kind") != "note_inline":
            out.append(flow_arrow())
        first = False
        kind = nd.get("kind", "box")
        if kind == "decision":
            out.append(flow_box(nd["text"], w, shade="FCE4D6"))
        elif kind == "note":
            out.append(raw_para(run_xml(nd["text"], {"i": True, "sz": 18, "color": "767171"}),
                                '<w:jc w:val="center"/>'))
        else:
            out.append(flow_box(nd["text"], w))
    out.append(para())
    return "".join(out)

# ---------- image placeholder ----------
def image_placeholder(caption, w=None):
    w = w or int(PAGE_W * 0.7)
    box = ('<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:jc w:val="center"/>'
           '<w:tblLayout w:type="fixed"/><w:tblBorders>%s</w:tblBorders></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid>'
           '<w:tr><w:trPr><w:trHeight w:val="1100"/></w:trPr>%s</w:tr></w:tbl>') % (
        w,
        "".join('<w:%s w:val="dashed" w:sz="8" w:space="0" w:color="A6A6A6"/>' % s
                for s in ("top","left","bottom","right")),
        w,
        cell("[ Insert image here ]", w, shade="F2F2F2", borders=False,
             align="center", valign="center"))
    cap = raw_para(run_xml(caption, {"i": True, "sz": 18, "color": "404040"}),
                   '<w:jc w:val="center"/><w:spacing w:after="160"/>')
    return box + cap

# ---------- markdown driver ----------
HEAD = {1: "Title", 2: "Heading1", 3: "Heading2", 4: "Heading3"}
FIG_RE = re.compile(r"^-\s*\*\*Figure (\d+):\*\*\s*(.*)$")

def is_sep(line):
    s = line.strip().strip("|")
    return bool(s) and "-" in s and all(ch in "-:| " for ch in s)

def split_row(line):
    return [c.strip() for c in line.strip().strip("|").split("|")]

def two_stage_flowchart():
    return flowchart([
        {"text": "Confirmed / suspected PJI (2018 ICM\u2013MSIS criteria)", "kind": "decision"},
        {"text": "STAGE 1: Implant removal + radical debridement; insert antibiotic-loaded cement spacer (static / articulating); send \u22653 deep tissue cultures + histology"},
        {"text": "Organism-directed IV antibiotics \u2248 6 weeks (microbiology / ID guidance)"},
        {"text": "Antibiotic-free interval; serial ESR & CRP trend; repeat aspiration if indicated"},
        {"text": "Infection controlled? (normalising ESR/CRP, negative aspirate)", "kind": "decision"},
        {"text": "If NO \u2192 repeat debridement / spacer exchange, adjust antibiotics, then re-evaluate", "kind": "note"},
        {"text": "If YES \u2192 STAGE 2: reimplant revision prosthesis (appropriate constraint \u00b1 stems / augments / cones)"},
        {"text": "Postoperative rehabilitation and clinico-radiological surveillance"},
    ])

def page_break():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

def process(lines, page_break_sections=False):
    body = []
    i = 0
    n = len(lines)
    section_seen = False
    while i < n:
        line = lines[i]
        s = line.strip()

        # fenced code block
        if s.startswith("```"):
            j = i + 1
            buf = []
            while j < n and not lines[j].strip().startswith("```"):
                buf.append(lines[j]); j += 1
            block = "\n".join(buf)
            if "Assessed for Eligibility" in block:  # patient-flow -> CONSORT flowchart
                body.append(flowchart([
                    {"text": "All patients undergoing Revision TKR at KEM Hospital (Jan 2021 \u2013 Dec 2027)"},
                    {"text": "Assessed for eligibility"},
                    {"text": "EXCLUDED \u2192 incomplete records; <6 months follow-up; declined consent; mega-prosthesis cases; isolated liner exchange", "kind": "note"},
                    {"text": "Enrolled / final cohort (recruit n = 38; target \u2265 30 evaluable)"},
                    {"text": "Retrospective arm (Jan 2021 \u2013 Jun 2026)  +  Prospective arm (Aug 2026 \u2013 Dec 2027)"},
                    {"text": "Outcome assessment \u2014 Clinical (KSS, WOMAC, VAS, ROM), Radiological (KS zones, alignment), Patient satisfaction. Time points: 6 wks, 3 mo, 6 mo, 1 yr, latest"},
                    {"text": "Statistical analysis (SPSS v26)"},
                    {"text": "Results & conclusions"},
                ]))
            else:  # formula / preformatted -> monospace shaded lines
                for bl in buf:
                    body.append(raw_para(
                        run_xml(bl if bl else "\u00a0", {"mono": True}),
                        '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>'
                        '<w:spacing w:after="0"/>'))
                body.append(para())
            i = j + 1
            continue

        # table
        if s.startswith("|") and i + 1 < n and is_sep(lines[i + 1]):
            header = split_row(line)
            rows = []
            j = i + 2
            while j < n and lines[j].strip().startswith("|"):
                rows.append(split_row(lines[j])); j += 1
            body.append(md_table(header, rows))
            i = j
            continue

        # headings
        m = re.match(r"^(#{1,4})\s+(.*)$", s)
        if m:
            lvl = len(m.group(1))
            extra = ""
            if page_break_sections and lvl == 2:
                if section_seen:
                    extra = '<w:pageBreakBefore/>'
                section_seen = True
            body.append(para(m.group(2), HEAD[lvl], ppr_extra=extra))
            i += 1
            continue

        # horizontal rule
        if s == "---":
            body.append(raw_para("", '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="BFBFBF"/></w:pBdr>'))
            i += 1
            continue

        # blockquote
        if s.startswith(">"):
            qt = s.lstrip(">").strip()
            body.append(para(qt, "Quote"))
            i += 1
            continue

        # figure list item (special handling)
        fm = FIG_RE.match(s)
        if fm:
            num = int(fm.group(1)); cap = fm.group(2)
            body.append(para("Figure %d. %s" % (num, re.sub(r"\*", "", cap)),
                             ppr_extra='<w:spacing w:before="120" w:after="60"/>',
                             base={"b": True}))
            if num == 6:
                body.append(para("(Study design flowchart \u2014 rendered in Section 5.10 above.)",
                                 base={"i": True}))
            elif num == 5:
                body.append(two_stage_flowchart())
            else:
                body.append(image_placeholder("Figure %d image to be inserted." % num))
            i += 1
            continue

        # bullet / numbered list
        lm = re.match(r"^(\s*)([-*]|\d+\.)\s+(.*)$", line)
        if lm:
            indent = 360
            prefix = "\u2022  " if lm.group(2) in ("-", "*") else (lm.group(2) + "  ")
            body.append(raw_para(runs_xml(prefix + lm.group(3)),
                                 '<w:ind w:left="%d" w:hanging="280"/><w:spacing w:after="40"/>' % indent))
            i += 1
            continue

        # blank line
        if s == "":
            i += 1
            continue

        # normal paragraph
        body.append(para(s))
        i += 1

    return "".join(body)

def convert(md):
    """Render cover page (centred) + title page + body with page breaks between sections."""
    global _FORCE_ALIGN
    idx = md.find("\n## INDEX")
    if idx == -1:
        return process(md.split("\n"), page_break_sections=True)
    front = md[:idx]
    body_md = md[idx + 1:]
    tp = front.find("### TITLE PAGE")
    if tp == -1:
        cover_md, titlepage_md = front, ""
    else:
        cover_md, titlepage_md = front[:tp], front[tp:]

    out = []
    # Cover page: centre-aligned
    _FORCE_ALIGN = "center"
    out.append('<w:p/>')  # a little top spacing
    out.append(process(cover_md.split("\n")))
    _FORCE_ALIGN = None
    out.append(page_break())
    # Title page (left-aligned tables)
    if titlepage_md.strip():
        out.append(process(titlepage_md.split("\n")))
        out.append(page_break())
    # Main body with a page break before each numbered section
    out.append(process(body_md.split("\n"), page_break_sections=True))
    return "".join(out)

# ---------- package ----------
def styles_xml():
    def style(sid, name, props, ppr="", default=False):
        return ('<w:style w:type="paragraph" %sw:styleId="%s"><w:name w:val="%s"/>'
                '<w:qFormat/><w:pPr>%s</w:pPr><w:rPr>%s</w:rPr></w:style>') % (
            'w:default="1" ' if default else "", sid, name, ppr, props)
    sty = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
        '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Nirmala UI"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>',
        style("Normal", "Normal", "", default=True),
        style("Title", "Title",
              '<w:b/><w:sz w:val="36"/><w:color w:val="1F3864"/>',
              '<w:jc w:val="center"/><w:spacing w:before="120" w:after="120"/>'),
        style("Heading1", "heading 1",
              '<w:b/><w:sz w:val="32"/><w:color w:val="1F3864"/>',
              '<w:spacing w:before="320" w:after="120"/><w:keepNext/>'
              '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="2" w:color="8EAADB"/></w:pBdr>'),
        style("Heading2", "heading 2",
              '<w:b/><w:sz w:val="28"/><w:color w:val="2E5496"/>',
              '<w:spacing w:before="240" w:after="100"/><w:keepNext/>'),
        style("Heading3", "heading 3",
              '<w:b/><w:sz w:val="26"/><w:color w:val="2E5496"/>',
              '<w:spacing w:before="200" w:after="80"/><w:keepNext/>'),
        style("Heading4", "heading 4",
              '<w:b/><w:sz w:val="24"/><w:color w:val="404040"/>',
              '<w:spacing w:before="160" w:after="60"/><w:keepNext/>'),
        style("Quote", "Quote",
              '<w:sz w:val="22"/><w:color w:val="404040"/>',
              '<w:ind w:left="340"/><w:spacing w:before="80" w:after="160"/>'
              '<w:shd w:val="clear" w:color="auto" w:fill="F7F7F0"/>'
              '<w:pBdr><w:left w:val="single" w:sz="18" w:space="6" w:color="C9A227"/></w:pBdr>'),
        '</w:styles>',
    ]
    return "".join(sty)

def document_xml(body):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<w:body>%s'
            '<w:sectPr>'
            '<w:footerReference w:type="default" r:id="rId2"/>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
            '</w:body></w:document>') % body

def footer_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t>1</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
            '</w:p></w:ftr>')

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
      '</Types>')

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '</Relationships>')

DOC_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
            '</Relationships>')

def core_xml():
    now = datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:title>Revision TKR Thesis Protocol</dc:title>'
            '<dc:creator>Dr. Asif Ahmed</dc:creator>'
            '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
            '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
            '</cp:coreProperties>') % (now, now)

def main():
    with open(SRC, encoding="utf-8") as f:
        md = f.read()
    body = convert(md)
    doc = document_xml(body)
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", doc)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", styles_xml())
        z.writestr("word/footer1.xml", footer_xml())
        z.writestr("docProps/core.xml", core_xml())
    print("Wrote", OUT, "(%d bytes)" % os.path.getsize(OUT))

if __name__ == "__main__":
    main()
