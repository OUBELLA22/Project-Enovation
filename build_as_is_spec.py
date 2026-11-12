#!/usr/bin/env python3
"""Generate the AS-IS functional specification as a Word OOXML document.

This generator intentionally uses only the Python standard library. It assembles
WordprocessingML parts directly and writes them as a DOCX ZIP package.
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

OUTPUT_NAME = "Functional_Specification_AS_IS_Wire_Harness_Routing.docx"
SHORT_TITLE = "AS-IS Wire Harness Routing Functional Specification"
DARK_BLUE = "17365D"
BLUE = "2E75B6"
LIGHT_BLUE = "D9EAF7"
PALE_BLUE = "EEF5FB"
VERY_LIGHT_BLUE = "F7FAFD"
MID_GREY = "666666"
BORDER = "A9C4DA"
WHITE = "FFFFFF"


def run(
    text: str,
    *,
    bold: bool = False,
    italic: bool = False,
    size: int | None = None,
    color: str | None = None,
    underline: bool = False,
) -> str:
    """Return a safely escaped WordprocessingML run."""
    props = []
    if bold:
        props.append("<w:b/>")
    if italic:
        props.append("<w:i/>")
    if underline:
        props.append('<w:u w:val="single"/>')
    if size is not None:
        props.append(f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>')
    if color is not None:
        props.append(f'<w:color w:val="{color}"/>')
    rpr = f"<w:rPr>{''.join(props)}</w:rPr>" if props else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def paragraph(
    content: str | list[str],
    *,
    style: str | None = None,
    align: str | None = None,
    before: int = 0,
    after: int = 120,
    keep_next: bool = False,
    page_break_before: bool = False,
    left: int | None = None,
) -> str:
    """Return a paragraph from plain text or prebuilt runs."""
    runs = [run(content)] if isinstance(content, str) else content
    props = []
    if style:
        props.append(f'<w:pStyle w:val="{style}"/>')
    props.append(f'<w:spacing w:before="{before}" w:after="{after}" w:line="276" w:lineRule="auto"/>')
    if align:
        props.append(f'<w:jc w:val="{align}"/>')
    if keep_next:
        props.append("<w:keepNext/>")
    if page_break_before:
        props.append("<w:pageBreakBefore/>")
    if left is not None:
        props.append(f'<w:ind w:left="{left}"/>')
    return f"<w:p><w:pPr>{''.join(props)}</w:pPr>{''.join(runs)}</w:p>"


def heading(text: str, level: int = 1, *, page_break_before: bool = False) -> str:
    return paragraph(
        text,
        style=f"Heading{level}",
        before=220 if level == 1 else 160,
        after=100,
        keep_next=True,
        page_break_before=page_break_before,
    )


def bullet(text: str, *, level: int = 0) -> str:
    return (
        '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>'
        f'<w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="1"/></w:numPr>'
        '<w:spacing w:after="55" w:line="276" w:lineRule="auto"/>'
        f'<w:ind w:left="{720 + 360 * level}" w:hanging="360"/></w:pPr>'
        f"{run(text)}</w:p>"
    )


def page_break() -> str:
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def cell(
    text: str,
    width: int,
    *,
    bold: bool = False,
    shade: str | None = None,
    color: str | None = None,
    align: str | None = None,
) -> str:
    shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>' if shade else ""
    tcpr = (
        f'<w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{shd}'
        '<w:tcMar><w:top w:w="90" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>'
        '<w:bottom w:w="90" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>'
        '<w:vAlign w:val="center"/></w:tcPr>'
    )
    p = paragraph([run(text, bold=bold, color=color)], align=align, after=20)
    return f"<w:tc>{tcpr}{p}</w:tc>"


def table(headers: list[str], rows: list[list[str]], widths: list[int]) -> str:
    """Return a banded table with a repeating dark-blue header."""
    if len(headers) != len(widths):
        raise ValueError("Header and width counts differ")
    if any(len(row) != len(headers) for row in rows):
        raise ValueError("A table row has the wrong number of cells")
    grid = "<w:tblGrid>" + "".join(f'<w:gridCol w:w="{w}"/>' for w in widths) + "</w:tblGrid>"
    tblpr = (
        '<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>'
        '<w:tblLayout w:type="fixed"/><w:tblBorders>'
        f'<w:top w:val="single" w:sz="6" w:color="{BORDER}"/>'
        f'<w:left w:val="single" w:sz="6" w:color="{BORDER}"/>'
        f'<w:bottom w:val="single" w:sz="6" w:color="{BORDER}"/>'
        f'<w:right w:val="single" w:sz="6" w:color="{BORDER}"/>'
        f'<w:insideH w:val="single" w:sz="4" w:color="{BORDER}"/>'
        f'<w:insideV w:val="single" w:sz="4" w:color="{BORDER}"/>'
        '</w:tblBorders></w:tblPr>'
    )
    out = [f"<w:tbl>{tblpr}{grid}", '<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>']
    out.extend(cell(value, widths[index], bold=True, shade=DARK_BLUE, color=WHITE) for index, value in enumerate(headers))
    out.append("</w:tr>")
    for row_index, values in enumerate(rows):
        shade = VERY_LIGHT_BLUE if row_index % 2 == 0 else PALE_BLUE
        out.append('<w:tr><w:trPr><w:cantSplit/></w:trPr>')
        out.extend(cell(value, widths[index], shade=shade) for index, value in enumerate(values))
        out.append("</w:tr>")
    out.append("</w:tbl>")
    out.append(paragraph("", after=110))
    return "".join(out)


def note_box(title: str, text: str, *, shade: str = LIGHT_BLUE) -> str:
    """Return a single-cell callout table."""
    return table([title], [[text]], [9500]).replace(f'w:fill="{DARK_BLUE}"', f'w:fill="{DARK_BLUE}"', 1).replace(
        f'w:fill="{VERY_LIGHT_BLUE}"', f'w:fill="{shade}"', 1
    )


def label_paragraph(label: str, text: str) -> str:
    return paragraph([run(label, bold=True, color=DARK_BLUE), run(text)])


def build_body() -> str:
    body: list[str] = []

    # Cover page
    body.append(paragraph([run("FUNCTIONAL SPECIFICATION DOCUMENT", bold=True, size=40, color=DARK_BLUE)], align="center", before=800, after=100))
    body.append(paragraph([run("AS-IS Process Discovery and Analysis", bold=True, size=34, color=BLUE)], align="center", after=40))
    body.append(paragraph([run("For Automation Purposes", bold=True, size=30, color=BLUE)], align="center", after=220))
    body.append(paragraph([run("Electrical Harness Design — Wire Harness Routing in CATIA V5", size=25, color=MID_GREY)], align="center", after=520))
    body.append(table(
        ["Document field", "Value"],
        [
            ["Unit / Process", "Electrical Harness Design — Wire Harness Routing in CATIA V5"],
            ["Direction", "To be confirmed"],
            ["Department", "To be confirmed"],
            ["Date", "12 November 2026"],
            ["Version", "1.0 (Draft for business validation)"],
            ["Prepared by", "To be confirmed"],
            ["Validated by", "To be confirmed"],
        ],
        [3000, 6500],
    ))
    body.append(note_box(
        "DRAFT — BUSINESS VALIDATION REQUIRED",
        "This document combines confirmed target requirements from the existing project specification with a working hypothesis of the current manual or semi-manual CATIA V5 process. The AS-IS description is not a confirmed statement of current operations. Business interviews, observation, sample review and role validation remain required before approval.",
        shade="FFF2CC",
    ))
    body.append(heading("Document status legend", 2))
    body.append(table(
        ["Status label", "Meaning"],
        [
            ["Confirmed from source specification", "Information explicitly stated in the existing wire harness routing specification."],
            ["Working assumption — business validation required", "A plausible AS-IS description inferred for discovery planning; it must be validated by business and engineering teams."],
            ["To be confirmed", "Information is unavailable or insufficiently evidenced; the item and proposed validation action are stated."],
        ],
        [3900, 5600],
    ))
    body.append(page_break())

    # About and static contents
    body.append(heading("About this document", 1))
    body.append(paragraph(
        "This functional specification records an initial AS-IS process view for wire harness routing in CATIA V5 and translates the confirmed automation target into traceable business and technical requirements. It supports process discovery, business validation, solution design and acceptance planning. It does not replace interviews, direct observation, engineering-rule confirmation or formal approval."
    ))
    body.append(paragraph(
        "The existing project specification primarily describes the future automated TO-BE solution. Accordingly, every statement about the current manual or semi-manual way of working is labelled as a working assumption unless confirmed evidence exists. Unknown organisation, ownership, workload, performance, system, retention and security details remain To be confirmed."
    ))
    body.append(heading("Table of contents", 1))
    body.append(paragraph([
        run(
            "Static contents list — page numbers are intentionally omitted because pagination depends on the Word rendering environment.",
            italic=True,
            color=MID_GREY,
        )
    ]))
    toc = [
        "1. Introduction",
        "   1.1 Purpose of the document",
        "   1.2 Scope of the study",
        "   1.3 Discovery methodology",
        "   1.4 Business glossary",
        "2. Process Overview",
        "   2.1 Name and objective of the process",
        "   2.2 Organisational ownership",
        "   2.3 Process triggers",
        "   2.4 Expected outcome(s)",
        "3. Stakeholder Mapping",
        "   3.1 Roles and responsibilities (RACI)",
        "   3.2 Internal stakeholders involved",
        "   3.3 External stakeholders",
        "   3.4 Approval and decision points",
        "4. Detailed AS-IS Process Description",
        "   4.1 Overview (simplified process flow)",
        "   4.2 Step-by-step description",
        "   4.3 Sub-processes and variants",
        "   4.4 Documents and data involved",
        "5. Business Rules",
        "   5.1 Applicable business rules",
        "   5.2 Special cases and exceptions",
        "   5.3 Blocking or waiting conditions",
        "6. Current Volume and Performance",
        "   6.1 Frequency and volume of transactions",
        "   6.2 Processing time per step",
        "   6.3 Peak activity periods",
        "   6.4 Existing performance indicators",
        "7. Information Systems and Tools",
        "   7.1 List of applications/tools used",
        "   7.2 Interfaces and data exchange between systems",
        "   7.3 Current level of digital maturity",
        "8. Diagnosis and Pain Points",
        "   8.1 Pain points identified by the teams",
        "   8.2 Sources of errors or quality issues",
        "   8.3 Repetitive or low-value-added tasks",
        "9. Automation Expectations and Needs (Optional)",
        "   9.1 Priorities expressed by the business teams",
        "   9.2 Regulatory constraints or need for human control",
        "   9.3 Desired level of automation",
        "10. Technical Requirements (Optional)",
        "   10.1 Functional Requirements",
        "   10.2 Non-Functional Requirements (Optional)",
        "   10.3 Acceptance criteria",
        "11. Appendices (mandatory)",
        "   11.1 Interview reports",
        "   11.2 Screenshots of current tools (mandatory)",
        "   11.3 Examples of documents/forms used (mandatory)",
        "   11.4 Open confirmation register",
    ]
    body.extend(paragraph(item, after=35, left=360 if item.startswith("   ") else 0) for item in toc)
    body.append(page_break())

    # 1 Introduction
    body.append(heading("1. Introduction", 1))
    body.append(heading("1.1 Purpose of the document", 2))
    body.append(label_paragraph("Confirmed from source specification. ", "The project aims to automate deterministic wire harness routing between start connector A and end connector B through a user-selected fixation F1 or F2 in CATIA V5. V1 must generate candidate geometry, calculate route length, detect obstacle intersections and clearance violations, and block false success."))
    body.append(label_paragraph("AS-IS discovery objective. ", "Document how routing is currently requested, interpreted, created, checked, corrected, approved and stored; identify decisions and controls that automation must preserve; and establish missing evidence for business validation."))
    body.append(heading("1.2 Scope of the study", 2))
    body.append(table(
        ["Boundary", "Definition", "Evidence status"],
        [
            ["Start", "Receipt or identification of a routing need and availability of the relevant CATIA model and engineering inputs.", "Working assumption — business validation required; exact trigger and intake channel To be confirmed."],
            ["In scope", "Selection or identification of A, B and F1/F2; routing through the selected fixation; centerline and harness geometry creation; smoothing/direction/radius considerations; obstacle and clearance review; length measurement; correction, replacement and save/hand-off activities.", "Working assumption for AS-IS. Target functions are confirmed from source specification."],
            ["End", "A routed model is saved and made available for human engineering/quality review or downstream use, with unresolved issues communicated.", "Working assumption — business validation required; exact approval and release event To be confirmed."],
            ["V1 automation limit", "Mandatory obstacle detection and reporting are included. Automatic avoidance, automatic rerouting around obstacles, electrical connectivity, 2D formboard and complex multi-branch routing are outside V1.", "Confirmed from source specification."],
        ],
        [1500, 5500, 2500],
    ))
    body.append(heading("1.3 Discovery methodology", 2))
    body.append(bullet("Completed: desk analysis of Cahier_des_Charges_Faisceau_Routing.docx and its reproducible generator."))
    body.append(bullet("To be performed: interviews with process ownership, CATIA users, automation, IT/administration, engineering/quality approval and downstream representatives."))
    body.append(bullet("To be performed: observation of representative routing cases in CATIA V5, including F1 and F2 variants, missing references, collisions and design changes."))
    body.append(bullet("To be performed: review of sample models, naming conventions, macros, check reports, logs, approval records and operating procedures, where authorised."))
    body.append(bullet("To be performed: baseline measurement of volume, touch time, waiting time, rework, collision findings, failure modes and end-to-end turnaround time."))
    body.append(heading("1.4 Business glossary", 2))
    body.append(table(
        ["Term", "Definition", "Status / validation need"],
        [
            ["A", "Start connector or start reference for the harness route.", "Confirmed from source specification."],
            ["B", "End connector or end reference for the harness route.", "Confirmed from source specification."],
            ["F1 / F2", "Alternative fixation references; the user selects one mandatory passage point.", "Confirmed as target concept. Whether each is a point or an oriented physical clip is To be confirmed."],
            ["Centerline", "Central curve used as the guide for the harness sweep.", "Confirmed from source specification."],
            ["Obstacle", "Model geometry within the configured or explicitly selected checking perimeter.", "Confirmed as target concept; exact source and exclusions To be confirmed."],
            ["Collision", "Intersection of an obstacle with the harness control envelope, including a required-clearance violation.", "Confirmed from source specification."],
            ["Clearance", "Configurable minimum distance between the harness surface and obstacles; value must be greater than or equal to zero.", "Confirmed rule; value, units and ownership To be confirmed."],
            ["Epsilon (ε)", "Maximum permitted deviation between the centerline and selected fixation, or numerical classification tolerance.", "Confirmed concept; threshold and units To be confirmed."],
            ["GSD", "Generative Shape Design, the CATIA V5 workbench identified for sweep generation.", "Confirmed target environment; installed licence To be confirmed."],
            ["Generated geometrical set", "Dedicated CATIA container holding generated geometry so it can be replaced cleanly and kept separate from source references.", "Confirmed requirement; exact name To be confirmed."],
            ["CLEAR", "Obstacle check completed reliably with no intersection and no clearance violation.", "Confirmed from source specification."],
            ["COLLISION", "At least one obstacle intersects the control envelope or violates clearance; generation is blocked.", "Confirmed from source specification."],
            ["CHECK_FAILED", "The obstacle check did not execute or its result is unreliable; generation is blocked and must never be treated as CLEAR.", "Confirmed from source specification."],
        ],
        [1900, 4900, 2700],
    ))

    # 2 Overview
    body.append(heading("2. Process Overview", 1, page_break_before=True))
    body.append(heading("2.1 Name and objective of the process", 2))
    body.append(table(
        ["Item", "AS-IS process view", "Automated target view", "Status"],
        [
            ["Process name", "Wire harness routing in CATIA V5.", "Deterministic wire harness routing and generation in CATIA V5.", "Process wording aligned to source; ownership To be confirmed."],
            ["Objective", "Create an engineering-acceptable route from A to B through a chosen fixation while satisfying applicable geometry and obstacle constraints.", "Generate, check and report the route repeatably without altering source references.", "Target confirmed; current objective is a working assumption."],
            ["Owner", "To be confirmed.", "To be confirmed.", "Validate accountable business role and delegated authority."],
            ["AS-IS result", "A manually or semi-manually created CATIA route and geometry, corrected through review and saved for approval or downstream use.", "Not applicable.", "Working assumption — business validation required."],
            ["Target result", "Not applicable.", "Generated route via F1 or F2, clean dedicated geometry, total length, and structured CLEAR/COLLISION/CHECK_FAILED report.", "Confirmed from source specification."],
        ],
        [1600, 2900, 3000, 2000],
    ))
    body.append(heading("2.2 Organisational ownership", 2))
    body.append(paragraph("The accountable division, department, process owner and operational team are To be confirmed. The working stakeholder model in Section 3 must be validated before this document is approved."))
    body.append(heading("2.3 Process triggers", 2))
    body.append(table(
        ["Trigger", "Classification", "Validation required"],
        [
            ["User selects F1 or F2 and requests generation/regeneration.", "Known target trigger — confirmed from source specification.", "Confirm user interface and approval point."],
            ["A connector or model reference changes and routing must be updated.", "Probable AS-IS trigger — working assumption.", "Observe change-management cases and identify notification channel."],
            ["A new design or routing request is received.", "Probable AS-IS trigger — working assumption.", "Confirm request source, minimum input pack and priority rules."],
            ["A collision, clearance issue or review comment requires rework.", "Probable AS-IS trigger — working assumption.", "Confirm who records, assigns and closes the issue."],
        ],
        [4200, 2800, 2500],
    ))
    body.append(heading("2.4 Expected outcome(s)", 2))
    body.append(bullet("AS-IS working assumption: an engineer produces a route that can be reviewed and used by authorised downstream teams; evidence, status and approval mechanism are To be confirmed."))
    body.append(bullet("TO-BE confirmed: deterministic generation through the selected fixation, no duplicate generated geometry, original references unchanged, route length returned and mandatory obstacle check completed."))
    body.append(bullet("TO-BE confirmed: only CLEAR permits successful generation; COLLISION and CHECK_FAILED block success. Automatic obstacle avoidance remains outside V1."))

    # 3 Stakeholders
    body.append(heading("3. Stakeholder Mapping", 1, page_break_before=True))
    body.append(paragraph("The following role-based RACI is a proposed discovery baseline. Names, organisation, final role titles and assignments are To be confirmed. R = Responsible, A = Accountable, C = Consulted, I = Informed."))
    body.append(heading("3.1 Roles and responsibilities (RACI)", 2))
    body.append(table(
        ["Activity", "Process Owner", "Harness Design Engineer / CATIA User", "CAD Automation Engineer", "CATIA Administrator / IT", "Engineering / Quality Approver", "Downstream teams"],
        [
            ["Define process policy and priorities", "A?", "C?", "C?", "I?", "C?", "I?"],
            ["Provide model and engineering inputs", "I?", "R?", "C?", "C?", "C?", "C?"],
            ["Choose F1/F2 and engineering parameters", "I?", "R?", "C?", "I?", "C/A?", "I?"],
            ["Create or regenerate route", "I?", "R?", "C?", "I?", "I?", "I?"],
            ["Develop and maintain automation", "I?", "C?", "R?", "C?", "C?", "I?"],
            ["Maintain CATIA environment/access", "I?", "C?", "C?", "R/A?", "I?", "I?"],
            ["Review geometry and obstacle outcome", "I?", "R?", "C?", "I?", "A/R?", "C/I?"],
            ["Approve final engineering result", "A?", "R?", "I?", "I?", "A/R?", "I?"],
            ["Consume approved output", "I?", "C?", "I?", "I?", "I?", "R/I?"],
        ],
        [1900, 1100, 1700, 1400, 1300, 1300, 800],
    ))
    body.append(note_box("RACI validation note", "All question-marked assignments are working assumptions. Confirm one accountable role per activity, segregation of duties, backup roles, escalation paths and the authorised downstream recipients."))
    body.append(heading("3.2 Internal stakeholders involved", 2))
    body.append(bullet("Process Owner — accountable organisation and named owner To be confirmed."))
    body.append(bullet("Harness Design Engineer / CATIA User — expected operational user; role title and responsibilities To be confirmed."))
    body.append(bullet("CAD Automation Engineer and CATIA Administrator / IT — expected technical support roles; ownership boundaries To be confirmed."))
    body.append(bullet("Engineering / Quality Approver and authorised downstream teams — approval and consumption roles To be confirmed."))
    body.append(heading("3.3 External stakeholders", 2))
    body.append(paragraph("No external customer, supplier, partner or regulatory body is identified in the source specification. External involvement, data-sharing constraints and notification requirements are To be confirmed."))
    body.append(heading("3.4 Approval and decision points", 2))
    body.append(bullet("Human selection or confirmation of fixation F1/F2 and applicable engineering parameters before generation — expected control, To be confirmed."))
    body.append(bullet("Human confirmation of the obstacle perimeter and interpretation of COLLISION or CHECK_FAILED outcomes — expected control, To be confirmed."))
    body.append(bullet("Final engineering/quality approval before downstream release — expected control, To be confirmed."))
    body.append(bullet("No AI output or automation success status replaces an authorised engineering decision."))

    # 4 Detailed process
    body.append(heading("4. Detailed AS-IS Process Description", 1, page_break_before=True))
    body.append(note_box("AS-IS HYPOTHESIS", "The process below is a working assumption — business validation required. The source specification describes the future TO-BE solution and does not confirm the actual current sequence, actors, tools, controls or records."))
    body.append(heading("4.1 Overview (simplified process flow)", 2))
    body.append(paragraph("Receive request/model → Open and inspect CATIA model → Identify A, B, F1 and F2 → Choose fixation → Confirm constraints → Create centerline A → F → B → Smooth and apply directions/radius rules → Create sweep → Inspect route and obstacle passage → Check collisions/clearance → Measure length → Correct or replace if needed → Save and submit for review."))
    body.append(heading("4.2 Step-by-step description", 2))
    as_is_steps = [
        ["1", "Requester / Harness Design Engineer?", "Routing need, model reference", "Receive or identify the routing request and required design context.", "Request channel To be confirmed", "Work item and model identified"],
        ["2", "Harness Design Engineer / CATIA User?", "CATIA model", "Open the relevant model and verify that it is usable.", "CATIA V5; exact document type To be confirmed", "Model opened or issue raised"],
        ["3", "Harness Design Engineer / CATIA User?", "Model geometry and naming", "Identify start A, end B and candidate fixations F1/F2.", "CATIA V5", "References identified"],
        ["4", "Harness Design Engineer / CATIA User?", "Design intent", "Choose F1 or F2 based on engineering context.", "CATIA V5 and/or external guidance To be confirmed", "Selected fixation"],
        ["5", "Harness Design Engineer / CATIA User?", "Diameter, clearance, epsilon, direction and bend constraints", "Check available routing constraints and resolve missing values.", "Source of rules To be confirmed", "Usable constraints or blocked case"],
        ["6", "Harness Design Engineer / CATIA User?", "A, selected F, B", "Create a centerline manually from A through F to B.", "CATIA V5", "Candidate centerline"],
        ["7", "Harness Design Engineer / CATIA User?", "Candidate centerline and engineering constraints", "Smooth transitions and apply connector directions and bend-radius considerations.", "CATIA V5 / GSD availability To be confirmed", "Adjusted centerline"],
        ["8", "Harness Design Engineer / CATIA User?", "Adjusted centerline and positive diameter", "Create the swept harness geometry.", "CATIA V5 / GSD", "Candidate harness solid/surface"],
        ["9", "Harness Design Engineer / CATIA User?", "Candidate route and model", "Review route passage through the selected fixation and general model fit.", "CATIA V5 visual inspection and/or measurement To be confirmed", "Pass or correction need"],
        ["10", "Harness Design Engineer / reviewer?", "Harness envelope, obstacles and clearance", "Inspect collision and clearance against the applicable obstacle perimeter.", "Exact CATIA check, macro or visual method To be confirmed", "Observed issues or preliminary clear assessment"],
        ["11", "Harness Design Engineer / CATIA User?", "Final candidate centerline", "Measure and record route length where required.", "CATIA V5; recording location To be confirmed", "Route length"],
        ["12", "Harness Design Engineer / CATIA User?", "Review findings", "Correct the route, replace prior generated/manual geometry and repeat checks as necessary.", "CATIA V5", "Revised route without stale geometry"],
        ["13", "Harness Design Engineer / Approver?", "Reviewed route", "Save the model and submit, release or communicate it for approval/downstream use.", "Storage, workflow and approval system To be confirmed", "Saved model and review hand-off"],
    ]
    body.append(table(["#", "Actor", "Input", "Action", "Tool", "Output"], as_is_steps, [500, 1550, 1500, 2550, 1700, 1700]))
    body.append(heading("4.3 Sub-processes and variants", 2))
    body.append(table(
        ["Variant / exception", "Working AS-IS response", "Validation question"],
        [
            ["F1 versus F2", "Create the route through the selected alternative fixation and reassess geometry, length and obstacles.", "Who selects the fixation, using which criteria, and can approval override it?"],
            ["Missing A, B, F1 or F2", "Stop or seek a corrected model/reference; do not create an untraceable substitute.", "How is the issue recorded and who corrects the source model?"],
            ["Collision or clearance issue", "Revise geometry and repeat inspection; do not approve an unresolved route.", "Is any waiver process permitted? Who approves it?"],
            ["Design change", "Reopen the route, update or replace affected geometry, rerun checks and resubmit for approval.", "How are changes detected, versioned and communicated?"],
            ["Impossible route", "Escalate constraints or design intent rather than report success.", "What constitutes impossibility and what escalation evidence is required?"],
        ],
        [2100, 4300, 3100],
    ))
    body.append(heading("4.4 Documents and data involved", 2))
    body.append(table(
        ["Item", "Use", "Current evidence status"],
        [
            ["CATIA model containing A, B, F1/F2 and surrounding geometry", "Primary design input and output context.", "Model structure, type, naming and availability To be confirmed."],
            ["Engineering routing rules", "Diameter, bend, direction, clearance, tolerance and approval criteria.", "Values and controlled source To be confirmed."],
            ["Routing request / change notice", "Communicates need, scope, priority and due date.", "Format and system To be confirmed."],
            ["Collision / clearance evidence", "Supports review and correction.", "Current method and retention To be confirmed."],
            ["Approved CATIA result", "Provides routed geometry to authorised downstream teams.", "Approval record, storage and recipients To be confirmed."],
        ],
        [2700, 3500, 3300],
    ))

    # 5 Business rules
    body.append(heading("5. Business Rules", 1, page_break_before=True))
    body.append(paragraph("Rules marked as target are confirmed from the source specification. Any claim that the same control is already performed in the current AS-IS process remains a working assumption until observed or evidenced."))
    body.append(heading("5.1 Applicable business rules", 2))
    rules = [
        ["BR-01", "Target", "IF a route is created for selected fixation F, THEN the centerline must satisfy ||C(tF) − F|| ≤ ε.", "Configured ε is the only allowed numerical exception.", "Block if ε is missing/invalid or passage cannot be demonstrated."],
        ["BR-02", "Target", "IF sweep geometry is requested, THEN harness diameter must be strictly greater than zero.", "None confirmed.", "Reject invalid diameter before CATIA geometry operations."],
        ["BR-03", "Target", "IF obstacle clearance is evaluated, THEN clearance must be greater than or equal to zero.", "None confirmed.", "Reject invalid value; missing/default policy To be confirmed."],
        ["BR-04", "Target", "IF automation runs, THEN original connectors, fixations and parts must never be modified.", "None.", "Abort and report any attempted source-reference modification."],
        ["BR-05", "Target", "IF geometry is generated, THEN it must be placed in a dedicated geometrical set.", "Exact set name To be confirmed.", "Block or roll back if isolation cannot be maintained."],
        ["BR-06", "Target", "IF generation or regeneration is repeated, THEN prior generated geometry must be deleted or replaced without duplicates.", "None.", "Do not report success while stale or duplicate generated geometry remains."],
        ["BR-07", "Target V1", "IF a candidate route exists, THEN obstacle detection must run before any successful-generation status.", "No bypass in V1.", "Missing checker, perimeter or reliable result produces CHECK_FAILED."],
        ["BR-08", "Target V1", "IF the check is reliable and detects no intersection or clearance violation, THEN status is CLEAR.", "No inferred or visual-only CLEAR by automation.", "Anything unreliable is CHECK_FAILED."],
        ["BR-09", "Target V1", "IF any obstacle intersects the control envelope or violates clearance, THEN status is COLLISION.", "Waiver policy, if any, To be confirmed and cannot change detection result.", "Block successful generation."],
        ["BR-10", "Target V1", "IF the check cannot execute or is unreliable, THEN status is CHECK_FAILED and never CLEAR.", "None.", "Block successful generation."],
        ["BR-11", "Target V1", "IF status is COLLISION or CHECK_FAILED, THEN successful generation is blocked.", "A centerline preview may remain only if explicitly marked invalid; policy To be confirmed.", "Prevent false success and release."],
        ["BR-12", "Target", "IF obstacle checking runs, THEN the control envelope radius is diameter / 2 + clearance.", "Alternative fidelity rules To be confirmed only through approved change.", "Block if inputs or obstacle geometry are invalid."],
        ["BR-13", "AS-IS candidate", "IF the engineer changes F1/F2 or the design changes, THEN affected route geometry and checks should be revised.", "Current practice To be confirmed.", "Validate through observation and sample history."],
        ["BR-14", "Target", "IF natural language is accepted, THEN AI may produce only schema-validated JSON parameters.", "AI is optional.", "Reject invalid JSON before CATIA operations; AI cannot invent coordinates or engineering rules."],
    ]
    body.append(table(["Rule", "Basis", "IF / THEN statement", "Exceptions", "Blocker / control"], rules, [900, 1100, 3600, 1800, 2100]))
    body.append(heading("5.2 Special cases and exceptions", 2))
    body.append(table(
        ["Special case", "Required response", "Current AS-IS evidence"],
        [
            ["Missing or ambiguous A, B, F1 or F2", "Stop processing, request correction and never invent a substitute reference.", "To be confirmed through observation."],
            ["Invalid diameter, clearance or tolerance", "Reject the input before geometry creation or obstacle checking.", "Target rule confirmed; current handling To be confirmed."],
            ["COLLISION", "Block successful generation, identify affected obstacles and route the case for correction or approved escalation.", "Target rule confirmed; current resolution process To be confirmed."],
            ["CHECK_FAILED", "Block successful generation and resolve the check failure; never classify the case as CLEAR.", "Target rule confirmed; current handling To be confirmed."],
            ["Design or fixation change", "Replace affected generated geometry, rerun all checks and resubmit for approval.", "Working assumption — business validation required."],
        ],
        [2600, 4300, 2600],
    ))
    body.append(heading("5.3 Blocking or waiting conditions", 2))
    body.append(table(
        ["Condition (IF)", "Action (THEN)", "Frequency"],
        [
            ["Required reference or engineering parameter is missing", "Block the route and request the missing or corrected input.", "To be confirmed"],
            ["Obstacle perimeter is empty, invalid or unresolved", "Return CHECK_FAILED and block successful generation.", "To be confirmed"],
            ["Collision or clearance violation is detected", "Return COLLISION, identify obstacles and await correction or authorised decision.", "To be confirmed"],
            ["CATIA licence, workbench or checking capability is unavailable", "Stop the affected operation, report the cause and await environment support.", "To be confirmed"],
            ["Human engineering/quality approval is pending", "Keep the result unreleased until the authorised approval is recorded.", "To be confirmed"],
        ],
        [3500, 4700, 1300],
    ))

    # 6 Performance
    body.append(heading("6. Current Volume and Performance", 1, page_break_before=True))
    body.append(note_box("NO BASELINE DATA AVAILABLE", "No confirmed volumes, turnaround times, KPIs, performance thresholds or peak patterns are available in the source specification. All values remain To be confirmed through measured process discovery; no numerical placeholder is implied."))
    body.append(heading("6.1 Frequency and volume of transactions", 2))
    body.append(table(
        ["Measure", "Current value", "Definition to confirm", "Proposed collection method", "Owner"],
        [
            ["Routing requests / cases", "To be confirmed", "Count by agreed case boundary and period.", "Sample request records and model/change history for an agreed representative period.", "To be confirmed"],
            ["End-to-end turnaround time (TAT)", "To be confirmed", "Elapsed time from accepted request to approved/released result.", "Timestamp study across representative cases; separate active and waiting periods.", "To be confirmed"],
            ["Hands-on processing time", "To be confirmed", "Engineer touch time for creation, checking, correction and save.", "Time observation or structured work sampling with participant agreement.", "To be confirmed"],
            ["Waiting time", "To be confirmed", "Time awaiting inputs, access, decisions, correction or approval.", "Event timestamps and interview validation.", "To be confirmed"],
            ["Rework rate", "To be confirmed", "Share of cases requiring route correction after first review/check.", "Review comments and model revision history; agree denominator first.", "To be confirmed"],
            ["Collision / clearance finding rate", "To be confirmed", "Share of checked candidates with one or more findings.", "Structured observation and future check reports; avoid treating CHECK_FAILED as collision.", "To be confirmed"],
            ["CHECK_FAILED / unable-to-check rate", "To be confirmed", "Share of cases where reliable checking cannot complete.", "Record separate failure reason categories during pilot.", "To be confirmed"],
            ["Peak demand / seasonality", "To be confirmed", "Periods, programmes or milestones associated with higher demand.", "Review time-series data and interview planning roles.", "To be confirmed"],
            ["First-pass acceptance", "To be confirmed", "Share accepted without route rework after initial submission.", "Approval/review records after definition alignment.", "To be confirmed"],
            ["Automation performance baseline and target", "To be confirmed", "Runtime by model size/complexity and operation stage.", "Benchmark representative models; report generation and obstacle-check time separately.", "To be confirmed"],
        ],
        [1700, 1400, 2400, 3000, 1000],
    ))
    body.append(heading("6.2 Processing time per step", 2))
    body.append(paragraph("Average touch time and waiting time for each step in Section 4.2 are To be confirmed. Measure representative cases separately for model preparation, routing, geometry creation, collision/clearance checking, rework, save and approval."))
    body.append(heading("6.3 Peak activity periods", 2))
    body.append(paragraph("No peak period, programme milestone or seasonality pattern is confirmed. Review request and revision history over an agreed representative period and validate the result with planning and engineering roles."))
    body.append(heading("6.4 Existing performance indicators", 2))
    body.append(paragraph("No existing KPI is confirmed. Determine whether turnaround time, first-pass acceptance, rework, collision findings, CHECK_FAILED rate, backlog or model-release quality are currently measured before defining an automation baseline."))

    # 7 Systems
    body.append(heading("7. Information Systems and Tools", 1, page_break_before=True))
    body.append(heading("7.1 List of applications/tools used", 2))
    body.append(table(
        ["System / tool", "Current AS-IS status", "Target status", "Interface / maturity", "Open validation"],
        [
            ["CATIA V5", "Expected current design tool but exact usage is To be confirmed.", "Confirmed target CAD environment.", "Interactive CAD plus target automation interface.", "Exact release, architecture, licences and deployment."],
            ["Generative Shape Design (GSD)", "Availability and current use To be confirmed.", "Confirmed target workbench for sweep generation.", "CATIA workbench/API maturity depends on installed environment.", "Licence and supported geometry operations."],
            ["Windows / COM", "Current workstation and automation use To be confirmed.", "Confirmed target basis for CATIA COM automation if external automation is selected.", "Technology option; production choice not confirmed.", "Windows version, 32/64-bit compatibility, API availability."],
            ["CATPart", "To be confirmed.", "Acceptance includes reopening a CATPart, but exact operating model type is not confirmed.", "Document persistence check required.", "Actual input/output structure and naming."],
            ["CATProduct", "To be confirmed.", "Potential assembly context only; not confirmed.", "Potential CATIA document relationship.", "Whether obstacles/references are resolved from an assembly."],
            ["PLM / PDM", "To be confirmed.", "No target system confirmed.", "No confirmed interface.", "Storage, versioning, approval and access model."],
            ["Email / collaboration channel", "To be confirmed.", "No target system confirmed.", "Possible request/notification channel only.", "Actual intake and escalation channel."],
            ["Excel / local tracking", "To be confirmed.", "No target use confirmed.", "Possible manual record only.", "Whether any tracker exists and is controlled."],
            ["Existing CATIA macros", "To be confirmed.", "CATScript/VBA is a prototype option; Python COM is another option.", "Target technology choice not final.", "Inventory, ownership, support and code-signing rules."],
            ["Optional AI service", "No current use confirmed.", "Optional NL-to-validated-JSON layer only.", "Must not compute geometry or assert success; data transfer requires approval.", "Security, confidentiality, hosting and authorisation."],
        ],
        [1500, 1900, 2400, 2200, 1500],
    ))
    body.append(heading("7.2 Interfaces and data exchange between systems", 2))
    body.append(bullet("Routing engine inputs: references A/B/F1/F2, selected fixation, diameter, clearance, epsilon, connector directions/bend constraints where confirmed, and a resolved obstacle perimeter."))
    body.append(bullet("CATIA interface: read source references, generate/update isolated geometry, invoke a version-compatible obstacle-check method and persist the result without modifying source references."))
    body.append(bullet("Output: route length and structured status CLEAR, COLLISION or CHECK_FAILED, including affected obstacles and location/zone or minimum distance when available."))
    body.append(bullet("Technology maturity: concept/specification. Exact API, document type, licence and deployment design remain To be confirmed."))
    body.append(heading("7.3 Current level of digital maturity", 2))
    body.append(paragraph("The design activity is expected to be digitally executed in CATIA V5, but the current degree of manual work, macro use, structured data exchange, automatic checking and workflow integration is To be confirmed. The proposed automation is at specification stage; no production deployment is evidenced."))

    # 8 Diagnosis
    body.append(heading("8. Diagnosis and Pain Points", 1, page_break_before=True))
    body.append(paragraph("The source specification implies why automation may be useful but does not document measured AS-IS problems. The following diagnosis is therefore a set of hypotheses for interviews and observation, not confirmed findings."))
    body.append(heading("8.1 Pain points identified by the teams", 2))
    body.append(table(
        ["Pain-point hypothesis", "Potential impact", "Possible root cause", "Evidence status / validation"],
        [
            ["Repetitive manual centerline and sweep creation", "Engineer effort, slower iteration and avoidable handling variation.", "Repeated modelling operations for A → F → B.", "Working assumption; observe representative F1/F2 cases and measure touch time."],
            ["Visual or fragmented collision/clearance inspection", "Missed obstacles, uncertain clearance, late rework or conservative routing.", "No confirmed mandatory structured current check/report.", "Working assumption; inspect actual tools, reports and decision criteria."],
            ["Rework after design or fixation changes", "Repeated edits and delayed approval/downstream use.", "Dependent geometry and checks may require manual update.", "Working assumption; sample revision histories and review comments."],
            ["Duplicate or stale geometry", "Ambiguous model state, wrong geometry consumed and difficult cleanup.", "Regeneration may not replace prior geometry consistently.", "Risk inferred from explicit no-duplicate target; validate current frequency."],
            ["Inconsistent routing decisions", "Variable quality and review effort.", "Rules, tolerances or techniques may be tacit or applied differently.", "Working assumption; compare cases and controlled guidance."],
            ["Limited traceability of inputs and check outcomes", "Harder diagnosis, approval and audit reconstruction.", "Inputs, perimeter and status may not be recorded together.", "Working assumption; review current artefacts and retention practice."],
            ["Dependence on specialist CATIA expertise", "Training burden, bottlenecks and continuity risk.", "Geometry creation and judgement require experienced users.", "Working assumption; validate skill matrix and exception handling."],
            ["Missing or ambiguous references", "Blocked work or incorrect interpretation.", "Naming/model preparation rules may be incomplete.", "Target requirement addresses this risk; actual AS-IS incidence To be confirmed."],
        ],
        [2300, 2500, 2500, 2200],
    ))
    body.append(heading("8.2 Sources of errors or quality issues", 2))
    body.append(bullet("Potential source: missing, ambiguous or inconsistently named CATIA references. Incidence and examples are To be confirmed."))
    body.append(bullet("Potential source: visual-only or inconsistently configured collision and clearance checks. Current method is To be confirmed."))
    body.append(bullet("Potential source: tacit diameter, bend, direction, clearance or tolerance rules. Controlled sources and variation are To be confirmed."))
    body.append(bullet("Potential source: stale geometry after design changes or repeated generation. Actual occurrence is To be confirmed."))
    body.append(heading("8.3 Repetitive or low-value-added tasks", 2))
    body.append(bullet("Repeated creation and adjustment of centerlines and swept geometry for equivalent routing patterns — working assumption."))
    body.append(bullet("Repeated deletion/replacement of prior geometry after fixation or model changes — working assumption."))
    body.append(bullet("Manual collection of route length and collision/clearance evidence — working assumption."))
    body.append(bullet("Re-entry of parameters or recreation of checks across iterations — To be confirmed through observation."))

    # 9 Automation
    body.append(heading("9. Automation Expectations and Needs (Optional)", 1, page_break_before=True))
    body.append(heading("9.1 Priorities expressed by the business teams", 2))
    priorities = [
        ["1", "Deterministic routing", "Compute A → selected F → B repeatably from validated inputs."],
        ["1", "F1/F2 choice", "Allow explicit human selection and route through the chosen fixation within epsilon."],
        ["1", "Geometry generation", "Create centerline and swept section in a dedicated geometrical set."],
        ["1", "Clean regeneration", "Replace prior generated output without duplicates or source-reference changes."],
        ["1", "Length", "Calculate and return the generated route length."],
        ["1", "Mandatory obstacle detection", "Check every candidate using diameter / 2 + clearance, with a resolved perimeter and structured report."],
        ["1", "Fail-closed status", "Permit success only for reliable CLEAR; block COLLISION and CHECK_FAILED."],
        ["Future", "Automatic obstacle avoidance", "Automatically reroute around obstacles; explicitly outside V1."],
    ]
    body.append(table(["Priority", "Capability", "Expectation"], priorities, [1100, 2900, 5500]))
    body.append(heading("9.2 Regulatory constraints or need for human control", 2))
    body.append(paragraph("No regulatory requirement is identified in the source specification. Applicable regulations, internal engineering standards and mandatory approval controls are To be confirmed. The following human checkpoints are expected from the proposed operating model:"))
    body.append(bullet("Select or approve F1/F2 according to design intent."))
    body.append(bullet("Confirm the obstacle perimeter, exclusions and geometry fidelity."))
    body.append(bullet("Provide or approve diameter, clearance, epsilon, connector directions, bend radius and other engineering parameters."))
    body.append(bullet("Review COLLISION and CHECK_FAILED details, resolve model/input issues and decide whether an invalid preview should be retained under the approved policy."))
    body.append(bullet("Approve the final engineering result before downstream release."))
    body.append(heading("9.3 Desired level of automation", 2))
    body.append(label_paragraph("Recommended level. ", "Partial / assisted automation. The system performs deterministic generation, regeneration, measurement and mandatory checking; authorised humans retain engineering input, scope and approval decisions."))
    body.append(label_paragraph("Optional AI. ", "AI may translate natural-language intent into strictly schema-validated JSON. It must never invent coordinates, engineering rules, thresholds or model references; compute geometry; bypass obstacle checking; or declare CLEAR/success without the deterministic checker."))

    # 10 Technical requirements
    body.append(heading("10. Technical Requirements (Optional)", 1, page_break_before=True))
    body.append(heading("10.1 Functional Requirements", 2))
    functional = [
        ["FR-01", "EF-01", "Read the 3D positions of A, F1, F2 and B from the CATIA model.", "Steps 2–3", "Must"],
        ["FR-02", "EF-02", "Allow the user to select fixation F1 or F2.", "Step 4", "Must"],
        ["FR-03", "EF-03", "Calculate the route A → selected fixation → B.", "Step 6", "Must"],
        ["FR-04", "EF-04", "Ensure the generated centerline passes through the selected fixation within configured epsilon: ||C(tF) − F|| ≤ ε.", "Steps 6 and 9", "Must"],
        ["FR-05", "EF-05", "Generate harness geometry as a curve/centerline plus swept section.", "Steps 6–8", "Must"],
        ["FR-06", "EF-06", "Place all generated geometry in a dedicated geometrical set.", "Steps 6–8", "Must"],
        ["FR-07", "EF-07", "On selection or input change, delete or replace prior generated geometry without duplicates.", "Step 12", "Must"],
        ["FR-08", "EF-08", "Never modify original connectors, fixations or parts.", "All generation steps", "Must"],
        ["FR-09", "EF-09", "Show a clear error for missing references or an impossible route and prevent false success.", "Steps 2–6", "Must"],
        ["FR-10", "EF-10", "Calculate and return total route length.", "Step 11", "Must"],
        ["FR-11", "EF-11", "Acquire the obstacle perimeter from configuration or explicit user selection and identify the checked geometries.", "Step 10", "Must"],
        ["FR-12", "EF-12", "Execute obstacle checking on every candidate before any successful-generation status.", "Step 10", "Must"],
        ["FR-13", "EF-13", "Check the centerline control envelope with radius equal to harness diameter / 2 + clearance.", "Step 10", "Must"],
        ["FR-14", "EF-14", "Return a structured report with status, affected obstacles and, when available, location/zone and minimum distance.", "Steps 10 and 13", "Must"],
        ["FR-15", "EF-15", "Block successful generation for COLLISION or CHECK_FAILED and never report false success. Retain a centerline preview only if clearly marked invalid.", "Steps 10–13", "Must"],
    ]
    body.append(table(["ID", "Source", "Requirement", "Related step / process", "Priority"], functional, [900, 900, 4800, 2000, 900]))
    body.append(heading("10.2 Non-Functional Requirements (Optional)", 2))
    nfrs = [
        ["NFR-01 Determinism", "The same validated inputs and model state shall produce the same route and status.", "Repeat identical trials on representative models; allowable numerical comparison method To be confirmed.", "Threshold To be confirmed."],
        ["NFR-02 Fail-closed reliability", "Invalid inputs, obstacle data, unavailable checker or unreliable result shall produce a blocking error or CHECK_FAILED, never CLEAR.", "Fault-injection and negative acceptance tests.", "No false success."],
        ["NFR-03 Logging and traceability", "Record validated inputs, model/reference identifiers, obstacle perimeter, tolerance/clearance, check status, affected obstacles, errors and available diagnostics.", "Review structured logs against test cases.", "Format, location, access and retention To be confirmed."],
        ["NFR-04 Data security", "Apply least privilege and approved handling rules; do not transmit CAD data to an external AI service without confirmed authorisation.", "Security review and deployment-control evidence.", "Classification, hosting and policy To be confirmed."],
        ["NFR-05 Compatibility", "Operate with the approved Windows, CATIA V5 version, document type, licences and API architecture.", "Qualification matrix in target environment.", "Exact versions and 32/64-bit constraints To be confirmed."],
        ["NFR-06 Maintainability", "Separate routing logic, CATIA integration and optional AI; centralise configuration; isolate generated geometry.", "Architecture/code review and controlled regeneration test.", "Support model and ownership To be confirmed."],
        ["NFR-07 Usability", "Present actionable validation errors and unambiguous CLEAR/COLLISION/CHECK_FAILED outcomes without hiding required human decisions.", "Observed user acceptance sessions.", "Usability criteria To be confirmed."],
        ["NFR-08 Performance", "Measure route generation and obstacle-check runtime by representative model size and complexity before setting a target.", "Benchmark baseline and agreed percentile/statistic on approved test set.", "Baseline and target To be confirmed; no threshold invented."],
        ["NFR-09 Persistence", "Saved generated geometry and check status shall remain interpretable after closing and reopening the approved CATIA document.", "Reopen acceptance test.", "Storage representation To be confirmed."],
    ]
    body.append(table(["ID / category", "Requirement", "Verification approach", "Measure / target"], nfrs, [1900, 3500, 2500, 1600]))
    body.append(heading("10.3 Acceptance criteria", 2))
    acceptance = [
        ["AC-01", "Select F1.", "Route passes through F1 within configured epsilon."],
        ["AC-02", "Select F2.", "Route passes through F2 within configured epsilon."],
        ["AC-03", "Generate twice with unchanged inputs.", "No duplicate generated geometry; dedicated set contains only the intended current output."],
        ["AC-04", "Start connector A is missing.", "Clear error; generation is blocked and no invalid source model modification occurs."],
        ["AC-05", "Selected fixation is missing.", "Clear error and generation is blocked."],
        ["AC-06", "No valid route can be produced.", "Impossible-route error; no false success."],
        ["AC-07", "Valid obstacle perimeter contains no obstacle conflict.", "CLEAR only after the obstacle check actually executes reliably."],
        ["AC-08", "Obstacle directly intersects the control envelope.", "COLLISION; obstacle identified; successful generation blocked."],
        ["AC-09", "Obstacle is outside the physical harness but inside required clearance.", "COLLISION due to clearance violation; successful generation blocked."],
        ["AC-10", "Multiple obstacles conflict.", "COLLISION and all detected affected obstacles are reported."],
        ["AC-11", "Required obstacle set is empty/invalid or cannot be resolved.", "CHECK_FAILED and successful generation blocked."],
        ["AC-12", "Obstacle checker is unavailable or fails internally.", "CHECK_FAILED, never CLEAR, and successful generation blocked."],
        ["AC-13", "Distance is at or near the configured epsilon/numerical threshold.", "Classification is stable and follows the approved threshold rule."],
        ["AC-14", "Natural-language layer produces invalid JSON or invalid fields.", "Input is rejected before any CATIA operation."],
        ["AC-15", "Save, close and reopen the CATPart.", "Expected generated result and obstacle-check status remain available and interpretable."],
        ["AC-16", "Connector/model input changes and regeneration is requested.", "Route is updated, old generated geometry is replaced, obstacle checking reruns and source references remain unchanged."],
        ["AC-17", "Diameter is zero/negative or clearance is negative.", "Input is rejected before generation/checking; no false status is produced."],
        ["AC-18", "COLLISION or CHECK_FAILED candidate preview is retained under approved policy.", "Preview is clearly marked invalid and cannot be mistaken for approved/generated success."],
    ]
    body.append(table(["ID", "Scenario", "Expected result"], acceptance, [1000, 3500, 5000]))

    # 11 Appendices
    body.append(heading("11. Appendices (mandatory)", 1, page_break_before=True))
    body.append(heading("11.1 Interview reports", 2))
    body.append(table(
        ["Interview", "Participant / role", "Date", "Status", "Evidence / notes"],
        [
            ["INT-01", "Process Owner — To be confirmed", "To be confirmed", "Not conducted", "Attach approved notes and decisions."],
            ["INT-02", "Harness Design Engineer / CATIA User — To be confirmed", "To be confirmed", "Not conducted", "Observe at least representative F1/F2 and exception cases."],
            ["INT-03", "CAD Automation Engineer — To be confirmed", "To be confirmed", "Not conducted", "Confirm architecture, APIs and support model."],
            ["INT-04", "CATIA Administrator / IT — To be confirmed", "To be confirmed", "Not conducted", "Confirm environment, licences, access, logs and security."],
            ["INT-05", "Engineering / Quality Approver — To be confirmed", "To be confirmed", "Not conducted", "Confirm acceptance and waiver/escalation rules."],
            ["INT-06", "Downstream representative — To be confirmed", "To be confirmed", "Not conducted", "Confirm output use and release expectations."],
        ],
        [1000, 2900, 1400, 1400, 2800],
    ))
    body.append(heading("Interview questions", 3))
    interview_questions = [
        "What event starts and ends one routing case, and which role owns each decision?",
        "Which CATIA document types, releases, workbenches, licences, naming rules and model structures are used?",
        "How are A, B, F1 and F2 identified, and what happens when a reference is absent or ambiguous?",
        "Who chooses F1/F2, and which engineering criteria determine the choice?",
        "How are connector directions, bend radius, diameter, clearance and epsilon obtained and approved?",
        "How is the obstacle perimeter defined, and which items are included or excluded?",
        "How are intersections and clearance currently checked, recorded, reviewed and resolved?",
        "When geometry changes, how are stale outputs removed and checks repeated?",
        "What human approvals are required before release, and can any collision be waived?",
        "What volumes, touch times, waiting times, rework, collision findings and failure reasons can be measured?",
        "Where are models, evidence and logs stored, who can access them, and how long are they retained?",
        "Which downstream teams consume the output and what failure modes matter most to them?",
    ]
    body.extend(bullet(question) for question in interview_questions)

    body.append(heading("11.2 Screenshots of current tools (mandatory)", 2))
    body.append(paragraph("No screenshots were available in the source specification. Capture only authorised, non-sensitive examples and redact identifiers where required."))
    body.append(table(
        ["ID", "Screenshot to capture", "Purpose", "Status"],
        [
            ["SS-01", "CATIA tree showing A, B, F1/F2 and relevant containers", "Validate model structure and naming.", "To be confirmed / not provided"],
            ["SS-02", "Current manual centerline creation", "Validate AS-IS steps and tools.", "To be confirmed / not provided"],
            ["SS-03", "Current smoothing, direction and bend-radius setup", "Validate engineering controls.", "To be confirmed / not provided"],
            ["SS-04", "Current sweep creation and geometrical set", "Validate output type and organisation.", "To be confirmed / not provided"],
            ["SS-05", "Current collision/clearance inspection and result", "Validate check method and evidence.", "To be confirmed / not provided"],
            ["SS-06", "Missing-reference or failed-check message", "Validate error handling.", "To be confirmed / not provided"],
            ["SS-07", "Saved/reopened result and approval hand-off", "Validate persistence and release flow.", "To be confirmed / not provided"],
        ],
        [900, 3900, 2800, 1900],
    ))

    body.append(heading("11.3 Examples of documents/forms used (mandatory)", 2))
    body.append(table(
        ["Document / example", "Availability", "Use in this specification", "Action required"],
        [
            ["Cahier_des_Charges_Faisceau_Routing.docx", "Available", "Primary source for confirmed target scope, EF-01..EF-15, terminology and acceptance expectations.", "Maintain as source; do not modify for this AS-IS document."],
            ["build_cahier.py", "Available", "Reproducible source of the existing specification and OOXML reference.", "Maintain unchanged."],
            ["Representative CATIA model(s)", "To be confirmed", "Validate model structure, references, obstacles and performance.", "Provide approved anonymised or controlled samples."],
            ["Current work instruction / routing standard", "To be confirmed", "Validate actual AS-IS sequence and engineering rules.", "Identify controlled owner and revision."],
            ["Current collision/clearance evidence", "To be confirmed", "Validate check method and pain points.", "Provide representative reports/screenshots where authorised."],
            ["Request, change and approval records", "To be confirmed", "Validate triggers, TAT, rework and governance.", "Identify systems and representative sample."],
            ["Reference naming/model preparation guide", "To be confirmed", "Define reliable reference resolution.", "Obtain approved naming and exclusion rules."],
            ["Existing macro/script inventory", "To be confirmed", "Avoid duplication and assess reuse/support risks.", "Inventory authorised automation assets."],
        ],
        [2800, 1500, 3200, 2000],
    ))

    body.append(heading("11.4 Open confirmation register", 2))
    confirmations = [
        ["OCR-01", "Organisation", "Confirm Direction, Department and organisational scope.", "To be confirmed", "Open / To be confirmed"],
        ["OCR-02", "Actors and RACI", "Confirm role names, named participants, accountability, approvals and escalation.", "To be confirmed", "Open / To be confirmed"],
        ["OCR-03", "Exact AS-IS flow", "Validate start/end boundaries, sequence, variants, systems, records and controls by observation.", "To be confirmed", "Open / To be confirmed"],
        ["OCR-04", "Model type and naming", "Confirm CATPart/CATProduct context, tree structure, reference naming and generated set naming.", "To be confirmed", "Open / To be confirmed"],
        ["OCR-05", "Obstacle source", "Confirm perimeter source, inclusions, exclusions, required empty-set behaviour and geometry fidelity.", "To be confirmed", "Open / To be confirmed"],
        ["OCR-06", "Clearance and epsilon", "Confirm values, units, ownership, defaults and threshold classification rules.", "To be confirmed", "Open / To be confirmed"],
        ["OCR-07", "Collision policy / preview", "Confirm resolution/waiver policy and whether blocked candidates are deleted or retained as invalid previews.", "To be confirmed", "Open / To be confirmed"],
        ["OCR-08", "Volumes, TAT and KPIs", "Agree definitions, baseline period, data sources and measurement ownership.", "To be confirmed", "Open / To be confirmed"],
        ["OCR-09", "Performance target", "Benchmark representative models, then approve runtime measures and targets.", "To be confirmed", "Open / To be confirmed"],
        ["OCR-10", "CATIA version, licence and API", "Confirm release, GSD/interference capabilities, document type, Windows and 32/64-bit architecture.", "To be confirmed", "Open / To be confirmed"],
        ["OCR-11", "Logs", "Confirm log schema, location, access, monitoring and operational ownership.", "To be confirmed", "Open / To be confirmed"],
        ["OCR-12", "Retention", "Confirm retention periods and disposal rules for models, logs, reports and interview evidence.", "To be confirmed", "Open / To be confirmed"],
        ["OCR-13", "Security", "Confirm classification, least-privilege access, approved storage, external transfer and AI-service constraints.", "To be confirmed", "Open / To be confirmed"],
    ]
    body.append(table(["ID", "Topic", "Confirmation required", "Owner", "Status"], confirmations, [900, 1500, 4400, 1400, 1300]))
    body.append(note_box("Approval condition", "This draft should not be treated as a validated AS-IS baseline until the open confirmation register has assigned owners, evidence and decisions, and authorised business and engineering representatives have approved the resulting updates."))
    body.append(paragraph([run("— End of document —", italic=True, color=MID_GREY)], align="center", before=240, after=0))
    return "".join(body)


def build_package(output_path: Path) -> None:
    body_xml = build_body()
    sect_pr = (
        '<w:sectPr>'
        '<w:headerReference w:type="default" r:id="rId3"/>'
        '<w:footerReference w:type="default" r:id="rId4"/>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="600" w:footer="600" w:gutter="0"/>'
        '</w:sectPr>'
    )
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<w:body>{body_xml}{sect_pr}</w:body></w:document>'
    )

    header_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:p><w:pPr><w:jc w:val="right"/><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" w:color="2E75B6"/></w:pBdr></w:pPr>'
        f'{run(SHORT_TITLE, size=16, color=MID_GREY)}</w:p></w:hdr>'
    )
    footer_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:p><w:pPr><w:jc w:val="center"/><w:pBdr><w:top w:val="single" w:sz="6" w:space="4" w:color="2E75B6"/></w:pBdr></w:pPr>'
        f'{run(SHORT_TITLE + "  |  ", size=16, color=MID_GREY)}'
        '<w:r><w:rPr><w:sz w:val="16"/><w:color w:val="666666"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        '</w:p></w:ftr>'
    )

    styles_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:cs="Aptos"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:color w:val="222222"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:outlineLvl w:val="0"/><w:spacing w:before="260" w:after="120"/></w:pPr><w:rPr><w:b/><w:color w:val="{DARK_BLUE}"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:outlineLvl w:val="1"/><w:spacing w:before="220" w:after="100"/></w:pPr><w:rPr><w:b/><w:color w:val="{BLUE}"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:color w:val="{DARK_BLUE}"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="{BORDER}"/><w:left w:val="single" w:sz="4" w:color="{BORDER}"/><w:bottom w:val="single" w:sz="4" w:color="{BORDER}"/><w:right w:val="single" w:sz="4" w:color="{BORDER}"/><w:insideH w:val="single" w:sz="4" w:color="{BORDER}"/><w:insideV w:val="single" w:sz="4" w:color="{BORDER}"/></w:tblBorders></w:tblPr></w:style>
</w:styles>'''

    numbering_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0" w:multiLevelType="multilevel">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="○"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1080" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>'''

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''
    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''
    doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>'''
    core_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Functional Specification — AS-IS Wire Harness Routing</dc:title>
  <dc:subject>AS-IS process discovery and analysis for automation purposes</dc:subject>
  <dc:creator>To be confirmed</dc:creator>
  <cp:lastModifiedBy>To be confirmed</cp:lastModifiedBy>
  <cp:revision>1</cp:revision>
</cp:coreProperties>'''
    app_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Python standard library OOXML generator</Application>
  <AppVersion>1.0</AppVersion>
</Properties>'''

    parts = [
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", root_rels),
        ("word/document.xml", document_xml),
        ("word/_rels/document.xml.rels", doc_rels),
        ("word/styles.xml", styles_xml),
        ("word/numbering.xml", numbering_xml),
        ("word/header1.xml", header_xml),
        ("word/footer1.xml", footer_xml),
        ("docProps/core.xml", core_xml),
        ("docProps/app.xml", app_xml),
    ]
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w") as package:
        for name, content in parts:
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.create_system = 3
            info.external_attr = 0o100644 << 16
            package.writestr(info, content.encode("utf-8"))


def main() -> None:
    output = Path(__file__).resolve().parent / OUTPUT_NAME
    build_package(output)
    print(f"Written: {output}")


if __name__ == "__main__":
    main()
