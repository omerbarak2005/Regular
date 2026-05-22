#!/usr/bin/env python3
"""
M&A Origination Pipeline Builder
Populates Jim's workbook Pipeline tab with researched seed target contacts.
Usage: python populate_pipeline.py [--output MA_Pipeline_Populated.xlsx]
"""

import argparse
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

TODAY = date.today().strftime("%Y-%m-%d")

# ── Colors ──────────────────────────────────────────────────────────────────
NAV      = "1F3864"
WHITE    = "FFFFFF"
GREEN_BG = "E2EFDA"
GREEN_FG = "375623"
AMBER_BG = "FFF2CC"
AMBER_FG = "7F6000"
RED_BG   = "FCE4D6"
RED_FG   = "843C0C"
GRAY_BG  = "F2F2F2"
GRAY_FG  = "595959"
BLUE_BG  = "D6E4F7"

def _fill(c): return PatternFill("solid", fgColor=c)
def _font(bold=False, color="000000", size=10, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic, name="Calibri")
def _border():
    s = Side(style="thin", color="BDD7EE")
    return Border(left=s, right=s, top=s, bottom=s)
def _align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


# ── Pipeline columns (matches Jim's workbook exactly) ───────────────────────
PIPELINE_COLS = [
    ("Company Name",       20, False, "left"),
    ("Website",            28, False, "left"),
    ("State",               6, False, "center"),
    ("City/HQ",            16, False, "left"),
    ("Industry",           20, False, "left"),
    ("Est. Revenue",       14, False, "right"),
    ("Est. EBITDA",        13, False, "right"),
    ("Est. Employees",     14, False, "center"),
    ("Ownership Type",     16, False, "left"),
    ("Year Founded",       12, False, "center"),
    ("Multi-Location?",    14, False, "center"),
    ("Priority Score",     13, False, "center"),
    ("Status",             18, False, "left"),
    ("Contact Name",       20, False, "left"),
    ("Title",              22, False, "left"),
    ("Email",              34, False, "left"),
    ("Phone",              16, False, "center"),
    ("LinkedIn URL",       38, False, "left"),
    ("Last Outreach Date", 16, False, "center"),
    ("Next Step",          28, True,  "left"),
    ("Notes / Why Attractive", 44, True, "left"),
    ("Verification Needed",    22, True,  "left"),
]

# ── Researched pipeline rows ─────────────────────────────────────────────────
# Priority: 1=HOT (owner ID'd, email found, strong fit), 2=WARM, 3=COLD/verify
PIPELINE_ROWS = [
    {
        "Company Name":       "Russell Reid Waste Management",
        "Website":            "russellreid.com",
        "State":              "NJ",
        "City/HQ":            "Edison",
        "Industry":           "Waste Management / Environmental",
        "Est. Revenue":       "$30M–$60M",
        "Est. EBITDA":        "$5M–$10M",
        "Est. Employees":     "100–200",
        "Ownership Type":     "Family-owned (2nd gen)",
        "Year Founded":       "1943",
        "Multi-Location?":    "Yes (5 locations)",
        "Priority Score":     "1",
        "Status":             "Researching",
        "Contact Name":       "Gary M. Weiner",
        "Title":              "President",
        "Email":              "gweiner@russellreid.com",
        "Phone":              "(800) 356-4468",
        "LinkedIn URL":       "linkedin.com/company/russellreidresponsiblewastemanagement",
        "Last Outreach Date": "",
        "Next Step":          "Verify email; call main line to confirm Gary Weiner is decision-maker; send intro letter",
        "Notes / Why Attractive": "80+ year family business, 5 NJ/NY/PA/DE locations, septic & grease trap pumping. Morton Weiner (founder) acquired 1981; son Gary now runs it. Classic boomer transition setup. Waste services = recurring revenue + limited substitutes = PE loves this.",
        "Verification Needed": "Confirm email format; confirm no PE involvement; verify Gary is ready to discuss exit",
    },
    {
        "Company Name":       "Aiello Home Services",
        "Website":            "aiellohomeservices.com",
        "State":              "CT",
        "City/HQ":            "Windsor Locks",
        "Industry":           "HVAC / Plumbing / Electrical",
        "Est. Revenue":       "$22M",
        "Est. EBITDA":        "$3M–$4M",
        "Est. Employees":     "60",
        "Ownership Type":     "Family-owned (4th gen)",
        "Year Founded":       "~1935",
        "Multi-Location?":    "No",
        "Priority Score":     "1",
        "Status":             "Researching",
        "Contact Name":       "Michael Jezouit",
        "Title":              "President & CEO",
        "Email":              "m.jezouit@aiellohomeservices.com",
        "Phone":              "(860) 292-2600",
        "LinkedIn URL":       "linkedin.com/in/michael-jezouit-0b18a0126/",
        "Last Outreach Date": "",
        "Next Step":          "Send personalized email referencing 90-year family legacy; mention succession planning angle",
        "Notes / Why Attractive": "4th-generation family business, 90+ years. $22M revenue, 60 employees, CT market. Home services = fragmented market, PE loves roll-ups. Michael Jezouit as non-family CEO suggests owners may be open to transition. EMAIL CONFIRMED via multiple data sources (m.jezouit@aiellohomeservices.com).",
        "Verification Needed": "Confirm owners are Aiello family (not Jezouit); confirm no recent PE investment; check if any LOI outstanding",
    },
    {
        "Company Name":       "Confires Fire Protection",
        "Website":            "confires.com",
        "State":              "NJ",
        "City/HQ":            "South Plainfield",
        "Industry":           "Fire Protection / Life Safety",
        "Est. Revenue":       "$5M–$10M",
        "Est. EBITDA":        "$750K–$1.5M",
        "Est. Employees":     "33",
        "Ownership Type":     "Privately held",
        "Year Founded":       "~1985",
        "Multi-Location?":    "Yes (affiliated offices)",
        "Priority Score":     "2",
        "Status":             "Researching",
        "Contact Name":       "William Wrublevski",
        "Title":              "Owner",
        "Email":              "william.wrublevski@confires.com",
        "Phone":              "(888) 228-0917",
        "LinkedIn URL":       "linkedin.com/company/confires-fire-protection-services-llc/",
        "Last Outreach Date": "",
        "Next Step":          "Verify William Wrublevski is still active owner; also cc VP Mark Calleo (mark.calleo@confires.com)",
        "Notes / Why Attractive": "40+ years serving NJ/DE/PA. Fire protection is a recurring-revenue inspection/service business — extremely attractive to PE. Owner William Wrublevski likely retirement-age. Email format: first.last@confires.com confirmed via RocketReach.",
        "Verification Needed": "Confirm W. Wrublevski ownership stake; confirm email; check if part of any roll-up already",
    },
    {
        "Company Name":       "T.F. O'Brien & Co.",
        "Website":            "tfobrien.com",
        "State":              "NY",
        "City/HQ":            "New Hyde Park",
        "Industry":           "HVAC / Cooling & Heating",
        "Est. Revenue":       "$10M–$20M",
        "Est. EBITDA":        "$1.5M–$3M",
        "Est. Employees":     "25–35",
        "Ownership Type":     "Family-owned (3rd gen)",
        "Year Founded":       "1934",
        "Multi-Location?":    "No",
        "Priority Score":     "2",
        "Status":             "Researching",
        "Contact Name":       "Kerry O'Brien",
        "Title":              "Owner",
        "Email":              "k.obrien@tfobrien.com",
        "Phone":              "(516) 488-1800",
        "LinkedIn URL":       "linkedin.com/in/kerry-o-brien-b7906b6/",
        "Last Outreach Date": "",
        "Next Step":          "Reach out via LinkedIn + email; acknowledge 90+ year legacy; pitch confidential process",
        "Notes / Why Attractive": "90+ year Long Island HVAC institution. 3rd-gen family ownership (Kerry, Thomas Jr., Christopher O'Brien). Classic succession play — multiple heirs often means disagreement on direction, creating exit motivation. Strong brand in high-income Long Island market.",
        "Verification Needed": "Confirm current co-owner names; confirm revenue; check if any PE approached them before",
    },
    {
        "Company Name":       "Bruni & Campisi",
        "Website":            "bruniandcampisi.com",
        "State":              "NY",
        "City/HQ":            "Elmsford",
        "Industry":           "HVAC / Plumbing",
        "Est. Revenue":       "$29.9M",
        "Est. EBITDA":        "$4M–$6M",
        "Est. Employees":     "60–85",
        "Ownership Type":     "Family-owned (2nd gen)",
        "Year Founded":       "1979",
        "Multi-Location?":    "No",
        "Priority Score":     "1",
        "Status":             "Researching",
        "Contact Name":       "Keith Bruni",
        "Title":              "President",
        "Email":              "k.bruni@bruniandcampisi.com",
        "Phone":              "(914) 214-1550",
        "LinkedIn URL":       "linkedin.com/in/keith-bruni-6082b055/",
        "Last Outreach Date": "",
        "Next Step":          "Send email + LinkedIn connection; reference Westchester market strength and platform potential",
        "Notes / Why Attractive": "~$30M revenue, Westchester NY, 45+ years. HVAC+plumbing combo = higher ticket, recurring maintenance contracts. Keith Bruni (president) is 2nd gen; founders Mario Bruni and Frank Campisi are likely retirement age. Strong PE roll-up candidate.",
        "Verification Needed": "Confirm founders' current roles; confirm email format; check Westchester competitor landscape",
    },
    {
        "Company Name":       "Mowrey Elevator",
        "Website":            "mowreyelevator.com",
        "State":              "FL",
        "City/HQ":            "Marianna",
        "Industry":           "Elevator Sales, Service & Modernization",
        "Est. Revenue":       "$19.5M",
        "Est. EBITDA":        "$2.5M–$4M",
        "Est. Employees":     "50–75",
        "Ownership Type":     "Family-owned (founder-led)",
        "Year Founded":       "1976",
        "Multi-Location?":    "Yes (Marianna + Davie FL)",
        "Priority Score":     "1",
        "Status":             "Researching",
        "Contact Name":       "Tim Mowrey Sr.",
        "Title":              "Founder & President",
        "Email":              "t.mowrey@mowreyelevator.com",
        "Phone":              "(850) 526-4111",
        "LinkedIn URL":       "linkedin.com/in/tim-mowrey-29a39342/",
        "Last Outreach Date": "",
        "Next Step":          "Call + email Tim directly; elevator service is a niche PE loves — mention specific elevator roll-up comps",
        "Notes / Why Attractive": "50-year-old founder-led elevator company, $19.5M revenue, 50-75 employees, FL. Elevator service = govt-regulated, recurring maintenance contracts, very high switching costs. Mowrey branded as 'The Elevator Man' — strong regional reputation. Classic niche boomer business.",
        "Verification Needed": "Confirm Tim Sr. is still primary decision-maker vs. Tim Jr.; confirm no PE approach; verify email",
    },
]

# ── PE-acquired / disqualified ───────────────────────────────────────────────
DISQUALIFIED = [
    {
        "Company Name": "Sansone Air Conditioning",
        "Reason":       "ACQUIRED — Strikepoint Group Holdings (PE-backed), October 2022",
        "Website":      "sansone-ac.com",
        "State":        "FL",
        "Notes":        "Was family-owned (4th gen, since 1976). Acquired by PE platform Strikepoint. No longer a valid outreach target.",
    },
    {
        "Company Name": "Piper Fire Protection",
        "Reason":       "ACQUIRED — Fortis Fire & Safety (PE-backed), March 2023",
        "Website":      "piperfire.com",
        "State":        "FL",
        "Notes":        "Founded 1986 by Terry Johnson. Acquired by Fortis in March 2023. Remove from seed list.",
    },
    {
        "Company Name": "Stan's Heating Air & Plumbing",
        "Reason":       "PE-BACKED — Treaty Oak Equity (growth investment, Austin TX)",
        "Website":      "stansac.com",
        "State":        "TX",
        "Notes":        "Chris Strand owns it; received institutional investment from Treaty Oak Equity. Not a clean independent buyout target.",
    },
]


def build_workbook(output_path: str):
    wb = Workbook()

    # ── Sheet 1: Pipeline ──────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Pipeline"

    # Header
    for ci, (label, width, wrap, align) in enumerate(PIPELINE_COLS, 1):
        cell = ws.cell(row=1, column=ci, value=label)
        cell.fill      = _fill(NAV)
        cell.font      = _font(bold=True, color=WHITE, size=11)
        cell.alignment = _align(h="center", v="center")
        cell.border    = _border()
        ws.column_dimensions[get_column_letter(ci)].width = width
    ws.row_dimensions[1].height = 30

    # Data rows
    COL_NAMES = [c[0] for c in PIPELINE_COLS]
    for ri, row in enumerate(PIPELINE_ROWS, 2):
        p = row.get("Priority Score", "2")
        if p == "1":
            bg, fg = GREEN_BG, GREEN_FG
        elif p == "2":
            bg, fg = AMBER_BG, AMBER_FG
        else:
            bg, fg = GRAY_BG, GRAY_FG

        for ci, (label, _, wrap, halign) in enumerate(PIPELINE_COLS, 1):
            val  = row.get(label, "")
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.fill      = _fill(bg)
            cell.font      = _font(color=fg, size=10)
            cell.border    = _border()
            cell.alignment = _align(h=halign, v="top", wrap=wrap)

            # Email cells: blue underline style
            if label == "Email" and val:
                cell.font = Font(color="1155CC", underline="single", size=10, name="Calibri")
            # LinkedIn cells: blue
            if label == "LinkedIn URL" and val:
                cell.font = Font(color="0563C1", size=9, name="Calibri")
            # Notes: italic gray
            if label in ("Notes / Why Attractive", "Next Step"):
                cell.font = _font(italic=True, color=fg, size=9)
            # Priority badge
            if label == "Priority Score":
                badge = {"1": "375623", "2": "806000", "3": "595959"}.get(p, "595959")
                cell.fill = _fill(badge)
                cell.font = _font(bold=True, color=WHITE, size=10)
                cell.alignment = _align(h="center", v="center")

        ws.row_dimensions[ri].height = 72

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # ── Sheet 2: Disqualified (PE-backed) ─────────────────────────────────────
    dq = wb.create_sheet("Disqualified — PE Backed")
    dq_cols = [
        ("Company Name", 28),
        ("State", 8),
        ("Website", 28),
        ("Reason",  54),
        ("Notes",   54),
    ]
    for ci, (label, width) in enumerate(dq_cols, 1):
        cell = dq.cell(row=1, column=ci, value=label)
        cell.fill      = _fill("843C0C")
        cell.font      = _font(bold=True, color=WHITE, size=11)
        cell.alignment = _align(h="center", v="center")
        cell.border    = _border()
        dq.column_dimensions[get_column_letter(ci)].width = width
    dq.row_dimensions[1].height = 28

    for ri, row in enumerate(DISQUALIFIED, 2):
        for ci, (label, _) in enumerate(dq_cols, 1):
            val  = row.get(label, "")
            cell = dq.cell(row=ri, column=ci, value=val)
            cell.fill      = _fill(RED_BG)
            cell.font      = _font(color=RED_FG, size=10)
            cell.border    = _border()
            cell.alignment = _align(h="left", v="top", wrap=True)
        dq.row_dimensions[ri].height = 56

    dq.freeze_panes = "A2"

    # ── Sheet 3: Research Notes ────────────────────────────────────────────────
    rn = wb.create_sheet("Research Notes")
    rn.column_dimensions["A"].width = 26
    rn.column_dimensions["B"].width = 72

    notes = [
        ("RESEARCH NOTES",        f"Generated {TODAY} — Jim Paulson M&A Origination Project"),
        ("",                      ""),
        ("PRIORITY 1 TARGETS (send outreach this week)", ""),
        ("Russell Reid",          "President Gary M. Weiner (Morton Weiner's son). 5 locations NJ/NY/PA/DE. Email: gweiner@russellreid.com (estimated — verify). Call (800) 356-4468 to confirm direct line."),
        ("Aiello Home Services",  "Michael Jezouit, President/CEO. Email CONFIRMED: m.jezouit@aiellohomeservices.com (ZoomInfo + RocketReach). Phone: (860) 292-2600. 4th-gen family, $22M rev."),
        ("Bruni & Campisi",       "Keith Bruni, President. Email est.: k.bruni@bruniandcampisi.com. LinkedIn confirmed. Phone (914) 214-1550. $29.9M revenue, Westchester NY."),
        ("Mowrey Elevator",       "Tim Mowrey Sr., Founder/President. Email est.: t.mowrey@mowreyelevator.com. LinkedIn confirmed. Phone (850) 526-4111. $19.5M rev, FL."),
        ("",                      ""),
        ("PRIORITY 2 TARGETS (verify, then send)", ""),
        ("T.F. O'Brien",          "Kerry O'Brien, Owner. Email est.: k.obrien@tfobrien.com. LinkedIn confirmed. Phone (516) 488-1800. 90+ yr Long Island HVAC. Revenue unclear — call to verify."),
        ("Confires Fire",         "William Wrublevski, Owner. Email est.: william.wrublevski@confires.com. Email format first.last@confires.com confirmed. VP Mark Calleo: mark.calleo@confires.com."),
        ("",                      ""),
        ("REMOVED — PE-BACKED",   ""),
        ("Sansone AC",            "PASS: Acquired by Strikepoint Group Holdings (PE) October 2022. Remove from pipeline."),
        ("Piper Fire Protection", "PASS: Acquired by Fortis Fire & Safety (PE) March 2023. Remove from pipeline."),
        ("Stan's Heating",        "PASS: Received investment from Treaty Oak Equity (Austin TX). Not a clean independent target."),
        ("",                      ""),
        ("EMAIL FORMAT NOTES",    ""),
        ("Russell Reid",          "Try: gweiner@russellreid.com  |  gary.weiner@russellreid.com  |  g.weiner@russellreid.com"),
        ("T.F. O'Brien",          "Try: k.obrien@tfobrien.com  |  kobrien@tfobrien.com  |  kerry.obrien@tfobrien.com"),
        ("Mowrey Elevator",       "Try: t.mowrey@mowreyelevator.com  |  tmowrey@mowreyelevator.com  |  tim.mowrey@mowreyelevator.com"),
        ("Bruni & Campisi",       "Try: k.bruni@bruniandcampisi.com  |  kbruni@bruniandcampisi.com  |  keith.bruni@bruniandcampisi.com"),
        ("Confires",              "FORMAT CONFIRMED: first.last@confires.com  →  william.wrublevski@confires.com"),
        ("Aiello",                "EMAIL CONFIRMED: m.jezouit@aiellohomeservices.com"),
    ]

    for ri, (label, val) in enumerate(notes, 1):
        ca = rn.cell(row=ri, column=1, value=label)
        cb = rn.cell(row=ri, column=2, value=val)
        if label.startswith(("RESEARCH", "PRIORITY", "REMOVED", "EMAIL FORMAT")):
            ca.fill = _fill(NAV); ca.font = _font(bold=True, color=WHITE, size=11)
            cb.fill = _fill(NAV); cb.font = _font(bold=True, color=WHITE, size=11)
            rn.row_dimensions[ri].height = 24
        elif label in ("Russell Reid", "Aiello Home Services", "Bruni & Campisi", "Mowrey Elevator"):
            ca.fill = _fill(GREEN_BG); ca.font = _font(bold=True, color=GREEN_FG)
            cb.fill = _fill(GREEN_BG); cb.font = _font(color=GREEN_FG, size=9)
            rn.row_dimensions[ri].height = 30
        elif label in ("T.F. O'Brien", "Confires Fire"):
            ca.fill = _fill(AMBER_BG); ca.font = _font(bold=True, color=AMBER_FG)
            cb.fill = _fill(AMBER_BG); cb.font = _font(color=AMBER_FG, size=9)
            rn.row_dimensions[ri].height = 30
        elif label in ("Sansone AC", "Piper Fire Protection", "Stan's Heating"):
            ca.fill = _fill(RED_BG); ca.font = _font(bold=True, color=RED_FG)
            cb.fill = _fill(RED_BG); cb.font = _font(color=RED_FG, size=9)
            rn.row_dimensions[ri].height = 24
        else:
            ca.fill = _fill(GRAY_BG); ca.font = _font(color=GRAY_FG, size=9)
            cb.fill = _fill(GRAY_BG); cb.font = _font(color=GRAY_FG, size=9)
            rn.row_dimensions[ri].height = 18

        for c in (ca, cb):
            c.border    = _border()
            c.alignment = _align(h="left", v="center", wrap=True)

    wb.save(output_path)
    print(f"\n✓ Pipeline saved → {output_path}")
    print(f"  Sheet 1 'Pipeline':               {len(PIPELINE_ROWS)} contacts researched")
    print(f"  Sheet 2 'Disqualified — PE Backed': {len(DISQUALIFIED)} removed (already PE-owned)")
    print(f"  Sheet 3 'Research Notes':          email formats + verification checklist")
    print(f"\n  Priority 1 (send this week):  Russell Reid, Aiello, Bruni & Campisi, Mowrey Elevator")
    print(f"  Priority 2 (verify first):    T.F. O'Brien, Confires Fire Protection")
    print(f"\n  CONFIRMED email:  m.jezouit@aiellohomeservices.com  (Aiello)")
    print(f"  ESTIMATED emails: gweiner@russellreid.com, k.bruni@bruniandcampisi.com,")
    print(f"                    t.mowrey@mowreyelevator.com, k.obrien@tfobrien.com,")
    print(f"                    william.wrublevski@confires.com")


def main():
    parser = argparse.ArgumentParser(description="Populate M&A pipeline from researched contacts")
    parser.add_argument("--output", default="MA_Pipeline_Populated.xlsx", help="Output Excel file")
    args = parser.parse_args()
    build_workbook(args.output)


if __name__ == "__main__":
    main()
