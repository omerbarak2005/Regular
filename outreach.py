#!/usr/bin/env python3
"""
QSR Drive-Thru Coffee — M&A Outreach Engine
Pacific Northwest | 34 Units | $70M+ Rev | $13M+ EBITDA

Usage:
    python outreach.py                          # generate CSV + preview
    python outreach.py --hunter-key <KEY>       # also verify emails via Hunter.io
    python outreach.py --output my_list.csv     # custom output filename
    python outreach.py --priority 1             # only priority-1 firms
    python outreach.py --type pe                # filter by firm type: pe|strategic|family_office
"""

import csv
import sys
import json
import argparse
import urllib.request
import urllib.parse
from dataclasses import dataclass, field
from typing import Optional, List
from datetime import date

# ═══════════════════════════════════════════════════════════════════════════════
#  DEAL PROFILE  —  edit these fields to match the deal
# ═══════════════════════════════════════════════════════════════════════════════
DEAL = {
    "codename":     "Project Alpine",
    "concept":      "Specialty Beverage / Drive-Thru Coffee",
    "region":       "Pacific Northwest",
    "units":        34,
    "revenue":      "$70M+",
    "ebitda":       "$13M+",
    "margin":       "~18.5%",
    "ev_range":     "$78M–$130M",
    "owners":       3,
    "highlights": [
        "34 drive-thru units across Pacific Northwest — proven, scalable model",
        "$70M+ revenue, $13M+ adjusted EBITDA with ~18.5% margin (best-in-class for QSR beverage)",
        "Sticky, habitual consumer behavior — specialty coffee drives 5–7x weekly visit frequency",
        "Drive-thru-only format: lean labor model, minimal capex vs. sit-down concepts",
        "Three motivated co-owners aligned on timeline — clean ownership, no corporate complexity",
        "Significant white space for unit growth within and beyond current markets",
    ],
}

ADVISOR = {
    "name":   "Jim Paulson",
    "title":  "Managing Director",
    "firm":   "[Advisory Firm]",
    "email":  "jim@[advisorfirm].com",
    "phone":  "[phone]",
    "intern": "Omer Barak",
}

# ═══════════════════════════════════════════════════════════════════════════════
#  EXCLUSION LISTS  —  never email these
# ═══════════════════════════════════════════════════════════════════════════════
REJECTED = {
    "Roark Capital",
    "GreyLion",
    "Keystone Capital",
    "Garnett Station Partners",
}

ALREADY_ACTIVE = {
    "Gauge Capital",
    "Sun Holdings",
    "Tucker's Farm",
    "Trive Capital",
    "Triton Pacific",
    "Peak Rock Capital",
    "L Catterton",
    "Brentwood Associates",
    "Riverside",
    "CapitalSpring",
    "Center Bridge",
    "Blackstone Tac Opps",
    "Orangewood Partners",
    "Trivest",
    "FEP",
    "Eyas Capital",
}

# ═══════════════════════════════════════════════════════════════════════════════
#  DATA MODEL
# ═══════════════════════════════════════════════════════════════════════════════
@dataclass
class Contact:
    firm:             str
    name:             str
    title:            str
    email:            str
    phone:            str  = ""
    firm_type:        str  = "pe"          # pe | strategic | family_office | wealth
    priority:         int  = 2             # 1=highest  2=medium  3=lower
    email_confidence: str  = "probable"    # confirmed | probable | format_only
    notes:            str  = ""
    nda_status:       str  = "Not Sent"
    status:           str  = "Not Yet Contacted"

# ═══════════════════════════════════════════════════════════════════════════════
#  CONTACT DATABASE
#  priority 1 = direct QSR/beverage/PNW fit
#  priority 2 = strong LMM PE or strategic fit
#  priority 3 = family office, wealth, or longer-shot generalist
# ═══════════════════════════════════════════════════════════════════════════════
CONTACTS: List[Contact] = [

    # ── PRIORITY 1 : QSR / Beverage / PNW specialist ──────────────────────────

    Contact("Savory Fund", "Andrew K. Smith", "Managing Partner & Co-Founder",
            "asmith@mercatopartners.com", "", "pe", 1, "confirmed",
            "Pure-play restaurant PE ($750M AUM); direct drive-thru coffee comp."),

    Contact("CapitalSpring", "Richard Fitzgerald", "Co-Founder & Managing Partner",
            "rfitzgerald@capitalspring.com", "", "pe", 1, "confirmed",
            "Restaurant-focused debt + equity — built for QSR deals exactly this size."),

    Contact("CapitalSpring", "Erik Herrmann", "Partner & Head of Investments",
            "eherrmann@capitalspring.com", "", "pe", 1, "confirmed",
            "Restaurant-focused debt + equity — built for QSR deals exactly this size."),

    Contact("10 Point Capital", "Scott Pressly", "Founder & Managing Partner",
            "spressly@10pointcapital.com", "", "pe", 1, "probable",
            "Pure-play franchise/QSR PE; invests in drive-thru concepts nationally."),

    Contact("Sentinel Capital Partners", "John McCormack", "Co-Founder & Senior Partner",
            "jmccormack@sentinelpartners.com", "(212) 688-3100", "pe", 1, "probable",
            "Owns Checkers/Rally's drive-thru brand — direct operational comp."),

    Contact("KarpReilly", "Chris Reilly", "Co-Founder & Partner",
            "creilly@karpreilly.com", "(203) 504-9911", "pe", 1, "probable",
            "Consumer/restaurant growth equity, $100M–$500M EV sweet spot."),

    Contact("KarpReilly", "Allan Karp", "Co-Founder & Partner",
            "akarp@karpreilly.com", "", "pe", 1, "probable",
            "Consumer/restaurant growth equity, $100M–$500M EV sweet spot."),

    Contact("Endeavour Capital", "John von Schlegell", "Founder & Managing Director",
            "jvonschlegell@endeavourcapital.com", "(503) 223-2721", "pe", 1, "probable",
            "Portland-based PE — native PNW firm, highest geographic fit on the list."),

    Contact("Cascadia Capital", "Michael Butler", "Co-Founder & Chairman/CEO",
            "mbutler@cascadiacapital.com", "(206) 436-2500", "pe", 1, "probable",
            "Seattle M&A advisory with deep PNW F&B relationships; potential co-advisor or buyer intro."),

    Contact("Permanent Equity", "Brent Beshore", "Founder & CEO",
            "brent@permanentequity.com", "(573) 445-0678", "pe", 1, "confirmed",
            "Long-hold PE; ideal if owners care about legacy over max price."),

    Contact("FAT Brands", "Andy Wiederhorn", "Executive Chairman",
            "awiederhorn@fatbrands.com", "", "strategic", 1, "probable",
            "Aggressive brand acquirer (25+ brands); public company needing growth story."),

    Contact("Tavistock Restaurant Collection", "Joe Lewis", "Owner",
            "joe.lewis@tavistock.com", "", "strategic", 1, "probable",
            "Operator family office; owns multiple restaurant concepts."),

    Contact("Savory Fund (COO)", "Josh Boshard", "COO",
            "jboshard@mercatopartners.com", "", "pe", 1, "probable",
            "Day-to-day deal exec at Savory — good secondary contact."),

    # ── PRIORITY 2 : Strong LMM PE / strategic fit ────────────────────────────

    Contact("Frontenac Company", "Walter Florence", "Managing Partner",
            "wflorence@frontenac.com", "(312) 629-3152", "pe", 2, "confirmed",
            "Consumer sector specialist, Chicago LMM PE."),

    Contact("Huron Capital", "Brian Demkowicz", "Co-Founder & Chairman",
            "bdemkowicz@huroncapital.com", "(313) 962-5801", "pe", 2, "confirmed",
            "Operationally-focused LMM PE."),

    Contact("Alpine Investors", "Graham Weaver", "Founder & CEO",
            "gweaver@alpineinvestors.com", "", "pe", 2, "probable",
            "People-first PE firm; consumer services focus. Very active dealmaker."),

    Contact("Gridiron Capital", "Tom Burger", "Co-Founder & Managing Partner",
            "tburger@gridironcapital.com", "", "pe", 2, "confirmed",
            "35+ years middle-market, consumer brand focus."),

    Contact("Rotunda Capital Partners", "John Fruehwirth", "Founder & Managing Partner",
            "jf@rotundacapital.com", "(240) 482-0610", "pe", 2, "confirmed",
            "LMM PE, founded 2009; lower-mid market sweet spot matches deal size."),

    Contact("Southfield Capital", "Andy Levison", "Founder & Managing Partner",
            "alevison@southfieldcapital.com", "(203) 813-4100", "pe", 2, "confirmed",
            "Consumer/business services LMM."),

    Contact("Prairie Capital", "Nate Good", "Managing Partner",
            "ngood@prairie-capital.com", "(312) 265-2950", "pe", 2, "confirmed",
            "Chicago LMM PE, consumer/services thesis."),

    Contact("Prairie Capital", "Steve King", "Founding Partner",
            "sking@prairie-capital.com", "", "pe", 2, "probable",
            "Chicago LMM PE, consumer/services thesis."),

    Contact("Wellspring Capital Management", "Greg Feldman", "Executive Chairman & Co-Founder",
            "gfeldman@wellspringcapital.com", "", "pe", 2, "probable",
            "Mid-market PE, consumer brands focus."),

    Contact("Comvest Partners", "Robert O'Sullivan", "CEO",
            "robert@comvest.com", "", "pe", 2, "confirmed",
            "Lower middle market buyout, consumer exposure."),

    Contact("Comvest Partners", "Michael Falk", "Chairman",
            "michaelf@comvest.com", "", "pe", 2, "confirmed",
            "Lower middle market buyout, consumer exposure."),

    Contact("Bertram Capital", "Jeff Drazan", "Founder & Managing Partner",
            "jdrazan@bertramcapital.com", "", "pe", 2, "probable",
            "West Coast-based LMM PE, consumer/services."),

    Contact("Blackford Capital", "Martin Stein", "Founder & Managing Director",
            "mstein@blackfordcapital.com", "", "pe", 2, "probable",
            "Active LMM acquirer, consumer focus."),

    Contact("Voyager Capital", "Erik Benson", "General Partner",
            "benson@voyagercapital.com", "(206) 438-1804", "pe", 2, "confirmed",
            "Seattle-based — PNW geographic proximity."),

    Contact("Columbia Pacific Advisors", "Alex Cheung", "Managing Director",
            "acheung@columbiapacificadvisors.com", "", "pe", 2, "probable",
            "PNW-based family office/PE — strong geographic fit."),

    Contact("H.I.G. Capital", "Sami Mnaymneh", "Co-Founder & Managing Partner",
            "smnaymneh@higcapital.com", "", "pe", 2, "probable",
            "Large platform with active LMM consumer division."),

    Contact("Genstar Capital", "Jean-Pierre Conte", "Co-Founder & Managing Partner",
            "jpconte@genstarcapital.com", "", "pe", 2, "probable",
            "West Coast PE, consumer and industrial sectors."),

    Contact("Branford Castle Partners", "Robert Hurst", "Founder & Managing Partner",
            "rhurst@branfordcastle.com", "", "pe", 2, "probable",
            "LMM PE, consumer/industrial."),

    Contact("Wind Point Partners", "Bob Cummings", "Managing Partner",
            "bcummings@windpointpartners.com", "", "pe", 2, "probable",
            "Consumer/services LMM PE."),

    Contact("Blue Point Capital Partners", "Jon Pressnell", "Partner",
            "jpressnell@bluepointcapital.com", "", "pe", 2, "probable",
            "LMM PE, consumer/services — NOTE: listed as 'Blue Sea Capital' in sheet, firm is Blue Point."),

    Contact("Trinity Hunt Partners", "Blake Apel", "Managing Partner",
            "bapel@trinityhunt.com", "(214) 777-6600", "pe", 2, "probable",
            "Dallas LMM PE, consumer services. NOTE: listed as 'Jerry Kaufman' in sheet — correct is Blake Apel."),

    Contact("Edgewater Capital Partners", "Bob Girton", "Managing Partner",
            "bgirton@edgewatercapital.com", "(216) 292-3838", "pe", 2, "probable",
            "Ohio LMM PE. NOTE: listed as 'Todd Ofenloch' in sheet — correct MP is Bob Girton."),

    Contact("Pfingsten Partners", "Frank Beans", "CEO",
            "fbeans@pfingsten.com", "", "pe", 2, "probable",
            "Chicago LMM PE. NOTE: 'John Feeney' not confirmed at firm — correct CEO is Frank Beans."),

    Contact("Turnspire Capital Partners", "Christopher Zimmerman", "Founder & Managing Partner",
            "czimmerman@turnspire.com", "", "pe", 2, "probable",
            "LMM consumer/industrial PE."),

    Contact("McNally Capital", "Ward McNally", "Founder & Co-CEO",
            "wmcnally@mcnallycapital.com", "(312) 357-3710", "pe", 2, "probable",
            "Chicago LMM PE. NOTE: listed as 'William McNally' — correct first name is Ward."),

    Contact("Insignia Capital Group", "Chris Satterfield", "Managing Partner",
            "csatterfield@insigniacap.com", "(925) 399-8900", "pe", 2, "probable",
            "Walnut Creek CA LMM PE, consumer/services focus."),

    Contact("Arlon Group", "David Nussbaum", "Managing Partner",
            "dnussbaum@arlongroup.com", "", "pe", 2, "probable",
            "Food & agriculture PE — beverage/consumer angle is a fit. NOTE: 'Seth Boro' is now at Thoma Bravo, not Arlon."),

    Contact("Svoboda Capital Partners", "John Svoboda", "Founder & Managing Partner",
            "jsvoboda@svobodacapital.com", "", "pe", 2, "probable",
            "Chicago consumer/industrial LMM."),

    Contact("Spell Capital Partners", "J. Anthony Speller", "Founder & Managing Partner",
            "aspeller@spellcapital.com", "", "pe", 2, "probable",
            "Minneapolis LMM PE."),

    Contact("Bruin Capital", "George Pyne", "Founder & CEO",
            "gpyne@bruincapital.com", "", "pe", 2, "probable",
            "Consumer/brand-focused PE."),

    Contact("Dine Brands", "John Peyton", "CEO",
            "jpeyton@dinebrands.com", "", "strategic", 2, "probable",
            "Franchise development arm active in QSR — potential strategic acquirer."),

    Contact("Masked Rider Capital", "Jeff Horn", "CEO & Co-Founder",
            "jhorn@maskedridercapital.com", "", "pe", 2, "probable",
            "Restaurant-focused PE."),

    Contact("Dom Capital Group", "Jay Owen", "Co-Founder & Managing Partner",
            "jowen@domcapital.com", "", "pe", 2, "probable",
            "LMM PE."),

    Contact("New MainStream Capital", "Brian Kelley", "Managing Partner",
            "bkelley@newmainstreamcapital.com", "", "pe", 2, "probable",
            "Consumer LMM PE."),

    Contact("Apax Partners (Consumer)", "Andrew Sillitoe", "Co-CEO",
            "andrew.sillitoe@apax.com", "", "pe", 2, "probable",
            "Large global PE; consumer division relevant at upper end of EV range."),

    Contact("Broadtree Partners", "Michael Faber", "Managing Partner",
            "mfaber@broadtreepartners.com", "", "pe", 2, "probable",
            "Rochester LMM PE."),

    Contact("Staple Street Capital", "Stephen Murray", "Co-Founder & Managing Partner",
            "smurray@staplestreetcapital.com", "", "pe", 2, "probable",
            "LMM PE."),

    Contact("Alpine Investors (Consumer)", "Matt Wilensky", "Partner",
            "mwilensky@alpineinvestors.com", "", "pe", 2, "probable",
            "Alpine's consumer/services deal lead."),

    # ── PRIORITY 3 : Family office, wealth, longer-shot ──────────────────────

    Contact("Pritzker Private Capital", "Anthony Pritzker", "Principal",
            "apritzker@pritzker.com", "", "family_office", 3, "probable",
            "Family office with consumer brand experience."),

    Contact("Tao Capital Partners", "Nick Pritzker", "Principal",
            "npritzker@taocapital.com", "", "family_office", 3, "probable",
            "Consumer-focused family office."),

    Contact("Chenmark Capital", "James Sharpe", "Co-Founder & Managing Partner",
            "james@chenmarkcapital.com", "", "pe", 3, "probable",
            "Long-hold PE — good fit if owners want permanent capital."),

    Contact("Boyne Capital", "Jeffrey Sands", "Managing Partner",
            "jsands@boynecapital.com", "", "pe", 3, "probable",
            "Florida LMM PE."),

    Contact("Platte River Equity", "Michael Fayhee", "Managing Partner",
            "mfayhee@platteriverequity.com", "", "pe", 3, "probable",
            "Denver LMM PE — Western US geography helps."),

    Contact("New Heritage Capital", "Bret Kuchenbecker", "Managing Partner",
            "bkuchenbecker@newheritagecapital.com", "(617) 439-0688", "pe", 3, "probable",
            "Boston LMM PE. NOTE: 'Rich Maclean' not confirmed at firm."),

    Contact("New Heritage Capital", "Greg Katz", "Managing Partner",
            "gkatz@newheritagecapital.com", "", "pe", 3, "probable",
            "Boston LMM PE."),

    Contact("Balance Point Capital", "Nathan Elliott", "Partner",
            "nelliott@balancepointcapital.com", "(203) 652-8250", "pe", 3, "probable",
            "Westport CT LMM PE. NOTE: 'Michael Ouimet' not confirmed at firm."),

    Contact("Coldstream Wealth Management", "Krista Charlton", "CFP",
            "kcharlton@coldstreamwealth.com", "", "wealth", 3, "probable",
            "Bend OR-based — PNW HNW client base, may know buyers."),

    Contact("Ferguson Wellman Capital", "Ralph Cole", "CIO",
            "rcole@fergusonwellman.com", "", "wealth", 3, "probable",
            "Portland-based wealth manager — investor network."),

    Contact("Whittier Trust", "David Dahl", "President & CEO",
            "ddahl@whittier.com", "", "wealth", 3, "probable",
            "Multi-family office, consumer asset exposure."),

    Contact("Wetherby Asset Management", "Kim Kardell", "CEO",
            "kkardell@wetherby.com", "", "wealth", 3, "probable",
            "Wealth manager with direct investment appetite."),

    Contact("Endeavour Capital (Partner)", "John W. Dixon", "Partner",
            "jdixon@endeavourcapital.com", "", "pe", 3, "probable",
            "Portland PE secondary contact."),

    Contact("JMB Capital Partners", "Jonathan Brooks", "Partner",
            "jbrooks@jmbcapital.com", "", "pe", 3, "probable",
            "Chicago PE."),

    Contact("Fernandez Holdings", "Kirk Fernandez", "Founder & CEO",
            "kfernandez@fernandezholdings.com", "", "family_office", 3, "probable",
            "Family holding company."),

    Contact("GreenOak / BentallGreenOak", "John Carrafiell", "Co-Founder & Managing Partner",
            "john.carrafiell@bentallgreenoak.com", "", "pe", 3, "probable",
            "Real estate-heavy; consumer real estate angle only."),
]


# ═══════════════════════════════════════════════════════════════════════════════
#  EMAIL TEMPLATE ENGINE
# ═══════════════════════════════════════════════════════════════════════════════

_SUBJECT_PE = (
    "[Project Alpine] Drive-Thru Coffee | PNW | 34 Units | "
    "$70M Rev / $13M+ EBITDA | Seeking Capital Partner"
)
_SUBJECT_STRATEGIC = (
    "[Project Alpine] Drive-Thru Coffee Acquisition | PNW | 34 Units | "
    "$70M Rev / $13M+ EBITDA"
)
_SUBJECT_FO = (
    "[Project Alpine] Direct Investment | PNW Drive-Thru Coffee | "
    "34 Units | $13M+ EBITDA | ~18.5% Margin"
)

def _first(name: str) -> str:
    return name.split()[0]

def build_email(c: Contact) -> dict:
    """Return subject + body dict for a contact."""
    fn = _first(c.name)
    highlights_str = "\n".join(f"  • {h}" for h in DEAL["highlights"])

    if c.firm_type == "pe":
        subject = _SUBJECT_PE
        body = f"""Hi {fn},

I'm reaching out on behalf of {ADVISOR["name"]} at {ADVISOR["firm"]}, who is advising the three co-owners of a 34-unit specialty beverage / drive-thru coffee brand in the Pacific Northwest on a potential capital event.

Key highlights:
{highlights_str}

We estimate enterprise value in the {DEAL["ev_range"]} range at market-clearing multiples for this segment.

Given {c.firm}'s focus on consumer and lower middle-market investments, I believe this is a strong fit. We're happy to share a brief teaser under NDA if there's interest — no commitment required.

Would you have 15 minutes this week or next?

Best regards,
{ADVISOR["name"]}
{ADVISOR["title"]} | {ADVISOR["firm"]}
{ADVISOR["email"]} | {ADVISOR["phone"]}"""

    elif c.firm_type == "strategic":
        subject = _SUBJECT_STRATEGIC
        body = f"""Hi {fn},

I'm reaching out on behalf of three co-owners of a 34-unit drive-thru coffee concept in the Pacific Northwest exploring a strategic transaction.

The business is generating $70M+ in revenue and $13M+ in adjusted EBITDA on an ~18.5% margin — best-in-class for the specialty beverage drive-thru segment. The brand has strong unit-level economics and meaningful white space for unit growth.

Key highlights:
{highlights_str}

Given {c.firm}'s scale and presence in the restaurant/QSR space, we thought this could be an interesting bolt-on or platform acquisition. Happy to share more detail under NDA.

Would you be open to a brief call?

Best regards,
{ADVISOR["name"]}
{ADVISOR["title"]} | {ADVISOR["firm"]}
{ADVISOR["email"]} | {ADVISOR["phone"]}"""

    else:  # family_office / wealth
        subject = _SUBJECT_FO
        body = f"""Hi {fn},

I'm reaching out on behalf of {ADVISOR["name"]} at {ADVISOR["firm"]}, representing three co-owners of a 34-unit drive-thru coffee brand in the Pacific Northwest.

The business is a rare combination of scale and margin quality:
  • $70M+ revenue, $13M+ adjusted EBITDA (~18.5% margin)
  • Habitual, defensive consumer product — specialty coffee
  • Drive-thru-only model: lean operations, predictable cash flows
  • Established brand, 34 locations, Pacific Northwest presence

Enterprise value estimated at {DEAL["ev_range"]}. Given your office's focus on direct investments in high-quality consumer businesses, I wanted to bring this to your attention directly.

Happy to share a teaser under NDA.

Best regards,
{ADVISOR["name"]}
{ADVISOR["title"]} | {ADVISOR["firm"]}
{ADVISOR["email"]} | {ADVISOR["phone"]}"""

    return {"subject": subject, "body": body}


# ═══════════════════════════════════════════════════════════════════════════════
#  HUNTER.IO EMAIL VERIFICATION  (optional)
# ═══════════════════════════════════════════════════════════════════════════════

def verify_email_hunter(email: str, api_key: str) -> str:
    """Return 'deliverable', 'risky', 'undeliverable', or 'unknown'."""
    try:
        domain = email.split("@")[1]
        params  = urllib.parse.urlencode({"email": email, "api_key": api_key})
        url     = f"https://api.hunter.io/v2/email-verifier?{params}"
        with urllib.request.urlopen(url, timeout=8) as resp:
            data = json.loads(resp.read())
        return data.get("data", {}).get("result", "unknown")
    except Exception:
        return "unknown"

def get_email_format_hunter(domain: str, api_key: str) -> str:
    """Return the most common email format for a domain."""
    try:
        params = urllib.parse.urlencode({"domain": domain, "api_key": api_key})
        url    = f"https://api.hunter.io/v2/domain-search?{params}"
        with urllib.request.urlopen(url, timeout=8) as resp:
            data = json.loads(resp.read())
        pattern = data.get("data", {}).get("pattern", "unknown")
        return pattern
    except Exception:
        return "unknown"


# ═══════════════════════════════════════════════════════════════════════════════
#  FILTER & DEDUPLICATE
# ═══════════════════════════════════════════════════════════════════════════════

def get_targets(priority_filter: Optional[int] = None,
                type_filter: Optional[str] = None) -> List[Contact]:
    out = []
    seen_emails = set()
    for c in CONTACTS:
        if c.firm in REJECTED or c.firm in ALREADY_ACTIVE:
            continue
        if priority_filter and c.priority != priority_filter:
            continue
        if type_filter and c.firm_type != type_filter:
            continue
        if c.email in seen_emails:
            continue
        seen_emails.add(c.email)
        out.append(c)
    return sorted(out, key=lambda x: (x.priority, x.firm))


# ═══════════════════════════════════════════════════════════════════════════════
#  CSV EXPORT
# ═══════════════════════════════════════════════════════════════════════════════

SHEET_COLUMNS = [
    "Priority", "Date", "Company", "Contact", "Title", "Email",
    "Phone", "Firm Type", "Email Confidence", "Activity Type",
    "NDA Status", "Status After", "Logged By", "Next Steps",
    "Email Subject", "Email Body", "Notes",
]

def export_csv(targets: List[Contact], path: str, hunter_key: Optional[str] = None):
    today = date.today().strftime("%-m/%-d/%y")
    rows = []
    for c in targets:
        email_obj = build_email(c)
        verification = ""
        if hunter_key and c.email_confidence != "format_only":
            result = verify_email_hunter(c.email, hunter_key)
            verification = f"Hunter: {result}"

        rows.append({
            "Priority":         c.priority,
            "Date":             today,
            "Company":          c.firm,
            "Contact":          c.name,
            "Title":            c.title,
            "Email":            c.email,
            "Phone":            c.phone,
            "Firm Type":        c.firm_type,
            "Email Confidence": f"{c.email_confidence}" + (f" | {verification}" if verification else ""),
            "Activity Type":    "Initial Outreach",
            "NDA Status":       c.nda_status,
            "Status After":     "Outreach – No Response Yet",
            "Logged By":        "Omer",
            "Next Steps":       "Follow up in 1 week",
            "Email Subject":    email_obj["subject"],
            "Email Body":       email_obj["body"],
            "Notes":            c.notes,
        })

    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=SHEET_COLUMNS)
        writer.writeheader()
        writer.writerows(rows)

    print(f"\n✓ Exported {len(rows)} contacts → {path}")


# ═══════════════════════════════════════════════════════════════════════════════
#  CONSOLE SUMMARY
# ═══════════════════════════════════════════════════════════════════════════════

def print_summary(targets: List[Contact]):
    p1 = [c for c in targets if c.priority == 1]
    p2 = [c for c in targets if c.priority == 2]
    p3 = [c for c in targets if c.priority == 3]
    confirmed = [c for c in targets if c.email_confidence == "confirmed"]

    print("\n" + "═" * 70)
    print(f"  PROJECT ALPINE — OUTREACH PIPELINE  |  {date.today()}")
    print("═" * 70)
    print(f"  Deal:      {DEAL['units']}-unit drive-thru coffee | {DEAL['region']}")
    print(f"  Revenue:   {DEAL['revenue']}   EBITDA: {DEAL['ebitda']}   Margin: {DEAL['margin']}")
    print(f"  EV Range:  {DEAL['ev_range']}")
    print("─" * 70)
    print(f"  Total targets:       {len(targets)}")
    print(f"  Priority 1 (HOT):    {len(p1)}")
    print(f"  Priority 2 (WARM):   {len(p2)}")
    print(f"  Priority 3 (COLD):   {len(p3)}")
    print(f"  Confirmed emails:    {len(confirmed)}")
    print(f"  Excluded (rejected): {len(REJECTED)} firms")
    print(f"  Excluded (active):   {len(ALREADY_ACTIVE)} firms")
    print("─" * 70)

    print("\n  PRIORITY 1 — SEND FIRST\n")
    for c in p1:
        conf_tag = "✓" if c.email_confidence == "confirmed" else "~"
        print(f"  {conf_tag}  {c.firm:<40} {c.name}")
        print(f"     {c.email:<45} {c.phone}")

    print("\n  PRIORITY 2 — SEND SECOND WAVE\n")
    for c in p2:
        conf_tag = "✓" if c.email_confidence == "confirmed" else "~"
        print(f"  {conf_tag}  {c.firm:<40} {c.name}")
        print(f"     {c.email:<45} {c.phone}")

    print("\n  SHEET CORRECTIONS NEEDED (wrong names in original tracker)\n")
    corrections = [
        ("Rotunda Capital", "Charlie Homet", "John Fruehwirth", "jf@rotundacapital.com"),
        ("Southfield Capital", "Tom Lauer", "Andy Levison", "alevison@southfieldcapital.com"),
        ("Endeavour Capital", "John Sorte", "John von Schlegell", "jvonschlegell@endeavourcapital.com"),
        ("Arlon Group", "Seth Boro", "Seth Boro LEFT (now Thoma Bravo)", "contact firm directly"),
        ("Prairie Capital", "Kevin McTavish", "Nate Good (MP)", "ngood@prairie-capital.com"),
        ("Blue Sea Capital", "Jonathan Pressnell", "Blue POINT Capital Partners", "jpressnell@bluepointcapital.com"),
        ("Trinity Hunt", "Jerry Kaufman", "Blake Apel (MP)", "bapel@trinityhunt.com"),
        ("Edgewater Capital", "Todd Ofenloch", "Bob Girton (MP)", "bgirton@edgewatercapital.com"),
        ("Pfingsten Partners", "John Feeney", "Frank Beans (CEO)", "fbeans@pfingsten.com"),
        ("McNally Capital", "William McNally", "Ward McNally (correct first name)", "wmcnally@mcnallycapital.com"),
        ("New Heritage Capital", "Rich Maclean", "Bret Kuchenbecker / Greg Katz", "bkuchenbecker@newheritagecapital.com"),
    ]
    for firm, wrong, right, email in corrections:
        print(f"  ✗  {firm:<30} '{wrong}'  →  {right}")
        print(f"     Correct email: {email}")

    print("\n" + "═" * 70)


# ═══════════════════════════════════════════════════════════════════════════════
#  EMAIL PREVIEW
# ═══════════════════════════════════════════════════════════════════════════════

def print_email_previews(targets: List[Contact], n: int = 3):
    print(f"\n  SAMPLE EMAILS (first {n} priority-1 contacts)\n")
    p1 = [c for c in targets if c.priority == 1][:n]
    for c in p1:
        em = build_email(c)
        print("─" * 70)
        print(f"  TO:       {c.name} <{c.email}>")
        print(f"  SUBJECT:  {em['subject']}")
        print(f"\n{em['body']}")
    print("─" * 70)


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(description="QSR M&A Outreach Engine")
    parser.add_argument("--hunter-key", metavar="KEY",  help="Hunter.io API key for email verification")
    parser.add_argument("--output",     metavar="FILE", default="qsr_outreach.csv", help="Output CSV filename")
    parser.add_argument("--priority",   type=int, choices=[1, 2, 3], help="Filter by priority level")
    parser.add_argument("--type",       choices=["pe", "strategic", "family_office", "wealth"], help="Filter by firm type")
    parser.add_argument("--preview",    action="store_true", help="Print sample email bodies to console")
    args = parser.parse_args()

    targets = get_targets(priority_filter=args.priority, type_filter=args.type)

    print_summary(targets)

    if args.preview:
        print_email_previews(targets)

    export_csv(targets, args.output, hunter_key=args.hunter_key)

    print(f"\n  Next steps:")
    print(f"  1. Open {args.output} in Excel / Sheets — paste into your tracker")
    print(f"  2. Update ADVISOR dict at top of script with Jim's real email + phone")
    print(f"  3. Run --hunter-key <KEY> to verify emails (free tier: 25/mo at hunter.io)")
    print(f"  4. Send priority-1 wave first, follow up in 5 business days\n")


if __name__ == "__main__":
    main()
