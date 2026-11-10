"""Build Bella256 Insurance & Credentialing Tracker (xlsx, Google Sheets compatible)."""
import sys
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

OUT = sys.argv[1]

# ---------- style ----------
FONT = "Arial"
ROSE = "D3A4A5"; ROSE_DK = "BC8A8B"; TEAL = "8FC5C5"; BLUSH = "F7F0EF"; INK = "3F3F3F"
GRAY_HDR = "8C8C8C"; GRAY_FILL = "F2F2F2"
RED_FILL = "F8D7DA"; RED_TXT = "9C0006"; ORANGE_FILL = "FCE5CD"; YEL_FILL = "FFF2CC"; GRN_FILL = "D9EAD3"

f_base = Font(name=FONT, size=10, color=INK)
f_hdr = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_title = Font(name=FONT, size=16, bold=True, color=INK)
f_h2 = Font(name=FONT, size=11, bold=True, color="FFFFFF")
f_bold = Font(name=FONT, size=10, bold=True, color=INK)
f_example = Font(name=FONT, size=10, italic=True, color="808080")
fill = lambda c: PatternFill("solid", start_color=c, end_color=c)
thin = Side(style="thin", color="DDDDDD")
border = Border(bottom=thin)
DATE_FMT = "mm/dd/yyyy"
USD = '$#,##0.00'

wb = Workbook()

def red(ws, rng, formula):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=fill(RED_FILL), font=Font(name=FONT, color=RED_TXT, bold=True)))
def orange(ws, rng, formula):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=fill(ORANGE_FILL)))
def yellow(ws, rng, formula):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=fill(YEL_FILL)))
def green(ws, rng, formula):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=fill(GRN_FILL)))


def build_table(ws, cols, nrows, rows=None, example=None, freeze="B2"):
    """cols: list of dicts: h, w, kind('in'|'auto'), f (formula template with {r}), fmt, dv (formula1), dv_strict, note.
    rows: list of lists of prefilled input values (None for formula cols are ignored)."""
    ws.sheet_properties.tabColor = ROSE
    last = get_column_letter(len(cols))
    for i, c in enumerate(cols, 1):
        L = get_column_letter(i)
        cell = ws.cell(row=1, column=i, value=c["h"])
        cell.font = f_hdr
        cell.fill = fill(ROSE_DK if c.get("kind", "in") == "in" else GRAY_HDR)
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
        ws.column_dimensions[L].width = c.get("w", 14)
        if c.get("note"):
            cm = Comment(c["note"], "Tracker")
            cm.width, cm.height = 260, 110
            cell.comment = cm
        if c.get("dv"):
            dv = DataValidation(type="list", formula1=c["dv"], allow_blank=True)
            if c.get("dv_strict", True):
                dv.error = "Pick a value from the list (edit options on the Lists tab)."
                dv.errorTitle = "Use the dropdown"
            else:
                dv.showErrorMessage = False
            dv.add(f"{L}2:{L}{nrows + 1}")
            ws.add_data_validation(dv)
    ws.row_dimensions[1].height = 42
    for r in range(2, nrows + 2):
        for i, c in enumerate(cols, 1):
            cell = ws.cell(row=r, column=i)
            cell.font = f_base
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=c.get("wrap", False))
            if c.get("fmt"):
                cell.number_format = c["fmt"]
            if c.get("kind") == "auto":
                cell.value = c["f"].format(r=r)
                cell.fill = fill(GRAY_FILL)
    data = []
    if example:
        data.append(("ex", example))
    for row in rows or []:
        data.append(("row", row))
    for idx, (kind, row) in enumerate(data):
        r = idx + 2
        for i, v in enumerate(row, 1):
            if v is None or cols[i - 1].get("kind") == "auto":
                continue
            cell = ws.cell(row=r, column=i, value=v)
            if kind == "ex":
                cell.font = f_example
        if kind == "ex":
            for i, c in enumerate(cols, 1):
                if c.get("kind") == "auto":
                    ws.cell(row=r, column=i).font = f_example
    ws.freeze_panes = freeze
    ws.auto_filter.ref = f"A1:{last}{nrows + 1}"
    return nrows + 1


# =====================================================================
# Lists (dropdown sources)
# =====================================================================
lists = {
    "A": ("Application type", ["Initial enrollment", "Revalidation", "Re-credentialing", "Change of information", "Add location", "Contract renegotiation"]),
    "B": ("Application status", ["Not started", "Gathering documents", "Submitted", "In review", "Info requested", "Approved – awaiting contract", "Active / effective", "Denied", "Network closed", "Withdrawn"]),
    "C": ("In pipeline? (Y/N)", ["N", "Y", "Y", "Y", "Y", "Y", "N", "N", "N", "N"]),
    "D": ("Claim status", ["Draft – not sent", "Submitted", "Accepted by clearinghouse", "Pending", "Partially paid", "Paid in full", "Denied", "Appeal filed", "Corrected / rebilled", "Closed – written off"]),
    "E": ("Open A/R? (Y/N)", ["N", "Y", "Y", "Y", "Y", "N", "Y", "Y", "Y", "N"]),
    "F": ("Payer type", ["Medicare", "Medicaid", "Medicaid CMO", "Commercial", "Medicare Advantage", "Other"]),
    "G": ("Y/N", ["Y", "N"]),
    "H": ("Auth status", ["Not required", "Requested", "Pending", "Approved", "Denied", "Expired"]),
    "I": ("Roadmap status", ["Not started", "In progress", "Waiting on payer", "Done", "N/A"]),
    "J": ("Related to", ["Application", "Claim", "Auth / eligibility", "Document", "Other"]),
    "K": ("Document category", ["Business", "Identifiers", "Licenses & certs", "Accreditation", "Insurance & bonds", "Banking", "Compliance", "Enrollment IDs"]),
    "L": ("Renewal cycle", ["None", "Monthly", "120 days", "Annual", "2 years", "3 years", "5 years"]),
    "M": ("DME network status", ["Unknown", "Open", "Closed", "Managed by vendor"]),
    "N": ("Payment method", ["Card", "PayPal", "Stripe", "Cash", "Check", "HSA/FSA card"]),
    "O": ("Yes/No/Unknown", ["Yes", "No", "Unknown", "Pending"]),
}
wsL = wb.active
wsL.title = "Lists"
for col, (hdr, vals) in lists.items():
    c = wsL[f"{col}1"]; c.value = hdr; c.font = f_hdr; c.fill = fill(GRAY_HDR)
    c.alignment = Alignment(wrap_text=True)
    wsL.column_dimensions[col].width = 26
    for i, v in enumerate(vals, 2):
        wsL[f"{col}{i}"] = v
        wsL[f"{col}{i}"].font = f_base
wsL["Q1"] = "Edit or add options in these columns. Keep columns C and E (Y/N) lined up with the status next to them — the Dashboard uses them to decide what counts as 'in pipeline' and 'open A/R'."
wsL["Q1"].font = f_bold; wsL["Q1"].alignment = Alignment(wrap_text=True, vertical="top")
wsL.column_dimensions["Q"].width = 50
wsL.row_dimensions[1].height = 80
wsL.freeze_panes = "A2"
LST = lambda col, n=30: f"=Lists!${col}$2:${col}${n}"

# =====================================================================
# Codes (reference)
# =====================================================================
wsC = wb.create_sheet("Codes")
hcpcs = [
    ("L8000", "Mastectomy bra, without integrated breast form", "Medicare: RT/LT not needed for bra; check LCD for quantity"),
    ("L8001", "Mastectomy bra with integrated breast form, unilateral", "Append RT or LT"),
    ("L8002", "Mastectomy bra with integrated breast form, bilateral", ""),
    ("L8010", "Mastectomy sleeve", ""),
    ("L8015", "External breast prosthesis garment with mastectomy form (post-mastectomy camisole)", ""),
    ("L8020", "Breast prosthesis, mastectomy form (non-silicone)", "Append RT or LT"),
    ("L8030", "Breast prosthesis, silicone or equal, without integral adhesive", "Append RT or LT"),
    ("L8031", "Breast prosthesis, silicone or equal, with integral adhesive", "Append RT or LT"),
    ("L8032", "Nipple prosthesis, prefabricated, reusable", "Append RT or LT"),
    ("L8035", "Custom breast prosthesis, molded to patient model", "Append RT or LT"),
    ("L8039", "Breast prosthesis, not otherwise specified", "Needs description/invoice"),
    ("A4280", "Adhesive skin support attachment for external breast prosthesis, each", ""),
    ("A9282", "Wig, any type, each", "Medicare excludes wigs. Some commercial plans cover as 'cranial prosthesis' with a prescription"),
]
icd = [
    ("Z90.10", "Acquired absence of unspecified breast and nipple"),
    ("Z90.11", "Acquired absence of right breast and nipple"),
    ("Z90.12", "Acquired absence of left breast and nipple"),
    ("Z90.13", "Acquired absence of bilateral breasts and nipples"),
    ("C50.911", "Malignant neoplasm of unspecified site of right female breast"),
    ("C50.912", "Malignant neoplasm of unspecified site of left female breast"),
    ("C50.919", "Malignant neoplasm of unspecified site of unspecified female breast"),
    ("Z85.3", "Personal history of malignant neoplasm of breast"),
    ("I97.2", "Postmastectomy lymphedema syndrome"),
    ("L65.9", "Nonscarring hair loss, unspecified"),
    ("L63.9", "Alopecia areata, unspecified"),
    ("T45.1X5A", "Adverse effect of antineoplastic drugs, initial encounter (pair with hair-loss code)"),
]
carc = [
    ("CO-4", "Modifier missing or doesn't match the code", "Add/fix RT/LT or other modifier; resubmit corrected claim"),
    ("CO-16", "Claim missing information", "Read the RARC remark code; fix and resubmit"),
    ("CO-18", "Duplicate claim", "Check whether the original was paid — don't rebill"),
    ("CO-22", "Another payer may be primary", "Get other insurance info; bill primary first"),
    ("CO-27", "Coverage ended before date of service", "Re-verify eligibility; bill correct payer or patient"),
    ("CO-29", "Timely filing limit expired", "Appeal only with proof you filed on time (clearinghouse report)"),
    ("CO-45", "Charge exceeds fee schedule (contractual adjustment)", "Not a denial — write off the difference"),
    ("CO-50", "Not medically necessary", "Appeal with order, mastectomy documentation, notes"),
    ("CO-96", "Non-covered charge", "If ABN/waiver signed (GA modifier), bill patient"),
    ("CO-97", "Included in another service (bundled)", "Review coding"),
    ("CO-109", "Not covered by this payer / wrong contractor", "Send to the correct payer (e.g., DME MAC)"),
    ("CO-119", "Benefit maximum reached", "Bill patient if allowed by contract"),
    ("CO-151", "Frequency / quantity exceeds what's allowed", "Check replacement schedule; appeal with documentation"),
    ("CO-197", "Prior authorization missing", "Request retro-auth or appeal"),
    ("CO-B7", "Provider not credentialed/eligible on date of service", "Check effective date on Applications tab; hold claims until effective"),
    ("PR-1", "Deductible", "Bill patient"),
    ("PR-2", "Coinsurance", "Bill patient"),
    ("PR-3", "Copay", "Bill patient"),
]
mods = [
    ("RT", "Right side"), ("LT", "Left side"), ("NU", "New equipment"),
    ("KX", "Medical policy requirements met (Medicare)"),
    ("GA", "ABN (waiver) on file — patient may be billed if denied"),
    ("GY", "Statutorily excluded / not a Medicare benefit (e.g., wigs)"),
    ("GZ", "Expected denial, no ABN on file"),
]
def ref_table(ws, start_col, hdrs, rows, widths):
    for j, (h, w) in enumerate(zip(hdrs, widths)):
        L = get_column_letter(start_col + j)
        c = ws[f"{L}1"]; c.value = h; c.font = f_hdr; c.fill = fill(TEAL if j == 0 else GRAY_HDR)
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[L].width = w
    for i, row in enumerate(rows, 2):
        for j, v in enumerate(row):
            c = ws.cell(row=i, column=start_col + j, value=v)
            c.font = f_bold if j == 0 else f_base
            c.alignment = Alignment(wrap_text=True, vertical="top")
ref_table(wsC, 1, ["HCPCS", "Description", "Billing note"], hcpcs, [9, 44, 34])
ref_table(wsC, 5, ["ICD-10", "Description"], icd, [10, 44])
ref_table(wsC, 8, ["Denial code (CARC)", "Meaning", "What to do"], carc, [11, 36, 42])
ref_table(wsC, 12, ["Modifier", "Meaning"], mods, [9, 36])
n = len(hcpcs) + 3
wsC[f"A{n}"] = ("Medicare rules for breast prostheses (replacement frequency, documentation, orders) are in the DME MAC "
                "'External Breast Prostheses' LCD and policy article — check before billing. Add lymphedema compression "
                "codes here if you bill them (Medicare covers compression garments since 1/1/2024). Codes and "
                "descriptions are short summaries; confirm against the current HCPCS/ICD-10 code sets.")
wsC[f"A{n}"].font = Font(name=FONT, size=9, italic=True, color="666666")
wsC[f"A{n}"].alignment = Alignment(wrap_text=True, vertical="top")
wsC.merge_cells(f"A{n}:C{n+4}")
wsC.freeze_panes = "A2"
wsC.sheet_properties.tabColor = TEAL
HCPCS_DV = "=Codes!$A$2:$A$40"; ICD_DV = "=Codes!$E$2:$E$40"; CARC_DV = "=Codes!$H$2:$H$40"

# =====================================================================
# Payers
# =====================================================================
wsP = wb.create_sheet("Payers")
MCARE_ROUTE = "Accreditation + $50k surety bond first, then CMS-855S in PECOS (National Provider Enrollment contractor)"
MCAID_ROUTE = "GAMMIS provider enrollment (mmis.georgia.gov)"
CMO_ROUTE = "Separate credentialing with the CMO after you have a GA Medicaid ID"
COMM_ROUTE = "Ancillary/DME provider application via provider relations or payer website; CAQH only if requested"
MA_ROUTE = "Separate contract with the plan; requires active Medicare PTAN first"
CMO_NOTE = "Confirm this CMO is still active — Georgia re-procured its Medicaid CMO contracts (check dch.georgia.gov)."
payers = [
    ["Medicare (DMEPOS) – Original Medicare", "Medicare", MCARE_ROUTE, "DME MAC Jurisdiction C (CGS) via clearinghouse/EDI", 365, 120, "Y", "pecos.cms.hhs.gov · cgsmedicare.com/jc", None, "Open", "Timely filing = 12 months from date of service. Appeal (redetermination) = 120 days from the date you receive the denial notice."],
    ["Georgia Medicaid (FFS / GAMMIS)", "Medicaid", MCAID_ROUTE, "GAMMIS web portal", 180, 30, "N", "mmis.georgia.gov", None, "Unknown", "Defaults are placeholders — confirm timely filing and appeal limits in the GA Medicaid provider manual."],
    ["GA Medicaid CMO – Wellpoint (Amerigroup)", "Medicaid CMO", CMO_ROUTE, "CMO portal / Availity", 180, 30, "N", None, None, "Unknown", CMO_NOTE],
    ["GA Medicaid CMO – CareSource", "Medicaid CMO", CMO_ROUTE, "CMO portal / clearinghouse", 180, 30, "N", None, None, "Unknown", CMO_NOTE],
    ["GA Medicaid CMO – Peach State Health Plan", "Medicaid CMO", CMO_ROUTE, "CMO portal / clearinghouse", 180, 30, "N", None, None, "Unknown", CMO_NOTE],
    ["Anthem BCBS Georgia – Commercial", "Commercial", COMM_ROUTE, "Availity", 90, 60, "N", "availity.com", None, "Unknown", "Largest commercial payer in GA — prioritize."],
    ["Aetna – Commercial", "Commercial", COMM_ROUTE, "Availity / clearinghouse", 90, 60, "N", "availity.com", None, "Unknown", None],
    ["UnitedHealthcare – Commercial", "Commercial", COMM_ROUTE, "UHC Provider Portal / clearinghouse", 90, 60, "N", None, None, "Unknown", None],
    ["Cigna – Commercial", "Commercial", COMM_ROUTE, "Clearinghouse / Cigna portal", 90, 60, "N", None, None, "Unknown", None],
    ["Ambetter from Peach State (Marketplace)", "Commercial", COMM_ROUTE, "Clearinghouse / Ambetter portal", 90, 60, "N", None, None, "Unknown", "ACA Marketplace plan."],
    ["Kaiser Permanente Georgia", "Commercial", COMM_ROUTE, "Kaiser portal", 90, 60, "N", None, None, "Unknown", None],
    ["Alliant Health Plans", "Commercial", COMM_ROUTE, "Clearinghouse", 90, 60, "N", None, None, "Unknown", "Georgia-based plan."],
    ["Humana – Medicare Advantage", "Medicare Advantage", MA_ROUTE, "Availity / clearinghouse", 90, 60, "N", None, None, "Unknown", None],
    ["UnitedHealthcare – Medicare Advantage", "Medicare Advantage", MA_ROUTE, "UHC Provider Portal / clearinghouse", 90, 60, "N", None, None, "Unknown", None],
    ["Aetna – Medicare Advantage", "Medicare Advantage", MA_ROUTE, "Availity / clearinghouse", 90, 60, "N", None, None, "Unknown", None],
    ["Wellcare – Medicare Advantage", "Medicare Advantage", MA_ROUTE, "Clearinghouse / Wellcare portal", 90, 60, "N", None, None, "Unknown", None],
    ["Anthem – Medicare Advantage", "Medicare Advantage", MA_ROUTE, "Availity", 90, 60, "N", "availity.com", None, "Unknown", None],
]
pcols = [
    dict(h="Payer", w=40, wrap=True, note="Payer names here feed every Payer dropdown. Add new payers at the bottom."),
    dict(h="Payer type", w=16, dv=LST("F")),
    dict(h="How to enroll", w=44, wrap=True),
    dict(h="Where claims go", w=30, wrap=True),
    dict(h="Timely filing (days from DOS)", w=13, fmt="0", note="Drives the timely-filing deadline on the Claims tab. Defaults (except Medicare) are conservative placeholders — replace with the number in your contract/provider manual, then set 'Verified?' to Y."),
    dict(h="Appeal window (days)", w=12, fmt="0", note="Drives the appeal deadline on the Claims tab. Conservative placeholder unless Verified = Y."),
    dict(h="Deadlines verified?", w=11, dv=LST("G")),
    dict(h="Website / portal", w=28, wrap=True),
    dict(h="Provider relations phone", w=18),
    dict(h="DME network status", w=14, dv=LST("M")),
    dict(h="Notes", w=50, wrap=True),
]
build_table(wsP, pcols, 60, rows=payers)
red(wsP, "G2:G61", '=AND($A2<>"",$G2="N")')
PAYER_DV = "=Payers!$A$2:$A$61"
PTYPE = '=IF({c}{{r}}="","",IFERROR(INDEX(Payers!$B$2:$B$61,MATCH({c}{{r}},Payers!$A$2:$A$61,0)),""))'

# =====================================================================
# Applications
# =====================================================================
wsA = wb.create_sheet("Applications")
acols = [
    dict(h="App ID", w=9),
    dict(h="Payer", w=36, dv=PAYER_DV, wrap=True),
    dict(h="Payer type", w=15, kind="auto", f=PTYPE.format(c="B")),
    dict(h="Application type", w=18, dv=LST("A")),
    dict(h="Status", w=22, dv=LST("B"), note="Update after every payer contact. 'Info requested' = payer asked for something — respond fast, these stall apps."),
    dict(h="Date started", w=11, fmt=DATE_FMT),
    dict(h="Date submitted", w=11, fmt=DATE_FMT),
    dict(h="Tracking / reference #", w=16),
    dict(h="Payer contact", w=16),
    dict(h="Contact phone / email", w=20),
    dict(h="Last contact", w=11, fmt=DATE_FMT),
    dict(h="Next follow-up", w=11, fmt=DATE_FMT, note="Rule of thumb: 14 days after each contact while an app is pending."),
    dict(h="Days in process", w=10, kind="auto", fmt="0", f='=IF(G{r}="","",IF(N{r}<>"",N{r}-G{r},TODAY()-G{r}))'),
    dict(h="Effective date", w=11, fmt=DATE_FMT, note="Don't bill in-network before this date."),
    dict(h="Provider ID / PTAN", w=14),
    dict(h="Contract + fee schedule?", w=11, dv=LST("G")),
    dict(h="EFT / ERA set up?", w=10, dv=LST("G")),
    dict(h="Revalidation / re-cred due", w=12, fmt=DATE_FMT, note="Medicare DMEPOS: every 3 years. Medicaid: at least every 5 years. Commercial: per contract (often 3 years)."),
    dict(h="In pipeline?", w=9, kind="auto", f='=IF(E{r}="","",IFERROR(INDEX(Lists!$C$2:$C$30,MATCH(E{r},Lists!$B$2:$B$30,0)),""))'),
    dict(h="Follow-up flag", w=14, kind="auto", f='=IF(S{r}<>"Y","",IF(L{r}="","Set follow-up date",IF(L{r}<TODAY(),"OVERDUE",IF(L{r}<=TODAY()+7,"This week",""))))'),
    dict(h="Notes", w=44, wrap=True),
]
app_rows = []
for i, p in enumerate(payers, 1):
    app_rows.append([f"APP-{i:03d}", p[0], None, "Initial enrollment", "Not started"])
A_N = build_table(wsA, acols, 300, rows=app_rows)
red(wsA, f"T2:T{A_N}", '=$T2="OVERDUE"')
yellow(wsA, f"T2:T{A_N}", '=$T2="This week"')
orange(wsA, f"T2:T{A_N}", '=$T2="Set follow-up date"')
green(wsA, f"E2:E{A_N}", '=$E2="Active / effective"')
red(wsA, f"E2:E{A_N}", '=OR($E2="Denied",$E2="Network closed")')
orange(wsA, f"E2:E{A_N}", '=$E2="Info requested"')

# =====================================================================
# Documents (expirables)
# =====================================================================
wsD = wb.create_sheet("Documents")
dcols = [
    dict(h="Document", w=40, wrap=True),
    dict(h="Category", w=15, dv=LST("K")),
    dict(h="Number / ID", w=16),
    dict(h="Issued by", w=20, wrap=True),
    dict(h="Issued / last done", w=11, fmt=DATE_FMT),
    dict(h="Expires / next due", w=11, fmt=DATE_FMT, note="Enter the expiration or next-due date. Status turns yellow at 90 days, orange at 30, red when expired."),
    dict(h="Renewal cycle", w=11, dv=LST("L")),
    dict(h="Days left", w=8, kind="auto", fmt="0", f='=IF(F{r}="","",F{r}-TODAY())'),
    dict(h="Status", w=14, kind="auto", f='=IF(A{r}="","",IF(F{r}="",IF(G{r}="None","No expiration","Enter date"),IF(F{r}<TODAY(),"EXPIRED",IF(F{r}-TODAY()<=30,"Due ≤30 days",IF(F{r}-TODAY()<=90,"Due ≤90 days","Current")))))'),
    dict(h="Needed for", w=22, wrap=True),
    dict(h="Google Drive link to copy", w=24, note="Paste the share link to the scanned copy in your Drive folder."),
    dict(h="Owner", w=10),
    dict(h="Notes", w=46, wrap=True),
]
docs = [
    ["LLC registration – GA Secretary of State", "Business", None, "GA Secretary of State", None, None, "Annual", None, None, "All payers", None, "Lisa", "Annual registration due Jan 1 – Apr 1 each year (ecorp.sos.ga.gov)."],
    ["EIN letter (IRS CP-575 or 147C)", "Identifiers", None, "IRS", None, None, "None", None, None, "All payers"],
    ["NPI – Type 2 (Bella256 Boutique LLC)", "Identifiers", None, "NPPES", None, None, "None", None, None, "All payers", None, "Lisa", "Payers enroll the business NPI. Update NPPES within 30 days of any change."],
    ["NPI – Type 1 (Lisa, individual)", "Identifiers", None, "NPPES", None, None, "None", None, None, "Some commercial payers"],
    ["Mastectomy fitter certification (Lisa)", "Licenses & certs", None, None, None, None, "Annual", None, None, "Accreditation, all payers", None, "Lisa", "Set the renewal cycle and CE requirements to match your certifying body."],
    ["Local business license / occupation tax certificate", "Licenses & certs", None, "City of Madison / Morgan County", None, None, "Annual", None, None, "Medicare, Medicaid"],
    ["GA sales tax certificate of registration", "Business", None, "GA Dept. of Revenue", None, None, "None", None, None, "Business"],
    ["DMEPOS accreditation certificate", "Accreditation", None, "(your accrediting org)", None, None, "Annual", None, None, "Medicare, Medicaid", None, "Lisa", "Required before Medicare enrollment. Surveys reportedly moved to annual in 2026 — confirm with your accreditor."],
    ["Medicare surety bond ($50,000)", "Insurance & bonds", None, "(surety company)", None, None, "Annual", None, None, "Medicare", None, "Lisa", "$50,000 per NPI, from a Treasury-listed surety."],
    ["Comprehensive liability insurance (≥ $300,000)", "Insurance & bonds", None, None, None, None, "Annual", None, None, "Medicare, most payers", None, "Lisa", "Medicare supplier standard: at least $300,000; covers place of business, customers and employees."],
    ["Professional liability insurance", "Insurance & bonds", None, None, None, None, "Annual", None, None, "Commercial payers"],
    ["Workers' comp (if you have employees)", "Insurance & bonds", None, None, None, None, "Annual", None, None, "As applicable"],
    ["Lease / proof of physical location", "Business", None, None, None, None, "Annual", None, None, "Medicare site visit"],
    ["W-9 (current)", "Banking", None, None, None, None, "None", None, None, "All payers"],
    ["Bank letter or voided check (EFT)", "Banking", None, None, None, None, "None", None, None, "All payers"],
    ["CAQH ProView attestation (only if a payer requires it)", "Compliance", None, "CAQH", None, None, "120 days", None, None, "Some commercial payers", None, "Lisa", "Re-attest every 120 days."],
    ["OIG LEIE + SAM.gov exclusion check (owner & staff)", "Compliance", None, "oig.hhs.gov / sam.gov", None, None, "Monthly", None, None, "Medicare, Medicaid", None, "Lisa", "Save a screenshot/PDF of each monthly search."],
    ["HIPAA policies + annual staff training", "Compliance", None, None, None, None, "Annual", None, None, "Accreditation"],
    ["Medicare PTAN – revalidation due", "Enrollment IDs", None, "CMS", None, None, "3 years", None, None, "Medicare", None, "Lisa", "DMEPOS suppliers revalidate every 3 years. Report changes within 30 days."],
    ["GA Medicaid provider ID – revalidation due", "Enrollment IDs", None, "GA DCH", None, None, "5 years", None, None, "Medicaid"],
]
D_N = build_table(wsD, dcols, 100, rows=docs)
red(wsD, f"I2:I{D_N}", '=$I2="EXPIRED"')
orange(wsD, f"I2:I{D_N}", '=$I2="Due ≤30 days"')
yellow(wsD, f"I2:I{D_N}", '=$I2="Due ≤90 days"')
green(wsD, f"I2:I{D_N}", '=$I2="Current"')

# =====================================================================
# Claims
# =====================================================================
wsCl = wb.create_sheet("Claims")
ccols = [
    dict(h="Claim ID", w=10),
    dict(h="Patient (name or ID)", w=16, note="Patient name or internal ID. OK under your Google BAA — keep the sheet in the BAA-covered Workspace account and share only with staff who need it."),
    dict(h="Date of service", w=11, fmt=DATE_FMT),
    dict(h="Payer", w=32, dv=PAYER_DV, wrap=True),
    dict(h="Payer type", w=14, kind="auto", f=PTYPE.format(c="D")),
    dict(h="HCPCS", w=9, dv=HCPCS_DV, dv_strict=False),
    dict(h="Modifier(s)", w=9),
    dict(h="Units", w=6, fmt="0"),
    dict(h="ICD-10", w=10, dv=ICD_DV, dv_strict=False),
    dict(h="Billed $", w=10, fmt=USD),
    dict(h="Date submitted", w=11, fmt=DATE_FMT),
    dict(h="Payer claim # (ICN)", w=16),
    dict(h="Status", w=20, dv=LST("D")),
    dict(h="Allowed $", w=10, fmt=USD),
    dict(h="Paid $", w=10, fmt=USD),
    dict(h="Patient resp. $", w=10, fmt=USD),
    dict(h="Processed date (paid/denied)", w=11, fmt=DATE_FMT),
    dict(h="Denial code", w=9, dv=CARC_DV, dv_strict=False, note="See the Codes tab for what each code means and what to do."),
    dict(h="Next step / notes", w=36, wrap=True),
    dict(h="Timely filing deadline", w=11, kind="auto", fmt=DATE_FMT, f='=IF(OR(C{r}="",D{r}=""),"",C{r}+IFERROR(INDEX(Payers!$E$2:$E$61,MATCH(D{r},Payers!$A$2:$A$61,0)),90))', note="Date of service + payer's timely-filing days (Payers tab). Falls back to 90 days."),
    dict(h="Days left to file", w=9, kind="auto", fmt="0", f='=IF(OR(T{r}="",K{r}<>""),"",T{r}-TODAY())', note="Only shows for claims not yet submitted."),
    dict(h="Appeal deadline", w=11, kind="auto", fmt=DATE_FMT, f='=IF(OR(Q{r}="",AND(M{r}<>"Denied",M{r}<>"Partially paid")),"",Q{r}+IFERROR(INDEX(Payers!$F$2:$F$61,MATCH(D{r},Payers!$A$2:$A$61,0)),60))', note="Processed date + payer's appeal window (Payers tab). Payers count from the date on the denial notice — use that date as the processed date."),
    dict(h="Open A/R?", w=8, kind="auto", f='=IF(M{r}="","",IFERROR(INDEX(Lists!$E$2:$E$30,MATCH(M{r},Lists!$D$2:$D$30,0)),""))'),
    dict(h="Days outstanding", w=10, kind="auto", fmt="0", f='=IF(OR(W{r}<>"Y",K{r}=""),"",TODAY()-K{r})'),
    dict(h="Aging bucket", w=9, kind="auto", f='=IF(X{r}="","",IF(X{r}<=30,"0–30",IF(X{r}<=60,"31–60",IF(X{r}<=90,"61–90","90+"))))'),
    dict(h="Open balance $", w=11, kind="auto", fmt=USD, f='=IF(W{r}<>"Y","",MAX(0,IF(N{r}="",J{r}-O{r},N{r}-O{r}-P{r})))', note="Amount still expected from the payer: allowed − paid − patient responsibility (or billed − paid before the claim is processed)."),
]
claim_ex = ["EXAMPLE", "P-0001", date(2026, 11, 3), "Anthem BCBS Georgia – Commercial", None, "L8030", "RT", 1, "Z90.11", 350, date(2026, 11, 4), "26308XXXXXX", "Pending", None, None, None, None, None, "Example row — shows the format. Leave it or delete it; the Dashboard ignores it."]
C_N = build_table(wsCl, ccols, 1000, example=claim_ex)
red(wsCl, f"M2:M{C_N}", '=$M2="Denied"')
green(wsCl, f"M2:M{C_N}", '=$M2="Paid in full"')
red(wsCl, f"V2:V{C_N}", '=AND($V2<>"",$M2="Denied",$V2-TODAY()<=14)')
red(wsCl, f"U2:U{C_N}", '=AND($U2<>"",$U2<=30)')
red(wsCl, f"Y2:Y{C_N}", '=$Y2="90+"')
orange(wsCl, f"Y2:Y{C_N}", '=$Y2="61–90"')

# =====================================================================
# Auths & Eligibility
# =====================================================================
wsAu = wb.create_sheet("Auths & Eligibility")
aucols = [
    dict(h="Record ID", w=10),
    dict(h="Patient (name or ID)", w=16),
    dict(h="Payer", w=32, dv=PAYER_DV, wrap=True),
    dict(h="Member ID", w=14),
    dict(h="Date checked", w=11, fmt=DATE_FMT),
    dict(h="Coverage active?", w=9, dv=LST("G")),
    dict(h="Deductible remaining $", w=11, fmt=USD),
    dict(h="DME benefit / coinsurance notes", w=30, wrap=True),
    dict(h="Auth required?", w=10, dv=LST("O")),
    dict(h="HCPCS", w=9, dv=HCPCS_DV, dv_strict=False),
    dict(h="Auth status", w=12, dv=LST("H")),
    dict(h="Auth #", w=14),
    dict(h="Auth start", w=11, fmt=DATE_FMT),
    dict(h="Auth end", w=11, fmt=DATE_FMT),
    dict(h="Units approved", w=9, fmt="0"),
    dict(h="Days until auth ends", w=10, kind="auto", fmt="0", f='=IF(N{r}="","",N{r}-TODAY())'),
    dict(h="Notes / call ref #", w=36, wrap=True),
]
au_ex = ["EXAMPLE", "P-0001", "Anthem BCBS Georgia – Commercial", "4821", date(2026, 11, 2), "Y", 250, "DME 20% after deductible; in-network only", "No", "L8030", "Not required", None, None, None, None, None, "Ref# from Availity eligibility check"]
AU_N = build_table(wsAu, aucols, 300, example=au_ex)
orange(wsAu, f"P2:P{AU_N}", '=AND($P2<>"",$P2<=14,$K2="Approved")')
red(wsAu, f"F2:F{AU_N}", '=$F2="N"')

# =====================================================================
# Superbills
# =====================================================================
wsS = wb.create_sheet("Superbills")
scols = [
    dict(h="Superbill #", w=10),
    dict(h="Patient (name or ID)", w=16),
    dict(h="Date issued", w=11, fmt=DATE_FMT),
    dict(h="Date of service", w=11, fmt=DATE_FMT),
    dict(h="Items / HCPCS", w=24, wrap=True),
    dict(h="ICD-10", w=10, dv=ICD_DV, dv_strict=False),
    dict(h="Total charged $", w=11, fmt=USD),
    dict(h="Paid by patient $", w=11, fmt=USD),
    dict(h="Payment method", w=11, dv=LST("N")),
    dict(h="Patient's insurer", w=20),
    dict(h="Medicare patient?", w=9, dv=LST("G")),
    dict(h="Alert", w=28, kind="auto", wrap=True, f='=IF(K{r}="Y","Medicare: not reimbursable — supplier not Medicare-enrolled","")'),
    dict(h="Given to patient (date)", w=11, fmt=DATE_FMT),
    dict(h="Patient reimbursed?", w=11, dv=LST("O")),
    dict(h="Amount reimbursed $", w=11, fmt=USD),
    dict(h="Notes", w=36, wrap=True),
]
sb_ex = ["EXAMPLE", "P-0002", date(2026, 11, 5), date(2026, 11, 5), "L8000 x2; L8030-RT x1", "Z90.11", 425, 425, "Card", "Aetna (patient's plan)", "N", None, date(2026, 11, 5), "Pending", None, "Ask patient to report back what their insurer paid."]
S_N = build_table(wsS, scols, 500, example=sb_ex)
red(wsS, f"L2:L{S_N}", '=$L2<>""')

# =====================================================================
# Call Log
# =====================================================================
wsLog = wb.create_sheet("Call Log")
lcols = [
    dict(h="Entry ID", w=10),
    dict(h="Date", w=11, fmt=DATE_FMT),
    dict(h="Payer", w=32, dv=PAYER_DV, wrap=True),
    dict(h="Rep name", w=14),
    dict(h="Phone / channel", w=16),
    dict(h="Call reference #", w=16, note="Always ask for a call reference number — it's your proof if a payer later says you were told something different."),
    dict(h="Related to", w=14, dv=LST("J")),
    dict(h="Related ID (APP-/claim #)", w=12),
    dict(h="What was said", w=44, wrap=True),
    dict(h="Next action", w=28, wrap=True),
    dict(h="Next action due", w=11, fmt=DATE_FMT),
    dict(h="Done?", w=7, dv=LST("G")),
    dict(h="Flag", w=11, kind="auto", f='=IF(OR(K{r}="",L{r}="Y"),"",IF(K{r}<TODAY(),"OVERDUE",IF(K{r}<=TODAY()+7,"This week","")))'),
]
log_ex = ["EXAMPLE", date(2026, 11, 6), "Aetna – Commercial", "Jordan", "Provider services", "1106-XXXX", "Application", "APP-007", "Asked whether the DME network is open in GA. Rep said open; sent link to ancillary application.", "Submit ancillary application", date(2026, 11, 13), "N"]
LOG_N = build_table(wsLog, lcols, 1000, example=log_ex)
red(wsLog, f"M2:M{LOG_N}", '=$M2="OVERDUE"')
yellow(wsLog, f"M2:M{LOG_N}", '=$M2="This week"')

# =====================================================================
# Roadmap
# =====================================================================
wsR = wb.create_sheet("Roadmap")
road = [
    ("Foundation", "Confirm LLC is active with GA Secretary of State", "ecorp.sos.ga.gov. Annual registration due Jan 1 – Apr 1 every year."),
    ("Foundation", "EIN letter on file", "IRS CP-575 or 147C letter — every application asks for it."),
    ("Foundation", "Organizational (Type 2) NPI for the LLC", "nppes.cms.hhs.gov. Payers enroll the business NPI. Pick the right taxonomy (e.g., 332B00000X DME & Medical Supplies)."),
    ("Foundation", "Business bank account + bank letter / voided check", "Needed for EFT (direct deposit) on every enrollment."),
    ("Foundation", "Local business license / occupation tax certificate", "City of Madison / Morgan County."),
    ("Foundation", "Comprehensive liability insurance ≥ $300,000", "Medicare supplier standard (42 CFR 424.57(c)(10)): must cover place of business, customers and employees."),
    ("Foundation", "Written policies & procedures", "HIPAA privacy/security, complaint resolution, patient rights, delivery & instructions, returns/warranty. Required by supplier standards and accreditation."),
    ("Foundation", "Set up monthly exclusion checks", "OIG LEIE (oig.hhs.gov/exclusions) + SAM.gov for owner and any staff. Log on Documents tab."),
    ("Foundation", "Choose clearinghouse / billing software", "Needed before your first in-network claim: sends electronic claims (837P), returns remittances (835 ERA). Check it supports DME/HCPCS, the CGS DME MAC, GA Medicaid, eligibility checks, and signs a BAA."),
    ("Medicare DMEPOS", "Confirm no enrollment moratorium is active", "Check CMS 'Provider Enrollment Moratoria' page. The 2026 moratorium on new DMEPOS medical supply companies expired 8/27/2026; CMS can reimpose one."),
    ("Medicare DMEPOS", "Confirm posted-hours requirement", "Supplier standards require being open to the public at least 30 hrs/week unless an exemption applies — confirm before applying (matters for appointment-only boutiques)."),
    ("Medicare DMEPOS", "Pick an accreditation organization & get quotes", "CMS-approved: BOC, ACHC, The Compliance Team, HQAA, CHAP, The Joint Commission."),
    ("Medicare DMEPOS", "Pass accreditation survey; save accreditation letter", "Required before Medicare will enroll you. Add to Documents tab with expiration."),
    ("Medicare DMEPOS", "Buy $50,000 surety bond", "Per NPI, from a Treasury-listed surety company. Renews yearly."),
    ("Medicare DMEPOS", "Submit CMS-855S in PECOS + pay application fee", "pecos.cms.hhs.gov. Goes to the National Provider Enrollment contractor. Log as APP-001 on Applications tab."),
    ("Medicare DMEPOS", "Pass the unannounced site visit", "Be open during posted hours; signage, inventory, patient files and policies on site."),
    ("Medicare DMEPOS", "Receive PTAN + effective date", "Record on Applications and Documents tabs."),
    ("Medicare DMEPOS", "Set up EFT (CMS-588)", "Medicare pays by EFT only."),
    ("Medicare DMEPOS", "Decide participating vs. non-participating", "CMS-460 participation agreement = accept assignment on all Medicare claims."),
    ("Medicare DMEPOS", "EDI enrollment with CGS (DME MAC Jurisdiction C)", "cgsmedicare.com/jc — Georgia DME claims go to CGS. Link your clearinghouse and set up ERA."),
    ("Medicare DMEPOS", "Calendar revalidation + 30-day change reporting", "Revalidate every 3 years. Report ownership, address, hours changes within 30 days in PECOS."),
    ("Georgia Medicaid", "Check GA Medicaid DME enrollment requirements", "mmis.georgia.gov → Provider Enrollment. Confirm whether Medicare enrollment or accreditation is required first."),
    ("Georgia Medicaid", "Submit GAMMIS enrollment application (DME)", "Upload NPI, W-9, licenses, accreditation, insurance, bank info."),
    ("Georgia Medicaid", "Receive Medicaid provider ID; set up EFT / ERA", "Record on Applications and Documents tabs."),
    ("Georgia Medicaid", "Credential with each Georgia Medicaid CMO", "Separate application per CMO. Confirm the current CMO list at dch.georgia.gov (contracts were re-procured)."),
    ("Georgia Medicaid", "Calendar Medicaid revalidation", "At least every 5 years."),
    ("Commercial", "Call provider relations: is the DME/ancillary network open?", "Log in Call Log with reference #. Update 'DME network status' on Payers tab."),
    ("Commercial", "Request and submit ancillary (DME) application", "Some payers use CAQH ProView — only set it up if a payer asks."),
    ("Commercial", "Follow up every 14 days", "Update Applications tab after every contact."),
    ("Commercial", "Review contract & fee schedule before signing", "Compare rates for your HCPCS codes to your ABC cost."),
    ("Commercial", "Get effective date + provider ID — don't bill in-network before it", "Claims before the effective date deny (CO-B7)."),
    ("Commercial", "Register on portals / Availity; set up EFT & ERA", "Availity covers eligibility, claims and auths for many GA payers."),
    ("Commercial", "Calendar re-credentialing", "Usually every 2–3 years per contract."),
    ("Medicare Advantage", "Get Medicare PTAN first", "Most MA plans require active Medicare enrollment."),
    ("Medicare Advantage", "Apply to each MA plan separately", "Humana, UHC, Aetna, Wellcare, Anthem. Some route DME through a network vendor — ask."),
    ("Medicare Advantage", "Track contracts & effective dates", "Applications tab."),
    ("Ongoing", "Monday: review the Dashboard", "Overdue follow-ups, expiring documents, denials, appeal deadlines."),
    ("Ongoing", "Log every payer call with a reference #", "Call Log tab."),
    ("Ongoing", "Monthly OIG / SAM exclusion checks", "Documents tab row — update the date each month."),
    ("Ongoing", "CAQH re-attestation every 120 days (if used)", "Documents tab."),
    ("Ongoing", "Annual renewals", "Liability insurance, surety bond, accreditation, business license, GA annual registration."),
    ("Ongoing", "Keep records 7 years", "Orders, proof of delivery, claims. Medicare requires proof of delivery be kept 7 years."),
]
rcols = [
    dict(h="Track", w=17),
    dict(h="Step", w=6, kind="auto", f='=IF(A{r}="","",COUNTIF($A$2:A{r},A{r}))'),
    dict(h="Task", w=44, wrap=True),
    dict(h="Details / where", w=62, wrap=True),
    dict(h="Status", w=14, dv=LST("I")),
    dict(h="Date done", w=11, fmt=DATE_FMT),
    dict(h="Notes", w=30, wrap=True),
]
road_rows = [[t, None, task, det, "Not started"] for t, task, det in road]
R_N = build_table(wsR, rcols, len(road) + 20, rows=road_rows, freeze="C2")
green(wsR, f"E2:E{R_N}", '=$E2="Done"')
yellow(wsR, f"E2:E{R_N}", '=OR($E2="In progress",$E2="Waiting on payer")')

# =====================================================================
# Dashboard
# =====================================================================
wsDash = wb.create_sheet("Dashboard", 0)
wsDash.sheet_properties.tabColor = INK
for col, w in {"A": 2, "B": 40, "C": 14, "D": 4, "E": 40, "F": 14, "G": 10}.items():
    wsDash.column_dimensions[col].width = w
wsDash["B1"] = "Bella256 Insurance & Credentialing Dashboard"; wsDash["B1"].font = f_title
wsDash["B2"] = "As of"; wsDash["B2"].font = f_bold
wsDash["C2"] = "=TODAY()"; wsDash["C2"].number_format = DATE_FMT; wsDash["C2"].font = f_bold
wsDash["E2"] = "Everything here is automatic — update the other tabs and this page recalculates."
wsDash["E2"].font = Font(name=FONT, size=9, italic=True, color="666666")

def section(cell_rng_label, row, col_label, title):
    c = wsDash[f"{col_label}{row}"]; c.value = title; c.font = f_h2; c.fill = fill(ROSE_DK)
    nxt = chr(ord(col_label) + 1)
    wsDash[f"{nxt}{row}"].fill = fill(ROSE_DK)
    if col_label == "E":
        wsDash[f"G{row}"].fill = fill(ROSE_DK)

def kpi(row, col, label, formula, fmt="0", alert=None):
    lc = wsDash[f"{col}{row}"]; lc.value = label; lc.font = f_base; lc.border = border
    vcol = chr(ord(col) + 1)
    vc = wsDash[f"{vcol}{row}"]; vc.value = formula; vc.font = f_bold; vc.number_format = fmt; vc.border = border
    vc.alignment = Alignment(horizontal="right")
    if alert:
        red(wsDash, f"{vcol}{row}", alert.format(c=f"{vcol}{row}"))
    return vc

NA = "<>EXAMPLE"
# --- Enrollment
section(None, 4, "B", "ENROLLMENT APPLICATIONS")
r = 5
for i in range(10):
    lc = wsDash[f"B{r}"]; lc.value = f"=Lists!$B${i+2}"; lc.font = f_base; lc.border = border
    vc = wsDash[f"C{r}"]; vc.value = f'=COUNTIF(Applications!$E$2:$E$1000,B{r})'; vc.font = f_bold; vc.border = border
    r += 1
kpi(16, "B", "Follow-ups OVERDUE", '=COUNTIF(Applications!$T$2:$T$1000,"OVERDUE")', alert="={c}>0")
kpi(17, "B", "Follow-ups due this week", '=COUNTIF(Applications!$T$2:$T$1000,"This week")')
kpi(18, "B", "In-pipeline apps with no follow-up date", '=COUNTIF(Applications!$T$2:$T$1000,"Set follow-up date")', alert="={c}>0")
# --- Documents
section(None, 20, "B", "DOCUMENTS & EXPIRATIONS")
kpi(21, "B", "EXPIRED", '=COUNTIF(Documents!$I$2:$I$500,"EXPIRED")', alert="={c}>0")
kpi(22, "B", "Due in ≤30 days", '=COUNTIF(Documents!$I$2:$I$500,"Due ≤30 days")', alert="={c}>0")
kpi(23, "B", "Due in ≤90 days", '=COUNTIF(Documents!$I$2:$I$500,"Due ≤90 days")')
kpi(24, "B", "Missing expiration date", '=COUNTIF(Documents!$I$2:$I$500,"Enter date")')
kpi(25, "B", "Current", '=COUNTIF(Documents!$I$2:$I$500,"Current")')
# --- Roadmap
section(None, 27, "B", "ROADMAP PROGRESS (% done)")
tracks = ["Foundation", "Medicare DMEPOS", "Georgia Medicaid", "Commercial", "Medicare Advantage"]
for i, t in enumerate(tracks):
    row = 28 + i
    lc = wsDash[f"B{row}"]; lc.value = t; lc.font = f_base; lc.border = border
    vc = wsDash[f"C{row}"]
    vc.value = (f'=IFERROR(COUNTIFS(Roadmap!$A$2:$A$200,B{row},Roadmap!$E$2:$E$200,"Done")/'
                f'(COUNTIF(Roadmap!$A$2:$A$200,B{row})-COUNTIFS(Roadmap!$A$2:$A$200,B{row},Roadmap!$E$2:$E$200,"N/A")),0)')
    vc.number_format = "0%"; vc.font = f_bold; vc.border = border

# --- Claims
CL = lambda col: f"Claims!${col}$2:${col}$2000"
section(None, 4, "E", "CLAIMS")
kpi(5, "E", "Open claims (#)", f'=COUNTIFS({CL("W")},"Y",{CL("A")},"{NA}")')
kpi(6, "E", "Open balance expected from payers", f'=SUMIFS({CL("Z")},{CL("W")},"Y",{CL("A")},"{NA}")', fmt=USD)
for col, txt in (("E", "Aging (days outstanding)"), ("F", "$ open"), ("G", "# claims")):
    c = wsDash[f"{col}7"]; c.value = txt; c.font = Font(name=FONT, size=9, bold=True, color="666666")
    c.alignment = Alignment(horizontal="left" if col == "E" else "right")
for i, b in enumerate(["0–30", "31–60", "61–90", "90+"]):
    row = 8 + i
    lc = wsDash[f"E{row}"]; lc.value = b; lc.font = f_base; lc.border = border
    vc = wsDash[f"F{row}"]; vc.value = f'=SUMIFS({CL("Z")},{CL("Y")},E{row},{CL("A")},"{NA}")'
    vc.number_format = USD; vc.font = f_bold; vc.border = border
    gc = wsDash[f"G{row}"]; gc.value = f'=COUNTIFS({CL("Y")},E{row},{CL("A")},"{NA}")'
    gc.font = f_bold; gc.border = border
red(wsDash, "F11:G11", "=$G$11>0")
kpi(12, "E", "Denied — needs action", f'=COUNTIFS({CL("M")},"Denied",{CL("A")},"{NA}")', alert="={c}>0")
kpi(13, "E", "Appeal deadlines in next 14 days", f'=COUNTIFS({CL("V")},">="&TODAY(),{CL("V")},"<="&TODAY()+14,{CL("M")},"Denied",{CL("A")},"{NA}")', alert="={c}>0")
kpi(14, "E", "Appeal deadlines MISSED", f'=COUNTIFS({CL("V")},"<"&TODAY(),{CL("M")},"Denied",{CL("A")},"{NA}")', alert="={c}>0")
kpi(15, "E", "Unsent claims ≤30 days from timely-filing limit", f'=COUNTIFS({CL("U")},"<=30",{CL("A")},"{NA}")', alert="={c}>0")
kpi(16, "E", "Paid by payers this year", f'=SUMIFS({CL("O")},{CL("Q")},">="&DATE(YEAR(TODAY()),1,1),{CL("A")},"{NA}")', fmt=USD)
# --- Auths
AU = lambda col: f"'Auths & Eligibility'!${col}$2:${col}$1000"
section(None, 18, "E", "AUTHS & ELIGIBILITY")
kpi(19, "E", "Auths requested / pending", f'=COUNTIFS({AU("K")},"Requested",{AU("A")},"{NA}")+COUNTIFS({AU("K")},"Pending",{AU("A")},"{NA}")')
kpi(20, "E", "Approved auths ending in ≤14 days", f'=COUNTIFS({AU("K")},"Approved",{AU("P")},"<=14",{AU("P")},">=0",{AU("A")},"{NA}")', alert="={c}>0")
# --- Superbills
SB = lambda col: f"Superbills!${col}$2:${col}$2000"
section(None, 22, "E", "SUPERBILLS (THIS YEAR)")
kpi(23, "E", "Superbills issued", f'=COUNTIFS({SB("C")},">="&DATE(YEAR(TODAY()),1,1),{SB("A")},"{NA}")')
kpi(24, "E", "Total charged on superbills", f'=SUMIFS({SB("G")},{SB("C")},">="&DATE(YEAR(TODAY()),1,1),{SB("A")},"{NA}")', fmt=USD)
# --- Call log
LG = lambda col: f"'Call Log'!${col}$2:${col}$2000"
section(None, 26, "E", "CALL LOG")
kpi(27, "E", "Next actions OVERDUE", f'=COUNTIFS({LG("M")},"OVERDUE",{LG("A")},"{NA}")', alert="={c}>0")
kpi(28, "E", "Next actions due this week", f'=COUNTIFS({LG("M")},"This week",{LG("A")},"{NA}")')

# =====================================================================
# Start Here
# =====================================================================
wsH = wb.create_sheet("Start Here", 0)
wsH.sheet_properties.tabColor = TEAL
wsH.column_dimensions["A"].width = 2
wsH.column_dimensions["B"].width = 28
wsH.column_dimensions["C"].width = 100
rows_h = [
    ("title", "Bella256 Insurance & Credentialing Tracker"),
    ("p", "One place for payer applications, credentialing documents, claims, authorizations, superbills and payer calls. Built for a small DME / mastectomy-products supplier in Georgia."),
    ("h", "PUT IT IN GOOGLE DRIVE"),
    ("kv", "1. Upload", "Sign in to your BAA-covered Google Workspace account → drive.google.com → New → File upload → choose this file."),
    ("kv", "2. Convert", "Right-click the file → Open with → Google Sheets → File → Save as Google Sheets."),
    ("kv", "3. Clean up", "Delete the original .xlsx from Drive so there's only one copy to update."),
    ("kv", "4. Documents folder", "Make a Drive folder 'Credentialing Docs' for scans (license, bond, insurance, accreditation). Paste each file's link on the Documents tab."),
    ("h", "PRIVACY / HIPAA — READ THIS"),
    ("kv", "BAA status", "Google BAA in place — patient names, member IDs and claim details are OK in this sheet."),
    ("kv", "Keep it covered", "Store it only in the Google Workspace account covered by the BAA — never a personal @gmail.com account. Share with named people only (no 'anyone with the link'). Turn on 2-Step Verification. Don't download copies to personal devices."),
    ("kv", "Still leave out", "SSNs and anything not needed to track the application or claim."),
    ("kv", "Website repo copy", "The copy in the website repo is a blank template. Never put real patient or business data there."),
    ("h", "HOW TO READ THE TABS"),
    ("kv", "Rose header", "You type here."),
    ("kv", "Gray header / gray cells", "Automatic formulas — don't type over them. (In Google Sheets, drag the last row's formulas down if you run out of rows.)"),
    ("kv", "EXAMPLE rows", "Gray italic rows show the format. The Dashboard ignores them. Delete them whenever you like."),
    ("kv", "Red / orange / yellow", "Red = overdue or expired. Orange = due within 30 days or needs attention. Yellow = coming up."),
    ("h", "TABS"),
    ("kv", "Dashboard", "Weekly check: overdue follow-ups, expiring documents, open claims and aging, denials, appeal deadlines."),
    ("kv", "Roadmap", "Step-by-step checklist for Medicare DMEPOS, Georgia Medicaid, commercial and Medicare Advantage enrollment."),
    ("kv", "Payers", "Every payer you're targeting, how to enroll, where claims go, and timely-filing/appeal limits (these drive deadlines on the Claims tab)."),
    ("kv", "Applications", "One row per payer application. Status, dates, reference numbers, follow-ups, effective date, provider IDs."),
    ("kv", "Documents", "Everything that expires: accreditation, surety bond, insurance, licenses, CAQH, exclusion checks, revalidations."),
    ("kv", "Claims", "One row per claim line. Status, payments, denials, appeal and timely-filing deadlines, aging."),
    ("kv", "Auths & Eligibility", "Benefit checks and prior authorizations."),
    ("kv", "Superbills", "Superbills issued under the self-pay model (out-of-network reimbursement for commercial plans)."),
    ("kv", "Call Log", "Every payer call: rep, call reference #, what was said, next action."),
    ("kv", "Codes", "HCPCS, ICD-10, modifiers and denial codes with what to do."),
    ("kv", "Lists", "Dropdown options. Add payers on the Payers tab; add other options here."),
    ("h", "WEEKLY ROUTINE (15 MIN)"),
    ("kv", "Monday", "Open Dashboard → clear anything red → call payers with overdue follow-ups → log each call → update statuses."),
    ("kv", "After every payer contact", "Update the Applications/Claims row + add a Call Log entry with the reference #."),
    ("kv", "Monthly", "Run OIG + SAM exclusion checks; update the date on the Documents tab."),
    ("h", "VERIFY BEFORE RELYING ON"),
    ("p", "Timely-filing and appeal limits on the Payers tab are conservative placeholders except Medicare (12 months filing / 120-day redetermination). Replace them with numbers from each contract or provider manual and mark 'Deadlines verified?' = Y. Code descriptions on the Codes tab are short summaries. Regulations change — check CMS and Georgia DCH before applying."),
    ("h", "SOURCES"),
    ("kv", "CMS", "Enroll as a DMEPOS Supplier — cms.gov/Medicare/Provider-Enrollment-and-Certification/Enroll-as-a-DMEPOS-Supplier"),
    ("kv", "Moratorium", "Federal Register 2026-03971 (DMEPOS moratorium, effective 2/27/2026); AOPA: 'CMS's Six-Month Moratorium on DMEPOS Supplier Enrollment has Expired' (8/27/2026)"),
    ("kv", "Supplier standards", "42 CFR 424.57 (liability insurance, posted hours, change reporting)"),
    ("kv", "Feature model", "Credentialing-software feature set (payer enrollment tracking, expirable monitoring, alerts) as listed for Accel Health on Capterra/GetApp"),
]
r = 1
for item in rows_h:
    kind = item[0]
    if kind == "title":
        wsH.cell(row=r, column=2, value=item[1]).font = f_title
        wsH.row_dimensions[r].height = 28
    elif kind == "p":
        c = wsH.cell(row=r, column=2, value=item[1]); c.font = f_base
        c.alignment = Alignment(wrap_text=True, vertical="top")
        wsH.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        wsH.row_dimensions[r].height = 30 if len(item[1]) < 200 else 58
    elif kind == "h":
        r += 1
        c = wsH.cell(row=r, column=2, value=item[1]); c.font = f_h2; c.fill = fill(ROSE_DK)
        wsH.cell(row=r, column=3).fill = fill(ROSE_DK)
    elif kind == "kv":
        a = wsH.cell(row=r, column=2, value=item[1]); a.font = f_bold; a.alignment = Alignment(vertical="top", wrap_text=True)
        b = wsH.cell(row=r, column=3, value=item[2]); b.font = f_base; b.alignment = Alignment(wrap_text=True, vertical="top")
        a.border = border; b.border = border
        if len(item[2]) > 110:
            wsH.row_dimensions[r].height = 28
    r += 1

# tab order
order = ["Start Here", "Dashboard", "Roadmap", "Payers", "Applications", "Documents", "Claims",
         "Auths & Eligibility", "Superbills", "Call Log", "Codes", "Lists"]
wb._sheets = [wb[n] for n in order]
wb.active = 0
wb.save(OUT)
print("saved", OUT)
