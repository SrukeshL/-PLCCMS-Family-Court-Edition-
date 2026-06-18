# ⚖ Family Court Custody Dashboard — Quick Start Guide
## Pennsylvania Family Court | 23 Pa.C.S. § 5328 (eff. Aug 2025)

---

## WHAT IS THIS?
A private, local tool to help you organize your custody case — whether you are representing
yourself (Pro Se) or working with an attorney. All data stays on YOUR computer.

---

## STEP 1 — INSTALL PYTHON (one time only)
1. Download Python from https://python.org/downloads
2. During install, check ✅ "Add Python to PATH"
3. Open Command Prompt (Windows) or Terminal (Mac/Linux)
4. Run: `pip install openpyxl pandas`

---

## STEP 2 — FILL IN THE EXCEL TEMPLATE
Open `family_court_custody_template.xlsx` and fill in the **yellow cells**:

| Sheet | What to Enter |
|-------|--------------|
| **Dashboard Config** | Your case name, docket number, court, party labels |
| **Case Overview** | Basic case details and hearing dates |
| **PA Custody Factors §5328** | Your evidence for each legal factor |
| **Evidence Tracker** | All documents, photos, videos (add Google Drive links) |
| **Timeline of Events** | All key dates and events in chronological order |
| **Communications Log** | Emails, texts, calls with other party and attorneys |
| **Cross-Examination Builder** | Questions to ask witnesses at hearing |
| **Hearing Prep Checklist** | Tasks to complete before your hearing |
| **Attorney Strategy [CONF]** | Strategy notes (attorney-client privileged) |
| **ODC Violation Tracker [CONF]** | Document any attorney misconduct |
| **Negotiation & Settlement** | Compare positions for out-of-court resolution |
| **Judge Pre-Trial View** | 48-hour advance summary for the court |

---

## STEP 3 — GENERATE YOUR DASHBOARD
1. Put ALL files in the SAME folder:
   - `family_court_custody_template.xlsx`
   - `generate_dashboard.py`
2. Open Command Prompt / Terminal in that folder
3. Run: `python generate_dashboard.py`
4. Open `family_court_dashboard.html` in your browser ✅

---

## STEP 4 — OPTIONAL: ENABLE ONLINE TRACKING
To share your dashboard online and track who views it:
1. Set "Tracking Enabled" = TRUE in Dashboard Config sheet
2. Set up a Google Apps Script webhook (see full documentation)
3. Paste the webhook URL in "Google Sheet Webhook" in config
4. Regenerate your dashboard
5. Host the HTML file on any web server or Google Sites

---

## CONFIDENTIALITY TIERS
| 🟢 Public | Safe to share with all parties and court |
| 🟡 Restricted | Your attorney only |
| 🔴 Confidential | Court only, upon request with signed agreement |
| ⚫ Sealed | Court order required — physical copies only |

**Medical and financial records:** NEVER upload to any online system.
Keep physical copies only. Reference them as "Available upon court request."

---

## PRIORITY LEVELS
| Level | Meaning |
|-------|---------|
| 🟢 LOW | Background / supporting information |
| 🔵 MEDIUM | Notable, relevant to the case |
| 🟠 HIGH | Significant impact on custody decision |
| 🔴 CRITICAL | Directly affects custody determination |
| ⚫ CATASTROPHIC | Immediate safety / requires urgent court action |

---

## PA CUSTODY FACTORS — 23 Pa.C.S. § 5328 (Updated Aug 2025)
Courts must consider ALL factors. **Safety factors (1)(2)(2.1)(2.2)** receive
**substantial weighted consideration**. No single factor is determinative.
Awards are **gender neutral** per § 5328(b).

---

## ODC — ATTORNEY MISCONDUCT
If your attorney violates the Rules of Professional Conduct:
- Document in the **ODC Violation Tracker** sheet
- Contact: Office of Disciplinary Counsel
- Website: www.padisciplinaryboard.org
- Philadelphia: (215) 560-6296
- Pittsburgh: (412) 565-3173
- ODC enforces ethical rules only — they do NOT represent you

---

## 48-HOUR PRE-TRIAL SUBMISSION
Complete the **Judge Pre-Trial View** sheet at least 48 hours before your hearing.
Generate your dashboard and submit to the court per your judge's instructions.
This helps the judge review your case in advance and makes the 12 minutes
of court time far more effective.

---

## DISCLAIMER
This tool is for case organization purposes only. It does not constitute legal advice.
Consult a licensed Pennsylvania attorney for legal guidance.
Pro Se resources: Philadelphia Family Court Self-Help Center — 1501 Arch Street, Philadelphia PA.

---

*Family Court Custody Dashboard | Pennsylvania | 23 Pa.C.S. § 5328*
