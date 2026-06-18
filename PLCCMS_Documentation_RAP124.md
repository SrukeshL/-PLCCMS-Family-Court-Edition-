# ⚖️ Parent/Lawyer Case Custody Management System (PLCCMS)
## Technical & User Documentation
### Formatted per Pennsylvania Rules of Appellate Procedure, Pa. R.A.P. 124

---

**Court:** Philadelphia Family Court | Commonwealth of Pennsylvania, USA
**Statute:** 23 Pa.C.S. § 5328 — Amended Act 11, June 30, 2025 (Effective August 2025)
**License:** GNU General Public License v3 — Public Good Software
**Purpose:** Free Forever — For Every Parent — For Every Court — For Every Child

---

## IMPORTANT NOTICES

> **PRIMARY PURPOSE**
> This tool is designed to help organize and increase efficiency while navigating through a custody case in family court. It does not provide legal advice and does not guarantee any case outcome.

> **NOT LEGAL ADVICE**
> This application does not create an attorney-client relationship. Always consult a licensed attorney in your jurisdiction for legal advice specific to your situation.

> **LOCAL USE FIRST**
> This application is designed for self-service and runs locally on your own computer only. Before making it available on the Internet or any shared network, kindly consult with qualified network security, cybersecurity, and Internet security professionals or experienced IT engineers to ensure your case data and personal information are properly protected.

> **AI TOOLS NOTICE**
> AI tools integrated in this application are for data processing and validation assistance ONLY. They do NOT guarantee any case outcome. AI tools can and do make mistakes. Never rely on AI output as legal advice. Always verify AI suggestions with a licensed attorney. Delete all AI chat sessions after your work session is complete.

> **CONFIDENTIALITY**
> Medical and financial records must never be uploaded to any online system. Keep physical copies only. Reference them in the application as "Available upon court request — physical copy maintained."

> **FREE PUBLIC GOOD SOFTWARE**
> This application is licensed under GNU GPL v3. It is free forever for all parents, courts, government bodies, legal aid organizations, and nonprofits. It cannot be commercialized by any party. All improvements must be shared back publicly.

---

## TABLE OF CONTENTS

| Section | Title |
|---------|-------|
| 1 | Introduction & Purpose |
| 2 | System Overview |
| 3 | Pennsylvania Custody Factors — 23 Pa.C.S. § 5328 |
| 4 | Application Features & Modules |
| 5 | Priority & Confidentiality System |
| 6 | ODC Professional Conduct Module |
| 7 | Single & Dual Party Mode |
| 8 | 48-Hour Pre-Trial Judicial View |
| 9 | AI Integration — Optional Processing Assistant |
| 10 | Privacy, Security & Local Deployment |
| 11 | Database & Court System Integration |
| 12 | GitHub Repository & Upload Guide |
| 13 | Government / Court Adoption |
| 14 | Installation & Quick Start |
| 15 | Roadmap & Future Development |
| 16 | License — GNU GPL v3 |
| 17 | Legal Disclaimer |
| Appendix A | PA Rules of Professional Conduct Reference |
| Appendix B | ODC Contact Information |
| Appendix C | Glossary of Terms |

---

## 1. Introduction & Purpose

### 1.1 Background

The Parent/Lawyer Case Custody Management System (PLCCMS) was developed in response to a critical and well-documented inefficiency in the family court system: the average family court judge handles in excess of two hundred (200) cases per week. At a standard eight-hour workday over five days, this results in approximately twelve (12) minutes of judicial time per custody case — twelve minutes in which a judge must review evidence, hear both parties, and render a decision that will determine a child's living arrangement, education, healthcare, and future.

### 1.2 Purpose Statement

This tool is designed to help organize and increase efficiency while navigating through a custody case in family court. It serves parents representing themselves (Pro Se), parents working with attorneys, attorneys in active custody proceedings, mediators, and — through its 48-hour pre-trial view — judges who benefit from structured advance case summaries.

### 1.3 Design Principles

The application was designed around seven core principles:

- **Privacy First:** All data remains on the user's local computer by default. No data is transmitted to any external server without explicit user configuration.
- **Legal Accuracy:** All custody factors reflect the current Pennsylvania statute, 23 Pa.C.S. § 5328, as amended by Act 11, effective August 2025.
- **Universal Access:** The application is free for all users regardless of race, culture, income, or legal representation status.
- **Flexibility:** Single-party use is fully supported. The other party may join at any stage of the proceedings without loss of any existing data.
- **Court Compatibility:** Output is structured for compatibility with Philadelphia Family Court procedures and designed to scale to any jurisdiction.
- **Attorney Accountability:** A professional conduct reference module strengthens the attorney-client relationship through transparency and shared understanding.
- **Public Good:** The application is released under GNU GPL v3 — it cannot be commercialized and will remain free forever.

### 1.4 Who This Application Serves

| User | How It Helps |
|------|-------------|
| Pro Se Parents | Organizes your entire case without legal training — guided by actual PA statute |
| Attorney-Represented Parents | Collaborate with your attorney more efficiently and stay fully informed |
| Family Law Attorneys | Strategy builder, client collaboration, evidence organizer, hearing prep |
| Both Parties Together | When both parents use the same format, a unified judge's view is generated |
| Mediators | Side-by-side negotiation view accelerates out-of-court settlements |
| Judges — 48-Hr Pre-Trial | Structured advance summary reduces cold-reading of disorganized files |
| Court Systems | Free adoption — GPL v3 ensures all improvements return to the public |
| Legal Aid Organizations | Free distribution to clients who cannot afford attorneys |

---

## 2. System Overview

### 2.1 Architecture

PLCCMS operates as a three-component local application:

| Component | File | Function |
|-----------|------|----------|
| Data Entry | family_court_custody_template.xlsx | 14-sheet Excel workbook. Parent or attorney enters case data into yellow input cells. |
| Generator | generate_dashboard.py | Python script that reads the workbook and produces a formatted HTML application. |
| Application Output | family_court_dashboard.html | Court-ready HTML application opened in any browser. Fully local — no server required. |
| Admin Panel | admin_panel.html | Optional security tracking panel for monitoring access when dashboard is shared online. |

### 2.2 Deployment Modes

| Mode | Description | Recommended For |
|------|-------------|-----------------|
| Local / Offline | Runs entirely on parent's computer. No internet required. Maximum privacy. | All individual parents and attorneys. Default mode. |
| Online / Shared | HTML hosted on web server. Access tracking via Google Sheets webhook. Security monitoring enabled. | Institutional use, court submission, or attorney-client sharing. Requires IT security review before deployment. |

### 2.3 Technology Stack

| Component | Technology | Purpose |
|-----------|-----------|---------|
| Data Entry | Microsoft Excel / LibreOffice Calc | Structured case data input |
| Processing | Python 3.8+ | Application generation |
| Excel Library | openpyxl | Read/write Excel workbooks |
| Data Processing | pandas | Data manipulation and processing |
| Output | HTML5 / CSS3 / JavaScript | Browser-based application |
| Tracking | Google Sheets + Apps Script | Optional access monitoring |
| Database | PostgreSQL or Oracle | Optional institutional storage |
| License | GNU GPL v3 | Free public good software |

---

## 3. Pennsylvania Custody Factors — 23 Pa.C.S. § 5328

### 3.1 Statutory Authority

All custody factor tracking in PLCCMS is based on Title 23, Chapter 53, Section 28 of the Pennsylvania Consolidated Statutes, as amended by Act 8 of 2024 and Act 11 of 2025 (effective sixty days from June 30, 2025, i.e., approximately August 28, 2025). The application pre-loads all current factors and displays the applicable statutory language alongside each evidence input field.

### 3.2 The Best Interest Standard

Under 23 Pa.C.S. § 5328(a), the court shall determine the best interest of the child by considering all relevant factors, giving substantial weighted consideration to those factors specified in paragraphs (1), (2), (2.1), and (2.2) which affect the safety of the child. No single factor is by itself determinative — the court examines the totality of the circumstances. 23 Pa.C.S. § 5328(a.2).

### 3.3 Safety Factors — Substantial Weighted Consideration

The following four factors receive substantial weighted consideration in all custody determinations and are displayed with elevated visual prominence in the application:

| Factor | Legal Standard | App Priority Range |
|--------|---------------|-------------------|
| (1) | Which party is more likely to ensure the safety of the child. | Critical — Catastrophic |
| (2) | The present and past abuse committed by a party or member of the party's household, including PFA and SVPO orders with abuse findings. | Critical — Catastrophic |
| (2.1) | Information under § 5329.1(a) — child abuse and involvement with protective services. | Critical — Catastrophic |
| (2.2) | Violent or assaultive behavior committed by a party. | Critical — Catastrophic |

### 3.4 All Remaining Factors

| Factor | Legal Standard |
|--------|---------------|
| (2.3i) | Which party more likely to encourage and permit frequent and continuing contact between the child and the other party, if consistent with child safety needs. |
| (2.3ii) | Attempts to turn the child against the other party (parental alienation) — except where reasonable safety measures are necessary. Good faith safety efforts do not constitute unwillingness to cooperate. |
| (3) | Willingness and ability to prioritize child's needs — care, stability, continuity; daily physical, emotional, developmental, educational, and special needs. |
| (4) | Need for stability and continuity in education, family life, and community life — except when change is necessary for safety. |
| (6) | The child's sibling and other familial relationships. |
| (7) | The well-reasoned preference of the child, based on developmental stage, maturity, and judgment. |
| (11) | The proximity of the residences of the parties. |
| (12) | Each party's employment schedule and availability, or ability to make appropriate childcare arrangements. |
| (14) | History of drug or alcohol abuse of a party or household member. |
| (15) | The mental and physical condition of a party or household member. |
| (16) | Any other relevant factor. |

### 3.5 Special Provisions

**Abuse Exception — § 5328(a.1):**
A factor shall not be adversely weighed against a party if the circumstances related to the factor were in response to abuse or necessary to protect the child or abused party, and the party alleging abuse does not pose a risk to the safety of the child at the time of the custody hearing. Temporary housing instability resulting from abuse shall not be considered against the party alleging abuse.

**Gender Neutral — § 5328(b):**
No party shall receive preference based upon gender in any custody award.

**30-Day Notice — § 5328(d):**
Within thirty (30) days of the receipt of a custody complaint, petition for modification, or petition to intervene, the court shall provide all parties with a copy of this section. PLCCMS includes a 30-day deadline tracker for this requirement.

**Factors Deleted by Prior Amendment:**
Factors (5), (8), (9), (10), and (13) were deleted by prior legislative amendment and are not included in the application.

---

## 4. Application Features & Modules

### 4.1 Excel Workbook — 14 Sheets

The application data entry workbook contains fourteen (14) structured sheets. Users fill in yellow-highlighted input cells only. Headers, formulas, and structure must not be modified.

| # | Sheet Name | Purpose | Access Level |
|---|-----------|---------|-------------|
| 1 | Dashboard Config | App settings, state selection, party mode, webhook URL | Admin |
| 2 | Case Overview | Court, docket number, parties, hearing dates, judge assigned | Standard |
| 3 | PA Custody Factors §5328 | All legal factors — your evidence, opposing claims, rebuttal | Standard |
| 4 | Evidence Tracker | Documents, photos, videos, court filings with Drive links | Standard |
| 5 | Timeline of Events | Chronological master record of all case events | Standard |
| 6 | Communications Log | All interactions — emails, texts, calls, letters | Standard |
| 7 | Cross-Examination Builder | Questions organized by witness and custody factor | Confidential |
| 8 | Hearing Prep Checklist | 48-hour pre-trial task tracker with due dates | Standard |
| 9 | Attorney Strategy [CONF] | Lawyer-client strategy — attorney-client privileged | Confidential |
| 10 | PA Rules Prof Conduct Ref | Professional conduct rules reference for client and attorney | Standard |
| 11 | ODC Violation Tracker [CONF] | Attorney misconduct documentation for potential ODC complaint | Confidential |
| 12 | Judge Pre-Trial View | 48-hour advance judicial summary — factual and concise | Standard |
| 13 | Negotiation & Settlement | Out-of-court resolution — side-by-side position tracker | Confidential |
| 14 | Visitor Log | Paste access tracking data from Google Sheets webhook | Admin |

### 4.2 Cross-Examination Builder

The Cross-Examination Builder provides a structured framework for preparing questions for any witness. Questions are organized into eight (8) sections:

- **Section A:** Background and Foundation Questions
- **Section B:** Parenting History and Daily Care
- **Section C:** Safety and Abuse Allegations
- **Section D:** Co-Parenting and Cooperation
- **Section E:** Communications and Access to Child
- **Section F:** Medical and Educational Decisions
- **Section G:** Financial Responsibilities
- **Section H:** Credibility and Inconsistencies

Each question is linked to a specific § 5328 custody factor and assigned a priority level. The builder includes fields for the expected answer, supporting evidence or citation, and question status (Draft / Ready / Used).

### 4.3 48-Hour Pre-Trial Judicial View

The Judge Pre-Trial View sheet generates a structured case summary designed for judicial advance review. Submission forty-eight (48) hours before the scheduled hearing enables the judge to review evidence, prepare questions, and identify agreed versus disputed issues before trial — maximizing the effective use of the twelve (12) minutes of available judicial time per case.

The view captures twelve structured sections: case summary, safety concerns, custody arrangement requested, the three strongest supporting points, key evidence, disputed facts, agreed items, the child's current situation, proposed parenting schedule, outstanding motions, witness summary, and relief requested.

### 4.4 Negotiation and Settlement Module

The Negotiation and Settlement tracker enables both parties to compare positions on fifteen (15) standard custody issues side by side. Agreed items require no court time. Disputed items are flagged for judicial attention. This module is designed to facilitate out-of-court resolution and reduce the volume of contested matters requiring judicial determination.

### 4.5 Key Application Features Summary

| Feature | Description |
|---------|-------------|
| 100% Local & Private | Runs on your computer — no cloud, no login, no data transmission |
| PA § 5328 Built-In | All custody factors pre-loaded, updated for Act 11 (August 2025) |
| State Dropdown | Select your state — custody factors and rules update automatically |
| Philadelphia Court Colors | Navy blue, gold, teal — aligned with Philadelphia Family Court palette |
| 5-Level Priority System | Low → Medium → High → Critical → Catastrophic |
| Single or Dual Party | Start alone; other party joins anytime — no data lost |
| Judge's 48-Hr View | Pre-trial summary designed for judicial advance review |
| Cross-Exam Builder | 8-section question builder organized by witness and topic |
| Evidence Tracker | Google Drive links + 4-tier confidentiality system |
| Settlement Negotiator | Side-by-side position comparison to find compromise zones |
| ODC Conduct Module | Professional conduct reference — builds trust, ensures accountability |
| AI Integration | Optional connection to Claude, ChatGPT, Gemini, or Copilot |
| Security Tracking | Optional Google Sheets webhook to track access when shared online |
| Database Support | Optional PostgreSQL or Oracle for institutional and court use |
| Court System Integration | API-ready for existing court management systems |

---

## 5. Priority & Confidentiality System

### 5.1 Five-Level Priority System

All evidence entries, timeline events, and custody factor assessments are assigned one of five priority levels. Priority levels are set collaboratively by the parent and attorney and may be adjusted at any time.

| Level | Color Code | When to Apply |
|-------|-----------|---------------|
| LOW | Green | Background or supporting context — not central to case outcome |
| MEDIUM | Blue | Notable and relevant — supports the overall picture |
| HIGH | Orange | Significant impact — court will likely consider carefully |
| CRITICAL | Red | Directly affects the custody determination |
| CATASTROPHIC | Dark Red | Immediate child safety concern — requires urgent court intervention |

### 5.2 Four-Tier Document Confidentiality System

All documents and evidence in PLCCMS are assigned one of four confidentiality tiers governing access and sharing:

| Tier | Icon | Who May Access | How Shared |
|------|------|----------------|------------|
| Public | 🟢 | All parties, attorneys, court, judge | Application directly or via Google Drive link |
| Restricted | 🟡 | Your attorney only | Password-protected section or in-person review only |
| Confidential | 🔴 | Court only, upon formal request | Physical copy with signed confidentiality agreement |
| Sealed | ⚫ | Judge only | Court order required — physical copies only |

> **Important:** Medical and financial records must never be uploaded to any online system. They must be maintained as physical copies only and referenced in the application as "Available upon court request — physical copy maintained by [Party Name]."

---

## 6. ODC Professional Conduct Module

### 6.1 Purpose and Rationale

The Professional Conduct Reference module was designed primarily to strengthen the attorney-client relationship — not to undermine it. Custody proceedings require parents to share the most sensitive personal information of their lives with their attorney. This level of trust is only possible when clients understand what their attorney has promised to do and what professional obligations govern the relationship.

The module serves eight distinct purposes:

**(1) Building Trust:**
Showing parents the rules their attorney is bound by — in the same application they use every day — demonstrates that the legal profession has a framework of enforceable obligations. This builds confidence rather than suspicion.

**(2) Transparency Protects Both Parties:**
When both parent and attorney see the same professional conduct rules in the same application, misunderstandings about fees, communication frequency, and scope of representation are resolved before they become disputes — protecting both the client and the attorney.

**(3) Prevention Over Complaint:**
Most ODC complaints arise from miscommunication rather than genuine misconduct. Early education on Rule 1.4 (Communication) and Rule 1.5 (Fees) sets correct expectations on both sides and prevents problems before they arise.

**(4) Protecting Good Attorneys:**
The structured ODC tracker requires specific dates, specific rule references, specific evidence, and meets the clear and convincing evidence standard. This structured format protects honest attorneys from vague or emotional complaints while ensuring genuine violations are properly documented.

**(5) Addressing the Knowledge Imbalance:**
Parents in custody proceedings are frequently emotionally overwhelmed and legally inexperienced. Having professional conduct rules visible in the same application they use daily informs the relationship without creating adversarial dynamics.

**(6) Broader Proceeding Integrity:**
Attorney misconduct affecting court proceedings — such as presenting false evidence (Rule 3.3) or withholding discovery (Rule 3.4) — may be documented by any party, raising the standard of practice across the entire proceeding.

**(7) Court System Alignment:**
The PA Supreme Court and ODC's own public materials encourage transparency and client education. Including the Rules of Professional Conduct aligns this application with values the court system already endorses.

**(8) Documentation While Active:**
ODC inquiry letters reveal that most complaints are poorly documented — missing timelines, no evidence, no fee agreement copies. This tracker solves the documentation problem while the attorney-client relationship is still active — not after it has irreparably broken down.

> *"The Professional Conduct Reference is included to strengthen the attorney-client relationship — not to undermine it. When both parent and attorney understand the same professional obligations, communication improves, expectations align, and focus stays where it belongs: on the best outcome for the child."*

### 6.2 Key Pennsylvania Rules of Professional Conduct

| Rule | Title | Most Relevant To |
|------|-------|-----------------|
| 1.1 | Competence | Attorney preparation and legal knowledge |
| 1.2 | Scope of Representation | Strategy decisions — attorney vs. client authority |
| 1.3 | Diligence | Timely filings, deadlines, responsiveness |
| 1.4 | Communication | Keeping client informed — most commonly violated rule |
| 1.5 | Fees | Reasonable fees, written fee agreements, billing |
| 1.6 | Confidentiality | Protection of all client information |
| 1.7 | Conflict of Interest | Multiple clients, personal relationships |
| 1.15 | Safekeeping of Property | Client retainer funds and accounting |
| 1.16 | Declining/Terminating | Withdrawal obligations and adequate notice |
| 3.3 | Candor to Tribunal | Truthfulness and accuracy before the court |
| 3.4 | Fairness to Opposing Party | Evidence and discovery obligations |
| 8.4 | Misconduct | Fraud, dishonesty — most serious violations |

### 6.3 ODC Limitations

The following limitations are prominently displayed in the application:

- ODC enforces the Rules of Professional Conduct only
- ODC does not represent the complainant personally
- ODC cannot obtain refunds or damages on the complainant's behalf
- ODC cannot interfere with pending legal proceedings
- Required evidentiary standard: clear and convincing evidence
- Complaints must be submitted as one email with one attachment — no zipped files
- Respond to all ODC inquiries within fourteen (14) days

---

## 7. Single & Dual Party Mode

### 7.1 Design Flexibility

One of the most important design principles of PLCCMS is that the application works fully for one parent acting alone and seamlessly expands when the other party joins — at any point in the case, without any loss of existing data.

### 7.2 Single Party Mode

In Single Party Mode (the default), one parent completes the workbook and generates the application independently. All Party B columns are present and waiting in the workbook. The generated application displays a factual, neutral court note: "Opposing party has not yet participated in pre-trial submission." This is informational — not adversarial.

### 7.3 Switching to Dual Party Mode — 3 Steps

When the other party decides to participate — at any stage of the case — the transition requires three steps:

1. Provide the other party with a copy of the blank Excel template
2. In the Dashboard Config sheet, set Party Mode to DUAL and enter the other party's filename in Party B Excel File
3. Run: `python generate_dashboard.py`

The unified view is generated immediately. All existing data is preserved exactly as entered.

### 7.4 Benefits of Dual Party Participation

| Benefit | Description |
|---------|-------------|
| Unified Judge's View | Judge sees both parties side by side — agreements in green, disputes in red |
| Faster Hearing | Agreed items require no court time — only genuine disputes need judicial attention |
| Settlement Acceleration | Both parties see each other's positions — enables informed out-of-court negotiation |
| Better Preparation | Both parties are required to articulate and evidence their positions clearly |
| Reduced Animosity | Structured format reduces emotional presentation — focuses on facts and law |

---

## 8. 48-Hour Pre-Trial Judicial View

The 48-hour pre-trial submission transforms the judicial experience from cold-reading a disorganized file to reviewing a structured, prioritized case summary in advance of the hearing.

The Judge Pre-Trial View sheet captures twelve (12) structured sections: case summary, safety concerns, custody arrangement requested, the three strongest supporting points, key evidence, disputed facts, agreed items, the child's current situation, proposed parenting schedule, outstanding motions, witness summary, and relief requested.

When both parties submit 48-hour views, the merge script generates a unified judicial application distinguishing agreements from disputes and surfacing Critical and Catastrophic priority items at the top of the judge's view.

The 48-hour workflow is as follows:

| Time | Action |
|------|--------|
| T-48 hours | Both parties (or one party) complete their Excel workbook |
| T-47 hours | Generate and submit dashboard to court via directed method |
| T-24 hours | Judge reviews application, prepares questions, identifies disputes |
| Trial Day | Judge arrives prepared — 12 minutes used for verification and ruling |
| Post-Trial | Ruling documented; both parties receive decision record |

---

## 9. AI Integration — Optional Processing Assistant

### 9.1 Critical Disclaimer

> ⚠️ **AI tools are for data processing and validation assistance ONLY. They do NOT guarantee any case outcome. AI tools can and do make mistakes. Never rely on AI output as legal advice. Always verify AI suggestions with a licensed attorney. Delete all AI chat sessions after your work session is complete.**
>
> This tool is designed for self-service and runs locally on your own computer only. Before making it available on the Internet or any network, consult with qualified network security, cybersecurity, and Internet security professionals or experienced IT engineers to ensure your case data and personal information are properly protected.

### 9.2 Supported AI Tools

| Tool | Provider | Recommended For | Subscription |
|------|----------|-----------------|-------------|
| Claude | Anthropic | Document analysis, legal writing, structured summaries | Claude Pro for large case files |
| ChatGPT | OpenAI | General drafting, question building, timeline summaries | ChatGPT Plus for large files |
| Google Gemini | Google | Google Drive evidence integration | Google One AI Premium |
| Microsoft Copilot | Microsoft | Word/Excel document integration | Microsoft 365 Copilot |

> **Subscription Note:** For cases involving large volumes of documents — hundreds of emails, lengthy transcripts, or multiple court filings — a paid AI subscription ($20–$30/month) is strongly recommended. Free tier context limits may produce incomplete or inaccurate results. Monthly AI subscription costs represent a fraction of one hour of attorney time. Cancel when your case concludes.

### 9.3 Permitted Uses

| Task | How AI Assists |
|------|---------------|
| Validate evidence entries | Verify evidence description clearly supports the custody factor cited |
| Review cross-exam questions | Suggest follow-up questions and identify logical gaps |
| Summarize communications | Condense lengthy email threads into clear timeline entries |
| Plain to court language | Translate everyday language into court-appropriate formal language for filings and submissions |
| Draft ODC complaint narrative | Structure the factual account clearly and chronologically |
| Check factor coverage | Identify which § 5328 factors have weak or no supporting evidence |
| Review judge's summary | Ensure the pre-trial view is factual, clear, and appropriately concise |
| Negotiation positions | Identify potential compromise zones based on factor analysis |

### 9.4 What AI Must NOT Be Used For

| Task | Reason |
|------|--------|
| Predicting case outcome | AI cannot know your judge, local practices, or full context |
| Legal strategy decisions | Requires a licensed attorney |
| Confirming case facts | AI can hallucinate — always independently verify |
| Medical, financial, or sealed documents | Never paste these into any AI tool |
| Permanent case data storage | Delete chat after every session — no exceptions |

### 9.5 Privacy Protocol — Follow Every Session

- Use the AI tool's temporary chat or incognito/no-history mode when available
- Use the child's initials only — never their full name
- Do not paste medical records, financial records, or sealed court documents into any AI tool
- Do not include attorney-client privileged strategy content
- Delete the entire conversation immediately after the session ends — without exception
- Verify the AI tool's current privacy policy regarding sensitive legal data before use

**How to delete your AI chat history:**

| Tool | Delete Steps |
|------|-------------|
| Claude | Settings → Delete Conversation or use temporary chat mode |
| ChatGPT | Sidebar → three dots next to chat → Delete |
| Gemini | myactivity.google.com → Delete activity |
| Copilot | Three dots → Clear conversation |

---

## 10. Privacy, Security & Local Deployment

### 10.1 Local Mode — Default and Recommended

In local mode — the default for all individual users — all case data remains on the user's computer only. No internet connection is required to run the application. No accounts, cloud synchronization, or analytics are enabled. The application is safe for sensitive legal, medical, and personal information.

### 10.2 Online Mode — Institutional Use Only

Before any online deployment, the user must obtain a review from qualified cybersecurity and IT security professionals. Family court case data is among the most sensitive personal information that exists. Improper online deployment can expose case strategy, child information, and personal safety to unauthorized parties.

When properly configured for online use, optional security features include:

- IP address and geographic access logging
- Timestamp recording of every page view and tab accessed
- Browser, device, and session duration tracking
- SHA-256 file hash tamper detection
- Parent notification of access events
- CSV export of all access logs for court submission if needed

---

## 11. Database & Court System Integration

### 11.1 Two-Tier Architecture

| Tier | Description | Recommended For |
|------|-------------|----------------|
| Tier 1 — Local | Excel → Python → HTML. No database required. Maximum privacy. | All individual parents and attorneys |
| Tier 2 — Institutional | Application → PostgreSQL or Oracle → Court portal. Full audit trail. | Courts, legal aid organizations, institutional deployments |

### 11.2 Database Options

| Database | Best For | Notes |
|----------|----------|-------|
| PostgreSQL | Open source courts, legal aid organizations, smaller jurisdictions | Free, powerful, widely supported |
| Oracle | Large court systems, state and federal government institutions | Used by many PA and federal government systems |

### 11.3 Court System Compatibility

- Tyler Technologies Odyssey — most common court management system in PA
- eCourt systems
- PACER — federal matters
- Custom court APIs via REST interface
- 48-hour pre-trial data push to court portal

---

## 12. GitHub Repository & Upload Guide

### 12.1 Files Safe to Upload

The following files are appropriate for the public GitHub repository:

- `README.md` — project documentation and homepage
- `LICENSE` — GNU GPL v3 license text
- `NOTICE` — copyright notice
- `CONTRIBUTING.md` — contribution guidelines
- `CHANGELOG.md` — version history
- `requirements.txt` — Python library list
- `.gitignore` — must be present and uploaded first
- `family_court_custody_template.xlsx` — blank template only — verify no personal data before uploading
- `generate_dashboard.py` — main Python generator script
- `scripts/merge_dashboards.py` — merge utility when ready
- `scripts/google_apps_script.js` — tracking script with no credentials
- `docs/` — all documentation files
- `jurisdiction/pennsylvania.json` — PA factors data file
- `screenshots/` — generic demo images only — no personal data

### 12.2 Files That Must Never Be Uploaded

- Any Excel workbook you have filled in with your own case data
- Generated HTML dashboard files — contain personal case data
- `admin_panel.html` — contains your webhook URL
- `output/` folder contents
- Any file containing personal names, docket numbers, or case details
- Your child's name, photos, school records, or any child information
- Medical records of any kind
- Financial documents of any kind
- Court filings with personal information
- Attorney-client communications
- ODC complaint documents
- Visitor logs containing IP addresses
- Google Sheets webhook URLs
- Any API keys, tokens, or credentials
- `.env` files or private configuration files

### 12.3 Pre-Commit Safety Checklist

Before every commit to GitHub, verify:

- `.gitignore` is in place and tested
- Excel template has all yellow cells blank — no personal data
- `git status` reviewed — only expected files showing
- No personal names, docket numbers, or case details in any file
- No API keys, webhook URLs, or credentials in any file
- Screenshots contain only generic placeholder data

---

## 13. Government & Court Adoption

### 13.1 GPL v3 and Government Use

PLCCMS is released under GNU GPL v3, which explicitly permits and supports free government and court adoption. Courts and government agencies may use, deploy, customize, and integrate this application at no cost.

### 13.2 What Courts and Government Agencies May Do

- Use this application internally — completely free
- Deploy to judges, clerks, and court staff — no license fees ever
- Customize for their specific court format and requirements
- Integrate with existing court management systems
- Mandate its use or recommend it to parties in cases
- Translate to other languages for their jurisdiction
- Host on government servers with appropriate IT security review
- Adopt as an official court standard or recommended tool

### 13.3 Government Obligations Under GPL v3

- Keep the original copyright notice in all copies
- Share any modifications or improvements publicly under GPL v3
- Cannot make the software proprietary or restrict access
- Must pass the same freedoms to all downstream users

### 13.4 The Ripple Effect

When a court system adopts and improves this application, those improvements must be shared back publicly under GPL v3. Every parent everywhere benefits from every court's enhancements — permanently and automatically.

| Level | Adoption | GPL v3 Benefit |
|-------|----------|---------------|
| Philadelphia Family Court | Adopts, improves 48-hr submission format | Shares back — all parents benefit |
| PA Supreme Court | Adopts statewide, improves factor weighting | Shares back — national visibility |
| Other states | Adapt for their custody factors | Share state-specific modules back |
| International courts | Adapt for Hague Convention | Share international modules back |

---

## 14. Installation & Quick Start

### 14.1 Prerequisites

- Python 3.8 or higher — download from python.org/downloads
- During installation, check "Add Python to PATH"
- Microsoft Excel or LibreOffice Calc (free alternative)
- Any modern web browser (Chrome, Firefox, Edge, Safari)

### 14.2 Installation Steps

```bash
# Step 1 — Clone the repository
git clone https://github.com/YOUR_USERNAME/PLCCMS.git
cd PLCCMS

# Step 2 — Install required Python libraries
pip install -r requirements.txt

# Step 3 — Open Excel template and fill in yellow cells
# Start with Sheet 1 — Dashboard Config

# Step 4 — Generate your application
python generate_dashboard.py

# Step 5 — Open in browser
# Windows:   start family_court_dashboard.html
# Mac:       open family_court_dashboard.html
# Linux:     xdg-open family_court_dashboard.html
```

### 14.3 First-Time Setup — Dashboard Config Sheet

| Setting | What to Enter |
|---------|--------------|
| Case Name | Your case nickname or name |
| Case Docket Number | Official court docket number |
| Court Name | Philadelphia Family Court or your court |
| Jurisdiction | Pennsylvania, USA |
| Party A Label | Your role — Petitioner, Mother, Father |
| Party B Label | Other party role |
| Party Mode | SINGLE (default) or DUAL |
| Tracking Enabled | FALSE (default) or TRUE for online use |

---

## 15. Roadmap & Future Development

| Phase | Features | Status |
|-------|----------|--------|
| 1 — Pennsylvania | PA § 5328 factors (Act 11, Aug 2025), 14-sheet workbook, HTML generator, single/dual party, ODC module, GPL v3 | Complete |
| 1 — Pending | Merge script, PDF export (Pa. R.A.P. 124), Google Apps Script webhook, state dropdown | In Progress |
| 2 — Multi-State | NJ, NY, DE, MD, FL, generic jurisdiction module, multilingual support | Planned |
| 3 — Institutional | PostgreSQL/Oracle connectors, REST API, Tyler Odyssey compatibility | Planned |
| 4 — International | Hague Convention module, Canada/UK/Australia/India, multi-language interface | Future |

---

## 16. License — GNU General Public License v3

PLCCMS is licensed under the GNU General Public License, Version 3 (GPL v3). This license was selected specifically because it guarantees that this application will remain free public good software — permanently and irrevocably.

| Permission / Obligation | Detail |
|------------------------|--------|
| Free for all parents | Any parent may use, copy, and run this application at no cost — forever |
| Free for courts and government | Any court, government body, or agency may use, deploy, and customize — free |
| Free for legal aid and nonprofits | Any legal aid organization or nonprofit may use and distribute — free |
| Improvements must be shared | Anyone who modifies and distributes must share modifications publicly under GPL v3 |
| Cannot be made proprietary | No party may convert this software into a closed, paid, or restricted product |
| Cannot be commercialized | No party may sell this software or build a paid commercial product from it |
| Copyright notice preserved | Original author's copyright notice must remain in all copies and derivatives |
| Free forever | These terms cannot be revoked — free public good software permanently |

> This project was created as purely public good software — to help organize and increase efficiency while navigating through a custody case in family court, for every parent, regardless of race, culture, income, or legal representation. It will remain free forever.

---

## 17. Legal Disclaimer

This application is not legal advice and does not create an attorney-client relationship.

This tool is designed to help organize and increase efficiency while navigating through a custody case in family court. It does not guarantee any particular custody outcome.

Pennsylvania statute information is provided for reference only. Always verify current law with a licensed Pennsylvania attorney. Statutes may be amended at any time — users are responsible for confirming that the factors and rules reflected in the application remain current at the time of their hearing.

The developers of this application are not attorneys and cannot provide legal advice. AI tool integration is for data processing and validation assistance only — AI tools can make mistakes. For legal advice specific to your situation, consult a licensed Pennsylvania family law attorney.

For Pro Se assistance:
- Philadelphia Family Court Self-Help Center — 1501 Arch Street, Philadelphia, PA 19102
- PA Courts Self-Help Center — www.pacourts.us/help-center

---

## Appendix A — PA Rules of Professional Conduct Reference

Source: Pennsylvania Rules of Professional Conduct — Supreme Court of Pennsylvania.
Violations may be reported to the Office of Disciplinary Counsel: www.padisciplinaryboard.org

| Rule | Title | Client Right | Common Violations to Watch For | ODC Reportable |
|------|-------|-------------|-------------------------------|----------------|
| 1.1 | Competence | Attorney must have the knowledge, skill, thoroughness, and preparation necessary for your representation. | Failing to file timely motions; inadequate preparation; missing deadlines | Yes |
| 1.2 | Scope of Representation | Attorney must abide by your decisions about objectives. You decide whether to accept settlement. | Taking actions without your authorization; ignoring client instructions | Yes |
| 1.3 | Diligence | Attorney must act with reasonable diligence and promptness. | Prolonged inaction; missed deadlines; 9+ months without filing | Yes |
| 1.4 | Communication | Attorney must keep you informed, respond promptly, and explain matters so you can make decisions. | Not returning calls; not informing you of hearings; sending communications without your knowledge | Yes — Most commonly violated |
| 1.5 | Fees | Fees must be reasonable. You are entitled to a written fee agreement. | Billing for work not done; refusing itemized billing; excessive fees | Yes |
| 1.6 | Confidentiality | Attorney may not reveal your information without your informed consent. | Disclosing strategy to opposing counsel; sharing private information | Yes |
| 1.7 | Conflict of Interest | Attorney cannot represent you if there is a conflict with another client or personal interest. | Representing both parties; personal relationship with opposing counsel | Yes |
| 1.15 | Safekeeping of Property | Attorney must keep your funds separate and provide accounting. | Commingling funds; not accounting for retainer | Yes |
| 1.16 | Declining/Terminating | If attorney withdraws, must give reasonable notice and protect your interests. | Abandoning case without notice at a critical time | Yes |
| 3.3 | Candor to Tribunal | Attorney must not make false statements to the court. | Presenting false evidence; misrepresenting facts to the judge | Yes |
| 3.4 | Fairness to Opposing Party | Attorney must not unlawfully obstruct access to evidence. | Hiding or destroying evidence; improperly influencing witnesses | Yes |
| 8.4 | Misconduct | Misconduct includes fraud, dishonesty, and conduct prejudicial to the administration of justice. | Lying to the court; misappropriating funds — most serious category | Yes — Most serious |

---

## Appendix B — ODC Contact Information

**Office of Disciplinary Counsel — Supreme Court of Pennsylvania**
Website: www.padisciplinaryboard.org

| Office | Address | Phone |
|--------|---------|-------|
| Philadelphia | 1601 Market St., Ste. 3320, Philadelphia PA 19103-2337 | (215) 560-6296 |
| Pittsburgh | 437 Grant Street, Frick Building, Ste. 1300, Pittsburgh PA 15219 | (412) 565-3173 |
| Harrisburg / Central | 820 Adams Ave., Ste. 170, Trooper PA 19403 | (610) 650-8210 |

**ODC Submission Requirements — Must Be Observed:**

- Submit complaint as ONE email with ONE attachment only
- Do not include zipped files
- Respond to all ODC inquiry letters within fourteen (14) days
- Required proof standard: clear and convincing evidence
- ODC does not represent complainants and cannot obtain damages on their behalf
- ODC cannot interfere with pending legal proceedings
- The disciplinary process is confidential until formal charges are filed by ODC

---

## Appendix C — Glossary of Terms

**23 Pa.C.S. § 5328:**
Pennsylvania Consolidated Statutes, Title 23 (Domestic Relations), Chapter 53, Section 28 — the governing statute for custody factor determinations in Pennsylvania.

**Act 11 of 2025:**
Pennsylvania legislation signed June 30, 2025, amending § 5328 with an effective date of sixty days from signing (approximately August 28, 2025).

**Catastrophic Priority:**
The highest priority level in PLCCMS — assigned to matters involving immediate child safety requiring urgent court intervention.

**Clear and Convincing Evidence:**
The evidentiary standard required by ODC to proceed on an attorney misconduct complaint — higher than preponderance of the evidence, lower than beyond reasonable doubt.

**Dual Party Mode:**
PLCCMS configuration in which both parties have submitted their Excel workbooks for merged comparison and unified judicial review.

**GAL:**
Guardian Ad Litem — a court-appointed attorney or advocate representing the child's best interests independently of both parties.

**GPL v3:**
GNU General Public License Version 3 — the open source license governing PLCCMS, ensuring it remains free public good software permanently and cannot be commercialized.

**ODC:**
Office of Disciplinary Counsel — the Pennsylvania Supreme Court body responsible for investigating attorney misconduct and enforcing the Rules of Professional Conduct.

**Pa. R.A.P. 124:**
Pennsylvania Rules of Appellate Procedure, Rule 124 — governing document format standards including 8.5 x 11 inch paper, 1 inch margins, 14pt font, and double spacing.

**PLCCMS:**
Parent/Lawyer Case Custody Management System — the full name of this application.

**Pro Se:**
Latin for "on one's own behalf" — a parent representing themselves without an attorney in court proceedings.

**PFA:**
Protection From Abuse order — a civil court order in Pennsylvania protecting a victim from an abuser.

**Single Party Mode:**
PLCCMS configuration in which only one parent has submitted data. Party B columns remain present and ready for future participation at any time.

**SVPO:**
Sexual Violence Protection Order — a civil court order in Pennsylvania protecting victims of sexual violence.

**Totality of Circumstances:**
The legal standard under § 5328(a.2) requiring courts to examine all custody factors together rather than relying on any single factor in awarding custody.

**Tyler Technologies Odyssey:**
A court case management system used by many Pennsylvania and national court systems, with which PLCCMS is designed to be compatible for institutional deployments.

**Webhook:**
A URL endpoint used by PLCCMS's optional tracking feature to send access event data to Google Sheets for security monitoring when the dashboard is shared online.

---

*Document prepared per Pennsylvania Rules of Appellate Procedure, Pa. R.A.P. 124*
*Parent/Lawyer Case Custody Management System (PLCCMS)*
*GNU General Public License v3 — Public Good Software*
*Pennsylvania Family Court | 23 Pa.C.S. § 5328 | Act 11, August 2025*
*Free Forever — For Every Parent — For Every Court — For Every Child*
