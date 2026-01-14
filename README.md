\# 🏰 Project Fortress: Zero-Trust Association Management System



\*\*A Serverless, Event-Driven Architecture for managing 1,100+ Real Estate Allottees.\*\*



\## 🚀 Project Overview

This project is a comprehensive backend system built on \*\*Google Apps Script\*\* (JavaScript) to automate the operations of a large Allottees Association. It replaces manual spreadsheet work with a \*\*Zero-Trust\*\* automated workflow that handles member onboarding, financial compliance, document security, and communication.



\## ⚡ Key Features (The 8 Engines)



The system is modularized into 8 distinct "Engines" for reliability:



\* \*\*1. The Gatekeeper (Security):\*\* Validates every Google Form submission against a master ledger. Prevents unauthorized access and blocks identity spoofing.

\* \*\*2. The Auditor (Compliance):\*\* A scheduled "Crawler" that performs deep API scans of Google Drive permissions. It detects "strangers" (unauthorized emails) and manages file/folder integrity.

\* \*\*3. The Insolvency Scanner (Legal):\*\* Dynamically cross-references member data against real-time NCLT (National Company Law Tribunal) watchlists.

\* \*\*4. Finance \& UPI Engine:\*\* Auto-generates UPI payment links with embedded Transaction Notes (`tn`) for accurate tracking.

\* \*\*5. The Template System:\*\* An attachment-ready email engine that merges data into HTML templates.

\* \*\*6. Receipt Generator:\*\* Auto-generates PDF receipts, archives them in a secure Vault, and emails them—all in one trigger.

\* \*\*7. The Welcome Dispatcher:\*\* Handles new member onboarding with link validation and audit logging.

\* \*\*8. The Time Machine (Backup):\*\* A self-healing backup system with auto-rotation policies (retains the last 5 snapshots).



\## 🛠️ Tech Stack

\* \*\*Language:\*\* JavaScript (Google Apps Script / ES6)

\* \*\*Database:\*\* Google Sheets (as a relational database)

\* \*\*Storage:\*\* Google Drive API (Advanced Permissions Management)

\* \*\*Triggers:\*\* Time-driven (Cron jobs) and Event-driven (Form Submit)



\## 🔒 Security Architecture

This project implements a \*\*Zero-Trust\*\* model:

1\.  \*\*Identity Verification:\*\* No user is trusted by default; email and flat number must match the master record.

2\.  \*\*Least Privilege:\*\* The system actively removes unauthorized viewers from Drive folders during the audit cycle.

3\.  \*\*Audit Trails:\*\* Every file movement, permission change, and email dispatch is logged immutably.



\## 📊 Workflow Diagram

```mermaid

graph TD

&nbsp;   A\[User Submits Form] -->|Trigger| B(Engine 1: Gatekeeper)

&nbsp;   B -->|Check Identity| C{Valid?}

&nbsp;   C -->|No| D\[Block \& Log Mismatch]

&nbsp;   C -->|Yes| E\[Move Files to Vault]

&nbsp;   E --> F\[Update Master Ledger]

&nbsp;   

&nbsp;   G\[Time Trigger] -->|Every Hour| H(Engine 2: Auditor)

&nbsp;   H -->|Scan Permissions| I{Strangers Found?}

&nbsp;   I -->|Yes| J\[Remove Stranger \& Alert Admin]

&nbsp;   I -->|No| K\[Verify File Integrity]

