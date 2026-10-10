# 🏫 Centralized Attendance Automation System

An enterprise-grade attendance tracking system built using **Excel VBA (Frontend)** and **SQL Server (Backend)** to automate daily attendance consolidation across 50+ CBSE School branches.

---

## 🎯 Problem Statement
Managing daily attendance data manually across 50+ CBSE school branches involves collecting individual Excel files, consolidating records, checking standard vs present counts, and generating formatted PDF reports. This manual workflow consumed over 1 hour daily and was highly prone to human errors and data inconsistencies.

---

## 💻 Prerequisites & System Requirements
Before running the branch-side application on any branch PC, ensure the following driver is installed:
- **ODBC Driver 18 for SQL Server:** Must be installed on each branch computer to establish a secure, encrypted connection with the central SQL Server database.

---

## 🏗️ System Architecture & Workflow
1. **Branch Level:** Each branch logs in securely using encrypted credentials. Branch users fill a dynamic userform (Classes 1 to 10 standard & present counts), which validates data and pushes it directly to the Central SQL Server via ADO commands.
2. **Central Admin Level:** Administrators can fetch real-time consolidated attendance reports, auto-format data tables, add serial numbers, and export reports instantly into formatted landscape PDF formats saved directly to the desktop.

---

## 🗄️ Database Schema
The backend uses MS SQL Server with optimized tables and constraints:
- `Branchlogin`: Manages branch-wise secure credentials and routing.
- `Attendance`: Stores daily entry counts with a composite unique constraint (`BranchName, EntryDate`) to prevent duplicate submissions.

*(You can check the full SQL script in `database/schema.sql`)*

---

## 🛠️ Tech Stack
- **Frontend / UI:** Microsoft Excel VBA (UserForms, ADODB, FileSystemObject)
- **Backend / Storage:** Microsoft SQL Server (ODBC Driver 18)
- **Reporting:** Automated Excel-to-PDF Export Engine