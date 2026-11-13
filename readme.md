# 📊 BPO Operations, Workforce Management (WFM) & Quality Assurance Analytics Suite

## 📌 Project Overview
This project consists of **4 independent Microsoft Excel / VBA files** built around real-world **BPO (Business Process Outsourcing) & Shared Services** operational datasets. 

Together, these workbooks demonstrate end-to-end proficiency in **BPO Operations Analytics, Workforce Management (WFM), Quality Assurance (QA) Auditing, Process Automation (VBA Macros), and Executive BI Dashboarding**. Each file operates as a standalone solution addressing a critical stage of the BPO analytics lifecycle—from automated raw data ingestion and formula-driven SLA tracking to relational data modeling (`Power Pivot` / `DAX`) and interactive executive reporting.

---

## 📂 The 4 Standalone Project Files

```text
📁 Local_BPO_Analytics_Project/
│
├── 📄 Raw Data for Import.xlsx   # [File 1] BPO Production & Audit Staging Dataset (1,000+ / 30,000+ rows)
├── 📄 Task 1.xlsx                # [File 2] BPO SLA, AHT, Utilization & Productivity Analysis Workbook
├── 📄 Task 2.xlsm                # [File 3] VBA-Automated Data Import & Workflow Processing Engine
└── 📄 Task 3.xlsm                # [File 4] BPO Executive BI Dashboard, Power Pivot Model & QA Analytics
```

### 1️⃣ `Raw Data for Import.xlsx` — *BPO Production & Audit Staging Dataset*
* **Purpose:** Serves as the raw transactional data extract typical of BPO workflow management systems (e.g., ServiceNow, Salesforce, or internal ticketing tools).
* **BPO Data Attributes:**
  * **Client & Process Hierarchy:** `Account` (Client), `Category`, `Task Name`, `Case Status` (Completed, Pending, Pended).
  * **Time & Volume Tracking:** `Assign Date`, `Complete Date`, `# Received`, `# Processed`, `Duration Tracked` (Total Handle Time), `Duration Chargeable` (Billable Client Time).
  * **Workforce & Training Segmentation:** `Processor` (Agent), `IsTrainningCase` (Nesting vs. BAU), `Trainning Mode Type` (Buddy, Classroom, Live).

### 2️⃣ `Task 1.xlsx` — *BPO SLA, AHT & Agent Productivity Analysis*
* **Purpose:** A standalone analytical workbook focused on formula-driven **Descriptive & Diagnostic Analysis** of agent-level and account-level BPO metrics.
* **Key Skills & Methodologies Showcased:**
  * **Turnaround Time (TAT) & SLA Adherence:** Evaluating `Assign Date` vs. `Complete Date` to measure processing speed and backlog aging.
  * **Capacity & Throughput Analysis:** Comparing `# Received` vs. `# Processed` to calculate **Clearance Rates** and identify process bottlenecks.
  * **Billable Utilization & Shrinkage:** Analyzing `Duration Chargeable` against `Duration Tracked` to measure productive efficiency and non-billable overhead.
  * **Excel Competencies:** Advanced Lookup Functions (`XLOOKUP`/`INDEX-MATCH`), Multi-condition Aggregations (`SUMIFS`, `COUNTIFS`, `AVERAGEIFS`), Logical Arrays, and Conditional Formatting for outlier detection.

### 3️⃣ `Task 2.xlsm` — *VBA-Automated ETL & Reporting Engine*
* **Purpose:** A standalone Macro-Enabled Workbook (`.xlsm`) engineered to automate repetitive BPO MIS (Management Information System) tasks and eliminate manual data-entry errors.
* **Key Skills & Methodologies Showcased:**
  * **VBA Macro Automation (`vbaProject.bin`):** Custom Visual Basic for Applications (VBA) scripts triggered via interactive UI controls to automate file ingestion, data cleansing, and report formatting.
  * **Automated ETL Workflow:** Programmatically imports external staging data (`Raw Data for Import.xlsx`), standardizes data types, and populates structured reporting templates in seconds.
  * **Operational Impact:** Demonstrates **Robotic Process Automation (RPA) / Macro scripting** capabilities that reduce daily MIS reporting turnaround time by over **80%**.

### 4️⃣ `Task 3.xlsm` — *Executive BPO BI Dashboard, Power Pivot Data Model & QA Analytics*
* **Purpose:** A comprehensive, macro-enabled **Business Intelligence (BI) application** combining **30,407 production records** (`Table1`) with **2,470 Quality Assurance (QA) audit logs** (`Table2`).
* **Key Skills & Methodologies Showcased:**
  * **Relational Data Modeling (`Power Pivot`):** Utilizes Excel's in-memory xVelocity engine to link Production Logs (`Processor`, `Task Name`, `Account`) with Quality Audit Logs (`Auditor`, `Audited`, `Errors`) without `VLOOKUP` performance lag.
  * **DAX (Data Analysis Expressions):** Custom calculated measures including `DISTINCTCOUNT` of cases, **Defect Density / Error Rate %**, **Audit Sample Coverage %**, and **First Time Right (FTR) / Accuracy %**.
  * **Six Sigma & Quality Assurance (QA) Analytics:** Cross-analyzes agent throughput against audit defect counts (`Errors`) to balance speed (AHT) vs. quality (Accuracy).
  * **Nesting & Onboarding Cohort Analysis:** Evaluates learning curves by comparing error rates and handle times between tenured BAU agents and trainees (`IsTrainningCase`, `Trainning Mode Type`).
  * **Interactive Executive Dashboard:** Features **6 interconnected Pivot Tables**, **5+ dynamic charts** (Combo, Pareto, Trend, and Distribution visuals), and synchronized **Slicers** for self-service filtering by Client Account, Task Category, Processor, and Auditor.

---

## 🎯 Core BPO & Analytics Keywords (ATS & Recruiter Optimized)

| Domain | Skills, Methodologies & Keywords |
| :--- | :--- |
| **BPO & WFM Domain** | Service Level Agreement (SLA), Average Handle Time (AHT), Turnaround Time (TAT), Full-Time Equivalent (FTE) Utilization, Billable vs. Non-Billable Hours, Backlog Aging, Throughput & Clearance Rate, Nesting / Training Ramp-up Analysis, MIS Reporting |
| **Quality Assurance (QA)** | Quality Audit Compliance, Defect Density, Error Rate %, First Time Right (FTR), Auditor Calibration, Pareto Analysis (80/20 Rule), Root Cause Analysis (RCA), Six Sigma Quality Control |
| **Advanced Excel & BI** | Power Query (ETL), Power Pivot, Relational Data Modeling, Star Schema, DAX Measures (`DISTINCTCOUNT`), Multi-Dimensional Pivot Tables, Slicers & Timelines, Dynamic Arrays, Conditional Formatting |
| **Automation & Programming** | Visual Basic for Applications (VBA), Macro-Enabled Workbooks (`.xlsm`), Automated File Import, UI Form Controls, Event-Driven Procedures, MIS Workflow Automation |

---

## 💡 Key Analytical Methodologies Demonstrated

1. **Efficiency vs. Quality Trade-off Matrix:** Correlating agent speed (`# Processed` & `Duration Tracked`) with quality audit outcomes (`Audited` & `Errors`) to identify high-performers vs. agents requiring targeted coaching.
2. **Client Profitability & Chargeability Analysis:** Measuring the ratio of `Duration Chargeable` to `Duration Tracked` across different client `Accounts` and `Categories` to uncover revenue leakage.
3. **Training Effectiveness (Cohort Variance):** Comparing defect rates and processing durations across `Trainning Mode Types` to optimize BPO onboarding curriculums.
4. **End-to-End MIS Automation:** Demonstrating how raw operational dumps (`Raw Data for Import.xlsx`) are systematically transformed through standalone formula templates (`Task 1.xlsx`), VBA scripts (`Task 2.xlsm`), and relational BI dashboards (`Task 3.xlsm`).

---
## Inside Visuals for the Project;

<img width="2877" height="1688" alt="image" src="https://github.com/user-attachments/assets/30f08b61-edfa-43ca-8031-4cd4103ba6f8" />
<img width="1636" height="1037" alt="image" src="https://github.com/user-attachments/assets/3bc9813a-b4a1-4e16-ba9c-03ceb877f678" />
<img width="1632" height="645" alt="image" src="https://github.com/user-attachments/assets/86bbf84e-0ea5-430e-a1d1-e8c0070d8f4a" />
<img width="2817" height="897" alt="image" src="https://github.com/user-attachments/assets/b072fd6c-d54d-4ec1-9686-69cb0766f764" />
<img width="2854" height="1326" alt="image" src="https://github.com/user-attachments/assets/af67bc2b-a676-452a-807e-5d810e60b096" />
<img width="2864" height="1342" alt="image" src="https://github.com/user-attachments/assets/1fc78b4b-644e-4ccd-90ec-a145f1afd238" />





