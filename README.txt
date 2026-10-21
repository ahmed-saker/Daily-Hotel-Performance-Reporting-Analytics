# 🏨 Daily Hotel Performance Reporting & Analytics

An automated Python reporting solution designed to analyze hotel performance, track daily production, compare results with the previous week, and distribute detailed Excel reports and visual insights via Microsoft Outlook.

## 📌 Project Overview

This project automates hotel performance reporting by integrating hotel portfolio data with sales and production metrics from an OLAP data cube. It generates individual performance reports for staff members and a consolidated summary report for management.

The solution helps monitor hotel activation, identify newly producing and dropped hotels, evaluate profitability, and support data-driven business decisions.

## ✨ Key Features

- **Automated Data Extraction:** Retrieves hotel performance metrics from an OLAP cube using MDX queries.
- **Daily Performance Monitoring:** Analyzes Bookings, Room Nights, Sales, Profit, and Profit Margin.
- **Week-over-Week Analysis:** Compares yesterday's performance with the same day of the previous week.
- **Hotel Activation Tracking:** Identifies newly producing hotels, dropped hotels, consistently producing hotels, and hotels without production on either comparison date.
- **Top 10 Hotel Analysis:** Generates charts ranking hotels by Bookings, Room Nights, Sales, and Profit.
- **Excel Report Automation:** Creates formatted Excel reports with conditional formatting, performance indicators, and summary sheets.
- **Automated Email Distribution:** Sends personalized reports to staff and consolidated summaries to management through Microsoft Outlook.
- **Scheduled Execution:** Supports daily automated execution at 9:00 AM.
- **Logging and Error Handling:** Records execution progress, data retrieval issues, and email delivery errors.

## 📊 Key Performance Indicators

| KPI | Description |
|---|---|
| Bookings | Number of bookings generated |
| Room Nights | Total room nights produced |
| Sales | Total sales generated |
| Profit | Total profit generated |
| Profit Margin (%) | Profit as a percentage of sales |
| Producing Hotels | Hotels generating bookings |
| New Hotels | Hotels producing yesterday but not on the comparison date |
| Dropped Hotels | Hotels producing on the comparison date but not yesterday |
| WoW Difference | Change in bookings, room nights, sales, and profit |

## 🛠️ Technology Stack

- **Python** — Automation and data processing
- **Pandas & NumPy** — Data manipulation and numerical analysis
- **Pyadomd & MDX** — OLAP cube connectivity and querying
- **OpenPyXL** — Excel report generation and formatting
- **Matplotlib** — Performance charts and visualizations
- **PyWin32** — Microsoft Outlook automation
- **Logging** — Execution monitoring and error tracking

## ⚙️ Workflow

1. Load the latest available Active Sheet containing hotel assignments.
2. Read staff information and email addresses from the configuration workbook.
3. Retrieve hotel performance data for yesterday and the same day of the previous week.
4. Clean hotel codes and merge the retrieved data with the hotel portfolio.
5. Calculate profit margins and week-over-week differences.
6. Identify new, dropped, active, and inactive hotels.
7. Generate individual Excel reports and Top 10 performance charts for each staff member.
8. Create a consolidated management summary.
9. Distribute reports through Microsoft Outlook.
10. Log execution status and processing errors.

## 📁 Expected Input Files

The automation expects the following business files to be available in the configured directories:

- **Active Sheet:** Contains hotel codes, accommodation names, staff assignments, and related hotel information.
- **ContractingEmails.xlsx:** Contains staff names and email addresses.

The exact file paths and column names should match the configuration and data structure used by the script.

## 🚀 Installation & Setup

### 1. Clone the repository

```bash
git clone https://github.com/YOUR_USERNAME/YOUR_REPOSITORY.git
cd YOUR_REPOSITORY
```

### 2. Create a virtual environment

```bash
python -m venv .venv
.venv\Scripts\activate
```

### 3. Install dependencies

```bash
pip install pandas numpy openpyxl matplotlib pywin32 pyadomd
```

### 4. Configure the environment

Before running the project, make sure you have:

- Windows with Microsoft Excel and Outlook available as required by the automation.
- Access to the relevant OLAP cube and its data source.
- The required Microsoft ADOMD.NET provider installed.
- Valid input Excel files and correctly configured file paths.
- Authorized Outlook account access for sending reports.
- Appropriate database or cube credentials configured securely.

**Security:** Credentials must be stored outside the source code, such as in environment variables or a secure secrets manager. Never commit passwords, private connection strings, internal email lists, or confidential business data to a public repository.

## ▶️ Running the Project

Run the scheduled reporting process:

```bash
python DailyTop10HotelsByProduction.py
```

The script is configured to check for its scheduled daily execution time of **9:00 AM**.

For testing, the main reporting function can be invoked directly after configuring the required dependencies and business environment.

## 📈 Business Value

This solution helps businesses:

- Reduce manual reporting and repetitive Excel tasks.
- Improve visibility into hotel production and profitability.
- Monitor hotel activation across staff portfolios.
- Detect changes in hotel production from week to week.
- Identify opportunities to improve hotel performance.
- Provide management with consistent, timely performance insights.

## 🔒 Data Privacy

This project is designed for authorized business environments. Connection details, employee information, internal reports, and production data should remain confidential. Use anonymized sample data when sharing demonstrations or publishing project examples.

## 👨‍💻 Author

**Ahmed Saker**

Data Analyst | Python Automation | SQL | Power BI | Excel | Data Visualization

Focused on building automated reporting solutions and transforming operational data into actionable business insights.
