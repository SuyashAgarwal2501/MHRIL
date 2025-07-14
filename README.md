# MHRIL
Resort Performance Insights Agent – Knowledge Base
This repository contains reference documents used by a Microsoft Copilot agent designed to automate the generation of monthly performance insights for resort operations. The agent uses structured Excel data to populate a narrative report template and respond to interactive queries.

Repository Contents:
1. V4 Placeholder Formulae.docx
Defines the calculation logic for each placeholder used in the performance insights report. Includes formulas for occupancy, revenue, ARR, NRRPOR, REVPAR, and other KPIs. This document ensures consistent and accurate data transformation from the Excel input to the final report.

2. insight_template_v4.docx
A structured Word template used to generate the monthly resort performance insights report. Contains narrative sections for Occupancy and Revenue, with placeholders (e.g., {{Increase_Room_Nights}}, {{ARR_vs_LY_Trend}}) that are dynamically replaced by the agent using extracted data.

3. Populated_Insight_Report.docx
A sample filled version of the insight report template. Demonstrates correct placeholder replacement, formatting, and phrasing logic using real or sample data. Serves as a reference for expected output structure and tone.

How the Agent Uses These Files:
The Excel file containing resort-level performance data is uploaded by the user each month.
The Word template and formula document are stored in the agent’s knowledge base and referenced during processing.
The filled report example helps the agent understand expected output formatting and phrasing.

Output:
The agent generates
A filled PDF report with continuous sentence layout (no page breaks)
Optionally, a filled Word document if formatting is preserved

Monthly Flexibility:
Excel file changes monthly and is uploaded by the user
Template and formula documents remain static and are referenced from this repository
