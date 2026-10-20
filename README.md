# IT-Ticket-Data-Analysis-using-Excel
IT Ticket Data Analysis – IT Service Desk Performance & Customer Satisfaction
Project Overview
This project analyzes IT service desk ticket data to evaluate support-team performance, ticket demand, resolution efficiency, and employee satisfaction. The analysis uses Microsoft Excel to clean and transform the data, create Pivot Tables, calculate performance metrics, and build an interactive service desk dashboard.

The project covers 97,498 IT tickets handled by 50 IT agents from 2016 to 2020. The findings are used to identify operational bottlenecks, compare agent performance, and recommend improvements to ticket management processes.
Business Problem
The IT support team manages a large and growing volume of tickets. Differences in agent performance and delays in resolving certain categories of requests can affect employee satisfaction and service quality.

The project aims to understand ticket patterns, identify underperforming agents and categories with longer resolution times, and determine whether investment should be prioritized toward ticket management software, agent training, or additional staffing.
Project Objectives
- Clean and prepare ticket and IT agent datasets for analysis.
- Examine ticket volume and trends over time.
- Analyze ticket distribution by issue type, request category, severity, and priority.
- Measure average resolution time and satisfaction ratings.
- Evaluate ticket-handling volume and performance across IT agents.
- Identify agents whose satisfaction and resolution-time metrics fall below team benchmarks.
- Explore quarterly performance and employee demographic information.
- Build a dashboard to communicate operational KPIs and findings.
- Recommend actions to improve service efficiency and employee satisfaction.
Dataset Overview
The Excel workbook contains ticket-level records, agent details, cleaned data, analysis worksheets, Pivot Tables, and a dashboard.
Dataset / Sheet	Description
Tickets	Ticket records, dates, employee and agent IDs, request category, issue type, severity, priority, resolution time, and satisfaction rating
IT Agents	Agent IDs, names, email addresses, and date-of-birth information
Cleaned Data	Prepared analysis dataset with additional fields such as numeric severity, month/year, day of week, employee age, year, and quarter
Objective & Subjective Question Sheets	Supporting analysis and responses to project questions
Pivot_table_dashboard	Aggregated metrics used to support dashboard reporting
Dashboard	IT service desk performance dashboard


Dataset scale: 97,498 ticket records, 50 IT agents, 16 reported data attributes, and a reporting period from 2016 through 2020.
Tools & Technologies
- Microsoft Excel – Data cleaning, transformation, calculations, Pivot Tables, and dashboard development
- Excel formulas – Data extraction, lookup, date calculations, and statistical analysis
- Pivot Tables – KPI summaries, category analysis, agent comparisons, and time-based analysis
- Data Visualization – Dashboard charts and performance reporting
- Microsoft PowerPoint – Presentation of findings and recommendations
Data Cleaning & Preparation
The project includes the following preparation steps:
- Reviewed the ticket and IT agent worksheets to understand the available fields.
- Corrected an inconsistent severity label by replacing 3-Mayor with 3-Major in 4,836 records.
- Used XLOOKUP to retrieve IT agent names from the agent master sheet using Agent ID.
- Extracted email domains using Excel text functions.
- Created date-of-birth values using DATE() and calculated agent ages using DATEDIF().
- Created additional analysis fields for numeric severity, month and year, day of week, employee age, year, and quarter.
- Used Pivot Tables to summarize ticket counts, resolution times, satisfaction ratings, and agent performance.
Excel Functions & Analysis Techniques
The project uses the following Excel features and functions:
- Pivot Tables – Ticket distribution, agent workload, request-category analysis, and quarterly trends
- XLOOKUP – Matching Agent IDs with agent names
- RIGHT, LEN, FIND – Extracting email domains
- DATE, DATEDIF, TODAY – Constructing dates and calculating age
- AVERAGE – Calculating average resolution time, satisfaction, and agent age
- CORREL – Measuring the relationship between numeric severity and resolution time
- LEFT, FIND, VALUE – Extracting numeric severity values
- Find & Replace – Correcting inconsistent category labels
- Date grouping and calculated fields – Monthly, yearly, quarterly, and weekday analysis
Key Performance Indicators
The project presentation reports the following overall metrics:
KPI	Result
Total IT tickets	97,498
Total IT agents	50
Reporting period	2016–2020
Average satisfaction rate	4.10 / 5.0
Average resolution time	4.55 days
Average daily ticket volume	53.36
Average age of IT agents	40.76 years


Key Analysis & Findings
1. Ticket Distribution by Issue Type
Issue Type	Ticket Count	Share
IT Request	73,220	75%
IT Error	24,278	25%


IT Requests account for approximately three-quarters of all tickets, making them the largest source of service desk demand.
2. Ticket Distribution by Request Category
Request Category	Ticket Count
System	39,002
Login Access	29,193
Software	19,570
Hardware	9,733


System and Login Access together account for approximately 70% of the total ticket volume.
3. Satisfaction Analysis
- Satisfaction rating 5 accounts for 50,770 tickets, approximately 52% of the total.
- Ratings 4 and 5 together represent approximately 80% of tickets, according to the project presentation.
- Rating 1 accounts for 9,907 tickets and represents an important group for service-quality investigation.
4. Ticket Volume Trends (2016–2020)
- Ticket volume increased from 13,051 in 2016 to 29,088 in 2020.
- The presentation reports an increase of approximately 123% over the five-year period.
- The rising workload indicates a need to improve ticket-handling processes and support capacity.
5. IT Agent Performance
The presentation identifies 15 out of 50 agents as below both performance benchmarks:
Metric	Team Average	Identified Underperforming Agents
Satisfaction rate	4.10	3.71
Resolution time	4.55 days	5.02 days


These agents have lower average satisfaction ratings and longer average resolution times than the overall team benchmarks.
6. Resolution Time by Request Category
The project documentation identifies the following categories as having average resolution times above the overall average of approximately 4.55 days:
Request Category	Average Resolution Time
Hardware	7.625 days
System	6.615 days
Software	5.238 days


These categories are potential areas for process review and improvements to the ticket management workflow.
7. Severity and Resolution Time
The project calculates a correlation of approximately -0.0405 between numeric severity and resolution time. This indicates a very weak negative linear relationship in the analyzed dataset.
Dashboard Overview
The Excel workbook includes an IT Service Desk Performance Dashboard designed to summarize ticket demand, resolution, and employee satisfaction.

The dashboard and supporting Pivot Tables cover:
- Overall ticket and service performance KPIs
- Ticket distribution by request category and priority
- Ticket volume by year and quarter
- Satisfaction-rate analysis
- Agent-level ticket workload and performance
- Quarterly satisfaction and category comparisons
Business Recommendations
Based on the analysis and the project presentation, the recommended priorities are:
1. Upgrade Ticket Management Software
- Prioritize improvements to ticket management software and supporting infrastructure.
- Review workflows for categories with longer resolution times.
- Explore automation to reduce manual work and improve resolution efficiency.
2. Train Underperforming IT Agents
- Provide structured training to the 15 agents identified below both performance benchmarks.
- Focus on improving resolution efficiency and employee satisfaction.
- Monitor agent-level performance after training to evaluate progress.
3. Evaluate Staffing Requirements
- Monitor ticket volume and workload as demand continues to grow.
- Consider hiring additional IT agents if software improvements and training do not sufficiently improve service performance.
- Use performance reviews and workload metrics to support staffing decisions.
Project Deliverables
- Excel workbook containing source data, cleaned data, calculations, analysis worksheets, Pivot Tables, and the dashboard.
- Project documentation describing the objective and subjective questions, analytical approach, findings, and recommendations.
- PowerPoint presentation summarizing the IT support performance story and proposed actions.
Repository Structure
IT-Ticket-Data-Analysis/
│
├── Doc Template (IT Tickets).docx
├── IT Ticket Analysis presentation.pptx
├── IT Tickets Analysis - Professional Dashboard.xlsx
└── README.md
The filenames shown above reflect the supplied project folder. Update them if you rename files in your GitHub repository.
How to Explore the Project
1. Download or clone the repository.
2. Open the Excel workbook in Microsoft Excel.
3. Review the Tickets, IT Agents, and Cleaned Data worksheets.
4. Explore the objective and subjective analysis worksheets for the questions and calculations.
5. Review the Pivot Tables and Dashboard sheet for summarized performance metrics and trends.
6. Refer to the project documentation and PowerPoint presentation for interpretation of the findings and recommendations.
Project Outcome
This project demonstrates how Excel can be used to clean and analyze operational data, measure IT service desk performance, identify workload and resolution bottlenecks, and communicate findings through dashboards and presentations. It combines spreadsheet-based data preparation, analytical formulas, Pivot Tables, KPI reporting, and business recommendations to support data-driven IT service management decisions.
