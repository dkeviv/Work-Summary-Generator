Executive Work Summary Generator 🚀
A Google Workspace Add-on that provides a powerful, interactive interface directly within Gmail to automatically analyze your Google Calendar and Gmail activity. It generates a professional, data-driven work summary, complete with executive-level insights and metrics, ready for performance reviews.

(Image: Example of a Google Workspace Add-on icon in the sidebar)

✨ Features
Interactive UI in Gmail: No more running scripts from an editor. A clean, self-explanatory interface lives in your Gmail sidebar.

Intelligent Analysis:

Automated Categorization: Classifies your work into key areas: Strategic Planning, Team Leadership, Operational Execution, and Client/Stakeholder Facing.

Relevance Scoring: Scores each activity based on keywords and collaboration to measure impact, not just volume.

Strategic Focus Score: Compares your actual work against ideal role priorities to identify alignment or gaps.

Topical Deep Dive:

Summarize all email conversations related to a specific project or keyword.

Optionally filter the topic summary to conversations with a specific person.

User-Friendly Inputs:

A welcome screen and guided form.

Easy calendar date pickers for selecting the analysis range.

Role-based presets (Executive, Manager, etc.) to auto-populate strategic weights.

VP-Ready Output:

Generates a polished, professionally formatted Google Doc.

Includes an Executive Summary, performance metric tables, activity distribution charts, and data-driven recommendations.

🛠️ Technology Stack
Google Apps Script: The core language for the logic and backend.

Gmail Add-on Framework: Using CardService to build the native user interface inside Gmail.

Google Workspace APIs: GmailApp, CalendarApp, and DocumentApp for data retrieval and report generation.

⚙️ Setup and Installation
Follow these steps to deploy the add-on for your own Google account.

1. Create the Apps Script Project
Go to the Google Apps Script dashboard.

Click New project.

Give the project a name, for example, "Work Summary Generator".

2. Add the Code
You will need to populate two files in your project.

A. The Manifest File (appsscript.json)

In the editor, click the Project Settings (⚙️) icon on the left.

Check the box for "Show appsscript.json manifest file in editor".

Return to the Editor (⛽) and click on the appsscript.json file.

Delete all content in the file and paste the code from this file.

B. The Script File (Code.gs)

Click on the Code.gs file.

Delete all content in the file and paste the code from this file.

3. Save and Deploy
Click the Save project (💾) icon.

Click the Deploy button and select Test deployments.

A dialog will open. Click Install, then Done.

You will be prompted to grant permissions. Review and Allow them. This is necessary for the script to read your calendar/email and create a doc.

🚀 How to Use
Refresh Gmail: After installing the test deployment, do a full refresh of your Gmail tab (Ctrl+R or Cmd+R).

Open the Add-on: Look for the Work Summary Generator icon in the right-hand sidebar and click it.

Start: On the welcome screen, click Create New Report.

Configure: Fill out the form:

Full Name: Your name for the report title.

Start/End Date: Use the calendar pickers to select the analysis period.

Topical Deep Dive (Optional): Enter a project name or keyword to get a special summary. You can also add a person's email to filter that summary.

Strategy Weights: Select a role preset (like "Manager") to automatically fill the weights, and adjust if needed.

Generate: Click the Generate Report button.

Done! A success card will appear with a button to open your newly generated Google Doc report.

📄 Report Structure
The generated Google Doc is professionally formatted with the following sections:

Topical Deep Dive (if a topic was provided)

Executive Summary

Performance Metrics

Activity Distribution

Detailed Activities

Work Patterns

Executive Recommendations

🔧 Customization
You can easily customize the analysis logic by editing the CONFIG constant at the top of Code.gs. The most common customization is to change the KEYWORDS to better match your organization's terminology.

📄 License
This project is licensed under the MIT License. See the  file for details.






