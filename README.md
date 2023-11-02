# Data Health Check

In this section, we will provide documentation for the code related to the Data Health Check process. The process involves using Google Sheets, Apps Script, Metabase, and the integration provided by the Metabase-Google Sheets add-on found in the 'Health_Check' directory.

## Data Retrieval and Update

To automate data retrieval and updates, we use triggers and the following steps:

1. **Data Retrieval**: We pull the data from Metabase and update it every two hours. This is done using the integration provided by the Metabase-Google Sheets add-on.

2. **Backup Data**: In a separate sheet, we generate tabs D1 to D(end of month) and create a query to pull data from yesterday and update the integration to extract the data from that QUESTION_NUMBER into yesterday's tab. This ensures that we lock in the data after 24 hours.

## Data Comparison and Verification

The core of the Data Health Check process involves comparing the data retrieved from the Metabase every two hours with the locked data from 24 hours ago. This is achieved through the following steps:

3. **Data Comparison**: We apply a Google Query function to append all the data from each day's tab into one consolidated location for easy comparison.

4. **Conditional Formatting**: To highlight discrepancies or changes in the data, we apply conditional formatting.

5. **Data Copying**: The compared data is copied into a third tab in the 'Health Check' sheet for further analysis.

For the actual JavaScript code to implement these steps, please refer to the code snippets below:
