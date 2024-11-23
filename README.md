# VBA---Filtering-and-Comparing-
1. Automated Filtering and File Comparison Tasks:
Today's Event Details File Comparison:

User provided an Event Details file and a Verified List.
The task involved filtering the Event Details file based on specific column values (Column AT = Yes, Column V = None, and Column Y = 2024).
The code compared Column A in the Event Details file to the Verified List, identifying any new records and copying them into a new file, named "New FC PO Issued, Scope None - Today's date".

EV NEX Discrepancy File Comparison:

User provided the "Latest EV NEX Discrepancy" file and the "Previous Week's EV NEX Discrepancy" file.
For both files, filters were applied based on specific criteria (Column G = MENA, Column AA > 0, and Column AQ = False).
The code compared both files for new records in two key scenarios:
Probability: Comparing Column AQ in both files. If new records were found in the latest file, those rows were copied to a new sheet named Probability.
Ownership: Comparing Column BJ in both files. New records were copied into a new sheet named Ownership.
If no new records were found, appropriate messages ("No new Probability" or "No new Ownership") were displayed.

2. Outcome Handling:
File Creation: In cases where new records were found (Probability or Ownership), a new file was generated with the relevant records, named EV NEX Defects.xlsx.
Conditional File Saving: If no new records were found, the output file was either not created or closed without saving.
User Notifications: The user was notified with messages like "Probability ready", "Ownership ready", "No new Probability", or "No new Ownership" based on the outcome.

3. Workflow Automation:
User Prompts: The VBA code prompted the user to upload the necessary files and handled file and sheet management based on their inputs.
Data Filtering: It filtered the files based on criteria and then performed comparisons between the two sets of data.
Efficient Record Handling: The new records were copied to a new file or sheet, maintaining the header and the required data.
These tasks streamline the process of comparing EV NEX discrepancy files and extracting new relevant records, saving significant manual work in the process.
