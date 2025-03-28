# VBA---Filtering-and-Comparing-

**EV NEX**

*The VBA code is a macro designed to compare two Excel files containing EV NEX discrepancy data. It filters records based on specific criteria and outputs discrepancies into a new workbook. Here's a summary of its functionality:*

 

### Function Overview:

1. **File Selection**: Prompts the user to open the latest and previous EV NEX discrepancy files.

2. **Data Filtering**: Applies filters to the worksheets to select records based on specific criteria:

   - Column G ("MENA")

   - Column AA (greater than 0)

   - Column AQ ("False")

3. **Output Workbook Creation**: Creates a new workbook to store discrepancies found.

4. **Comparison Logic**:

   - Compares rows in the latest worksheet to those in the previous worksheet.

   - If a row from the latest worksheet does not have a matching entry in the previous worksheet based on Column A, it is considered a discrepancy.

   - The code handles "Probability" and "Ownership" discrepancies separately.

5. **Output and Notification**: Copies headers and relevant rows into corresponding sheets ("Probability" and "Ownership") in the new workbook and notifies the user of the results.

6. **File Saving**: Saves the new workbook if discrepancies are found.
















**FC PO**

The provided VBA code is a macro designed for Microsoft Excel that performs the following tasks:

 

1. **Prompt for File Selection**: It prompts the user to open two Excel files: "Today's Event Details" and "Verified List". These files are assumed to have data in the first worksheet.

 

2. **Filter and Compare**: It filters rows in the "Event Details" file based on specific conditions:

   - Column AT must contain "Yes".

   - Column V must contain "None".

   - The year in Column Y must be 2024.

  

   It then checks if the value in Column A of these filtered rows exists in the "Verified List".

 

3. **Generate a New File**: If a row from "Event Details" does not have its Column A value in the "Verified List", it copies the entire row to a new worksheet in a newly created workbook. The header from "Event Details" is copied to the new worksheet.

 

4. **Save the New File**: The new workbook is saved in the default file path with a name that includes today's date.

 

5. **Notification and Cleanup**: It notifies the user once the new file is created and optionally closes the original workbooks without saving changes.

 

 

 

 
