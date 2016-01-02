# ALM--JIRA-Public-Repo
Import_Excel_2007_Macro_Scripts

ALM -> JIRA Imports
1.	Export latest data from ALM defects as .xlsx file. Make sure to select the following 15 column labels:

2.	Change defect type to “Bug”

3.	Change “Detected by” column to “********” and hide the column.

4.	Also, hide “defect type”, “Browser”, “Description”, “Environment” and “Summary”.

5.	Perform the following VLookup function in first cell within column “JIRA ID”:

i.	For “Lookup_Value” choose corresponding “Defect ID”.

ii.	For “Table_Array” choose the first two columns (i.e. “Defect ID” & “JIRA ID”) in the old master file.

iii.	For “Col_Index_Num” type in “2”.

iv.	For “[range_lookup]” type “false”.

v.	Now in your new master file, select all cells under “JIRA ID” column and the copy and paste (as values).

6.	Create a new workingfile (in .xlsx format) from the newly created master file (in step no. 1 above).

7.	In old master file – All we need to do is to check the time at which the previous import was done.

a.	In “Modified”: By filtering, check and record the last time the import was performed.

b.	Delete all filtered rows.

8.	For the“Status” column:
a.	Filter out: Closed, Dev Assigned, Fixed, Dev Rework and New (if found). Then delete rest of the data.

b.	Filter for each of the “Closed” and “Fixed” and then perform the following:

i.	For Closed:
	Delete all rows containing “#N/As” and clear “Comments” in the remaining rows.
	Change all values under the column “Assigned to” to “saiqa.chaudry”.
ii.	For Fixed, 

	Clear “#N/As” from “JIRA IDs” column and clear “Comments” in the remaining rows.
	Change all values under the column “Assigned to” to “chetan.puwar”. 
	Finally, change status from Fixed to “Resolved”.
    
c.	Clear the filter and select only for “Dev Assigned” and then perform the following :

i.	Clear “#N/As” from “JIRA IDs” column. 

ii.	Clear all comments but leave only those pertaining to new tickets (i.e. rows having blank/clear JIRA IDs).

iii.	Change “Dev Assigned” to “Dev Progress” for AE and to “Tech Analysis” for OPT.

iv.	Change “Assigned to” for newOPT tickets to “unassigned”.

 Perform these steps forAE tickets only (not OPT) with Status of “New”. Delete all the rows for OPT tickets only with the status of “New” if found.
 
9.	Make sure to CLEAR ALL FILTERS, then add “project key” as new column and perform the following:

a.	Type RP for advisor and OPT for optimization.

b.	In “Work Stream”, change “Optimization” that to “Core Optimizations” and “Advisor Essential” to “Wealth360 5.1”.

10.	In order to correctly map to JIRA change the values under the “Severity” column as follows:

a.	Change “2-critical” to “critical”.

b.	Change “3-medium” to “major”.

c.	Change “4-low” to “minor”.

d.	Change “5-enhancement” to “trivial”.

11.	Now, in order to fill out a “resolution” for already closed defects… Please do the following:

a.	Add an extra column next to “project key” and name it “resolution”.

b.	For RP tickets, type “Done” for all tickets with a status “Closed” and leave the rest blank.

c.	For OPT tickets, match the resolution to their corresponding ones on JIRA. (Note: You’ll have to do a JIRA issue search for the OPT tickets)

12.	Now, to add new defect tickets as new links into JIRA… Please do the following:

a.	Add an extra column next to “resolution” and name it “relates”.

b.	Perform the following VLookupusing the “UseCaseList” spreadsheet.

i.	For “Lookup_Value” choose corresponding “Use Case”number.

ii.	For “Table_Array” choose the first two columns in the “UseCaseList” spreadsheet”.

iii.	For “Col_Index_Num” type in “2”.

iv.	For “[range_lookup]” type “false”.

v.	Clear “#N/A” and “0” if found within any rows in the newly created column.

vi.	Select all cells under “relates” column and the copy and paste (as values).
