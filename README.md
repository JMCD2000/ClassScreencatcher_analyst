# ClassScreencatcher_analyst
This takes weekly Excel reports and builds combined tables and metrics in Access and Excel with VBA, when loaded manually. For automated loading of the weekly reports tables, I use the ScreeningHistory project to select the folder containing the excel files and then select this ClassScreecatcher_analyst project. The ScreeningHistory does all the data transformations, data validation, and loading via VBA.

There is an Excel file that I use to build the Lists and Varibles that are used throughout the VBA code. This makes the date names for the tables, assigns varibles to the dates and to the tables. This also provides the list order and list contents needed for the changes between Event dates.

I would have loved to have done this project in Python, but my work doesn't allow it. This runs using VBA written in Modules in MS Access 2016. The more tables there are the slower this runs. This is just an issue with using VBA and 32bit office. Again, it is what my work allows and has configured.

I run weekly Excel exports from an Oricle database that provides a status of open items, a burn rate, a class comparision, and responsibity for action. These weekly reports are used to build the data tables. This ClassScreencatcher_Analyst is then used to find the churn of rescreenings, changes in responsibility from Government to Contractor, Number of screenings and combinations, and other metrics. 

Alot of this code was originaly started in Excel, but became unmanigible due to the number of calculations needed and run. Moving the computation to VBA in the Database reduced the Excel size from 76MB to 12MB. These are now premade Excel files for AT, FCT, and Production to Sailaway.

I have now added the ability to insert the data tables primary values from the final table. This was a manual drill before and was a pain with multiple tables and when setting up for a new ship. I also added the ability for the table list to append columns to the data tables insted of doing it manualy. On the same line the columns can be droped when setting up a new DB.

Below was the manual clean up that was needed before direct importing the weekly Excel reports into the Access Screencatcher_Analyst database. The Concat formulas are put into the last two or three columns. It was easier to pre populate these values in excel instead of adding more complexity in Access or the VBA. The use of this manual cleanup was time consuming and could be error prone. This is now handeled on import with VBA.

1) Column header cleanup
(used on the Final only)
Trial_Card	Star	Pri	Safety	Screening	Act_1	Act_2	Status	Action_Taken	Date_Discovered	Date_Closed	Trial_ID	Event	TC_Screening	TC_Screening_AC1_AC2	Final_Sts_A_T
=CONCATENATE(E2,"/",F2,"/",G2," ",H2,"/",I2)	=CONCATENATE(E2,"/",F2,"/",G2)	=CONCATENATE(,H2,"/",I2)

(used on all other reports)
Trial_Card	Star	Pri	Safety	Screening	Act_1	Act_2	Status	Action_Taken	Date_Discovered	Date_Closed	Trial_ID	Event	TC_Screening	TC_Screening_AC1_AC2
=CONCATENATE(E2,"/",F2,"/",G2," ",H2,"/",I2)	=CONCATENATE(E2,"/",F2,"/",G2)

2) Clean up cell values to be compatible with Access
CONVERT
	STAR="*" TO "STAR"
	SCREEN="**" TO "AST"; clear out any bad symboles, "-", "@"
	AC1/2="****" TO "AST"; clear out any bad symboles, "-", "@"
	"ORACLE DATES" DASHes TO "MICROSOFT DATES" SLASHes
	CLEAR OUT EMPTY DATE FIELDS
	Columns A:I and L:P format as Text
	Columns J:K format as Date
	Sort A-Z on Trial_Card and Save

3) On import to Access save the new table with this format.
Nontrial beans
YYYY/MM/DD_LPD17

Trial beans
YYYY/MM/DD_LPD17_BT
YYYY/MM/DD_LPD17_AT
YYYY/MM/DD_LPD17_FCT
YYYY/MM/DD_LPD17_OWLD
YYYY/MM/DD_LPD17_Final

This starts with getting all the Excel reports into one file.
Building the "LPDXX Date Values and TC Numbers(v).xlsx" to get the dates and varibles needed for table names and the lists.
The data is copied from "LPDXX Date Values and TC Numbers(v).xlsx" into Modules "SetListAndVars" and "SetListAndVars_Summary.
Then using the Screencatcher_Analyst frmDatabaseOperations you build the data tables that the Excel reports will be loaded into.
Then using the ScreeningHistory(vX) project you point the the Excel folder and the Screencatcher_Analyst database. Any data errors or skipped rows or spreadsheets are logged in the ScreeningHistory(vX) database for review.
Then using the Screencatcher_Analyst frmDatabaseOperations you build and populate an all the aggrated tables and summaries.
These are then exported to Excel for charting and analysis with some premade Excel files for AT, FCT, and Production to Sailaway.

