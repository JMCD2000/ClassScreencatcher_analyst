# ClassScreencatcher_analyst
This takes weekly reports and builds combined tables and metrics in Access and Excel with VBA

I would have loved to have done this project in Python, but my work doesn't allow it. This runs using VBA written in Modules in MS Access 2016. The more tables there are the slower this runs. This is just an issue with using VBA and 32bit office. Again, it is what my work allows and has configured.

I run weekly exports from a database that exports to excel to provide a status of open items, a burn rate, a class comparision, and responsibity for action. These weekly reports are used to build the data tables. This ClassScreencatcher_Analyst is then used to find the churn of rescreenings, changes in responsibility from Government to Contractor, Number of screenings and combinations, and other metrics. 

Alot of this code was originaly started in Excel, but became unmanigible due to the number of calculations needed and run. Removing the computation to VBA at the Database reduced the Excel size from 76MB to 12MB.

Below is the manual clean up that is needed before importing the Excel into Access. The Concat formulas are put into the last two or three columns. It was easier to pre populate these values in excel instead of adding more complexity in Access or the VBA.

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
