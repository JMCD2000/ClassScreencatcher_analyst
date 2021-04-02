Option Compare Database
Option Explicit


Private Sub cmdAppendDateColumns_Click()
' This will append the date columns into the data tables
Dim continueTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdAppendDateColumns_Click"
continueTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder

'Call the AlterTablesAndColumns Module _
and run through the Sub
AppendDateFieldsToDataTables

If continueTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf continueTime = False Then
    'time error, do nothing
Else
    'untraped error
End If

End Sub


Private Sub cmdClearDataTables_Click()
'This is to empty the data tables
Dim continueTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdClearDataTables_Click"
continueTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder

'Call the AlterTablesAndColumns Module _
and run through the Sub
ClearDataTables

If continueTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf continueTime = False Then
    'time error, do nothing
Else
    'untraped error
End If

End Sub


Private Sub cmdClearDateReportTables_Click()
'This is to empty the dated report tables
Dim continueTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdClearDateReportTables_Click"
continueTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder

'Call the AlterTablesAndColumns Module _
and run through the Sub
ClearDateReportTables

If continueTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf continueTime = False Then
    'time error, do nothing
Else
    'untraped error
End If

End Sub


Private Sub cmdCreateDataTables_Click()
'Create the report data tables from the list
Dim continueTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdCreateDataTables_Click"
continueTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder

'Call the AlterTablesAndColumns Module _
and run through the Sub
CreateDateReportTables

If continueTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf continueTime = False Then
    'time error, do nothing
Else
    'untraped error
End If

End Sub

Private Sub cmdDropDateColumns_Click()
'Drop the date columns in the data tables
Dim continueTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdDropDateColumns_Click"
continueTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder

'Call the AlterTablesAndColumns Module _
and run through the Sub
DropDateFieldsFromDataTables

If continueTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf continueTime = False Then
    'time error, do nothing
Else
    'untraped error
End If

End Sub


Private Sub cmdInsertTableKeys_Click()
'Drop the date columns in the data tables
Dim continueTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdInsertTableKeys_Click"
continueTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder

'Call the AlterTablesAndColumns Module
InsertTrialCardNumbers

If continueTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf continueTime = False Then
    'time error, do nothing
Else
    'untraped error
End If
End Sub


Private Sub cmdRunAll_Click()
' This will run all the updates with one click
'Dim continueTime As Boolean
'Dim thisProcedure As String

'thisProcedure = "cmdRunAll_Click"
'continueTime = startTime(thisProcedure)

'Cycle through each sub, run time is handled in the subs
cmdSet_TABLE_Combined_Screenings_Click
cmdSet_TABLE_Screenings_Only_Click
cmdSet_TABLE_XX_Screen_Only_Click
cmdSet_TABLE_TC_Screen_Agg_Click
cmdSetSparseMatrix_Click
cmdSetXXSparseScreenMatrix_Click

'If continueTime = True Then
'    'stop time
'    endTime (thisProcedure)
'ElseIf continueTime = False Then
'    'time error, do nothing
'Else
'    'untraped error
'End If

End Sub


Private Sub cmdSet_Column_Summations_Click()
'This sets the counts by column
Dim continueTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdSet_Column_Summations_Click"
continueTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder
'Call the SetListAndVars_Summary Module
AddingToMySummaryDateLists 'Run the list builder

'Call the SetCountsSummaryByColumn Module _
and run through the Sub
Set_All_SparseMatrix_SummaryCounts
SetRunningDays
SetDateDifferences

'Empty the lists and varibles used
ClearMyDateLists 'Empty the lists that were used
ClearMySummaryDateLists 'Empty the lists that were used

If continueTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf continueTime = False Then
    'time error, do nothing
Else
    'untraped error
End If

End Sub


Private Sub cmdSet_Row_Summations_Click()
'This sets the counts by row
Dim continueTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdSet_Column_Summations_Click"
continueTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder
'Call the SetListAndVars_Summary Module
AddingToMySummaryDateLists 'Run the list builder

'Call the SetCountsSummaryByRow Module _
and run through the Sub
'SetRecordReScreenCounts

'Empty the lists and varibles used
ClearMyDateLists 'Empty the lists that were used
ClearMySummaryDateLists 'Empty the lists that were used

If continueTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf continueTime = False Then
    'time error, do nothing
Else
    'untraped error
End If

End Sub

Private Sub cmdSet_TABLE_Combined_Screenings_Click()
'This is to load or reload the screenings into the Combined_Screenings Table
Dim continueTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdSet_TABLE_Combined_Screenings_Click"
continueTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder

'Call the SetCombinedScreenings Module and run through all of the Subs
CurrentTable = "All_Combined_Screenings"
All_or_Events = "All"
SetFirstScreensAndEvents_CS 'This is to load or reload the screenings into the Combined_Screenings Table
SetNonShipEventScrns_CS 'This loads the screens that are not actual trials
SetLateAdds_TrialCards_CS 'This changes the screening from <Not Found> to <POST Trial>
SetClosedXX_TrialCards_CS 'This changes the screening from <Not Found> to <X/X>
SetTrialCardSplits_CS 'This changes the screening from <Not Found> to <SPLIT>
SetTrialCardsMissingFromReports_CS allColumnsList_BT, allColumnsList_AT
SetTrialCardsMissingFromReports_CS allColumnsList_AT, allColumnsList_FCT
SetTrialCardsMissingFromReports_CS allColumnsList_FCT, allColumnsList_OWLD
CurrentTable = Empty
All_or_Events = Empty

CurrentTable = "Events_Combined_Screenings"
All_or_Events = "Events"
SetFirstScreensAndEvents_CS 'This is to load or reload the screenings into the Combined_Screenings Table
SetLateAdds_TrialCards_CS 'This changes the screening from <Not Found> to <POST Trial>
SetClosedXX_TrialCards_CS 'This changes the screening from <Not Found> to <X/X>
SetTrialCardSplits_CS 'This changes the screening from <Not Found> to <SPLIT>

'Empty the lists and varibles used
ClearMyDateLists 'Empty the lists that were used
CurrentTable = Empty
All_or_Events = Empty

If continueTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf continueTime = False Then
    'time error, do nothing
Else
    'untraped error
End If

End Sub


Private Sub cmdSet_TABLE_Screenings_Only_Click()
'This is to load or reload the screenings into the Screenings_Only Table
Dim continueTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdSet_TABLE_Screenings_Only_Click"
continueTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder

'Call the SetScreeningsOnly Module and run through all of the Subs
CurrentTable = "All_Screenings_Only"
All_or_Events = "All"
SetFirstScreensAndEvents_SO 'This is to load or reload the screenings into the Combined_Screenings Table
SetNonShipEventScrns_SO 'This loads the screens that are not actual trials
SetLateAdds_TrialCards_SO 'This changes the screening from <Not Found> to <POST Trial>
SetClosedXX_TrialCards_SO 'This changes the screening from <Not Found> to <X/X>
SetTrialCardSplits_SO 'This changes the screening from <Not Found> to <SPLIT>
SetTrialCardsMissingFromReports_CS allColumnsList_BT, allColumnsList_AT
SetTrialCardsMissingFromReports_CS allColumnsList_AT, allColumnsList_FCT
SetTrialCardsMissingFromReports_CS allColumnsList_FCT, allColumnsList_OWLD
CurrentTable = Empty
All_or_Events = Empty

CurrentTable = "Events_Screenings_Only"
All_or_Events = "Events"
SetFirstScreensAndEvents_SO 'This is to load or reload the screenings into the Combined_Screenings Table
SetLateAdds_TrialCards_SO 'This changes the screening from <Not Found> to <POST Trial>
SetClosedXX_TrialCards_SO 'This changes the screening from <Not Found> to <X/X>
SetTrialCardSplits_SO 'This changes the screening from <Not Found> to <SPLIT>

'Empty the lists and varibles used
ClearMyDateLists 'Empty the lists that were used
CurrentTable = Empty
All_or_Events = Empty

If continueTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf continueTime = False Then
    'time error, do nothing
Else
    'untraped error
End If

End Sub


Private Sub cmdSet_TABLE_XX_Screen_Only_Click()
'This is to load or reload the screenings into the XX_Screen_Only Table
Dim continueTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdSet_TABLE_XX_Screen_Only_Click"
continueTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder

'Call the SetXXScreenOnly Module and run through all of the Subs
CurrentTable = "All_XX_Screen_Only"
All_or_Events = "All"
SetFirstScreensAndEvents_XXSO 'This is to load or reload the screenings into the Combined_Screenings Table
SetNonShipEventScrns_XXSO 'This loads the screens that are not actual trials
SetLateAdds_TrialCards_XXSO 'This changes the screening from <Not Found> to <POST Trial>
SetClosedXX_TrialCards_XXSO 'This changes the screening from <Not Found> to <X/X>
SetTrialCardSplits_XXSO 'This changes the screening from <Not Found> to <SPLIT>
SetTrialCardsMissingFromReports_CS allColumnsList_BT, allColumnsList_AT
SetTrialCardsMissingFromReports_CS allColumnsList_AT, allColumnsList_FCT
SetTrialCardsMissingFromReports_CS allColumnsList_FCT, allColumnsList_OWLD
CurrentTable = Empty
All_or_Events = Empty

CurrentTable = "Events_XX_Screen_Only"
All_or_Events = "Events"
SetFirstScreensAndEvents_XXSO 'This is to load or reload the screenings into the Combined_Screenings Table
SetLateAdds_TrialCards_XXSO 'This changes the screening from <Not Found> to <POST Trial>
SetClosedXX_TrialCards_XXSO 'This changes the screening from <Not Found> to <X/X>
SetTrialCardSplits_XXSO 'This changes the screening from <Not Found> to <SPLIT>

'Empty the lists and varibles used
ClearMyDateLists 'Empty the lists that were used
CurrentTable = Empty
All_or_Events = Empty

If continueTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf continueTime = False Then
    'time error, do nothing
Else
    'untraped error
End If

End Sub


Private Sub cmdSet_TABLE_TC_Screen_Agg_Click()
'This is to load or reload the screenings into the TC_Screen_Agg Table
Dim continueTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdSet_TABLE_TC_Screen_Agg_Click"
continueTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder

'Call the SetScreenAgg Module and run through all of the Subs
CurrentTable = "All_TC_Screen_Agg"
All_or_Events = "All"
SetFirstScreensAndEvents_TCSA 'This is to load or reload the screenings into the Combined_Screenings Table
Build_and_Set_Aggregated_Screen_and_Counts_TCSA
'zBuild_and_Set_Aggregated_Screen_Lifetime_TCSA
'zBuild_and_Set_Aggregated_Screen_OWLD_Limit_TCSA
'zBuild_and_Set_Aggregated_Screen_DEL_Limit_TCSA

CurrentTable = Empty
All_or_Events = Empty

CurrentTable = "Events_TC_Screen_Agg"
All_or_Events = "Events"
SetFirstScreensAndEvents_TCSA 'This is to load or reload the screenings into the Combined_Screenings Table
Build_and_Set_Aggregated_Screen_and_Counts_TCSA
'zBuild_and_Set_Aggregated_Screen_Lifetime_TCSA
'zBuild_and_Set_Aggregated_Screen_OWLD_Limit_TCSA
'zBuild_and_Set_Aggregated_Screen_DEL_Limit_TCSA

'Empty the lists and varibles used
ClearMyDateLists 'Empty the lists that were used
CurrentTable = Empty
All_or_Events = Empty

If continueTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf continueTime = False Then
    'time error, do nothing
Else
    'untraped error
End If

End Sub


Private Sub cmdSetSparseMatrix_Click()
'This is to load the Sparse Matrix tables
Dim continueTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdSetSparseMatrix_Click"
continueTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder

'Call the SetScreenSparseMatrix Module and run through all of the Subs
'This is to load or reload the screening sparse matrix into the <TABLE>_SparseMatrix
CurrentTable = "All_Screenings_Only_SparseMatrix"
SparseRefTable = "All_Screenings_Only"
All_or_Events = "All"
SetFirstScreensAndEvents_SOSM
Build_and_Set_Aggregated_Screen_SOSM 'This builds the sparse matrix that has a 0 if no change over the prior or has a 1 if there was a change.
Set_SparseMatrix_Counts_SMC

CurrentTable = Empty
SparseRefTable = Empty
All_or_Events = Empty

'This is to load or reload the screening sparse matrix into the <TABLE>_SparseMatrix
CurrentTable = "All_Combined_Screenings_SparseMatrix"
SparseRefTable = "All_Combined_Screenings"
All_or_Events = "All"
SetFirstScreensAndEvents_SOSM
Build_and_Set_Aggregated_Screen_SOSM 'This builds the sparse matrix that has a 0 if no change over the prior or has a 1 if there was a change.
Set_SparseMatrix_Counts_SMC

CurrentTable = Empty
SparseRefTable = Empty
All_or_Events = Empty

'Call the SetScreenSparseMatrix Module and run through all of the Subs
'This is to load or reload the screening sparse matrix into the <TABLE>_SparseMatrix
CurrentTable = "Events_Screenings_Only_SparseMatrix"
SparseRefTable = "Events_Screenings_Only"
All_or_Events = "Events"
SetFirstScreensAndEvents_SOSM
Build_and_Set_Aggregated_Screen_SOSM 'This builds the sparse matrix that has a 0 if no change over the prior or has a 1 if there was a change.
Set_SparseMatrix_Counts_SMC

CurrentTable = Empty
SparseRefTable = Empty
All_or_Events = Empty

'This is to load or reload the screening sparse matrix into the <TABLE>_SparseMatrix
CurrentTable = "Events_Combined_Screenings_SparseMatrix"
SparseRefTable = "Events_Combined_Screenings"
All_or_Events = "Events"
SetFirstScreensAndEvents_SOSM
Build_and_Set_Aggregated_Screen_SOSM 'This builds the sparse matrix that has a 0 if no change over the prior or has a 1 if there was a change.
Set_SparseMatrix_Counts_SMC

'Empty the lists and varibles used
ClearMyDateLists 'Empty the lists that were used
CurrentTable = Empty
SparseRefTable = Empty
All_or_Events = Empty

If continueTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf continueTime = False Then
    'time error, do nothing
Else
    'untraped error
End If

End Sub


Private Sub cmdSetXXSparseScreenMatrix_Click()
'This is to load the Screen Sparse Matrix table
Dim continueTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdSetXXSparseScreenMatrix_Click"
continueTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder

'Call the SetCombinedScreenings Module and run through all of the Subs
'This is to load or reload the screening sparse matrix into the <TABLE>_SparseMatrix
CurrentTable = "All_XX_Screen_Only_SparseMatrix"
SparseRefTable = "All_XX_Screen_Only"
All_or_Events = "All"
SetFirstScreensAndEvents_XXSM
Build_and_Set_Aggregated_Screen_XXSM 'This builds the sparse matrix that has a 0 if no change over the prior or has a 1 if there was a change.
Set_SparseMatrix_Counts_XXSMC

CurrentTable = Empty
SparseRefTable = Empty
All_or_Events = Empty

CurrentTable = "Events_XX_Screen_Only_SparseMatrix"
SparseRefTable = "Events_XX_Screen_Only"
All_or_Events = "Events"
SetFirstScreensAndEvents_XXSM
Build_and_Set_Aggregated_Screen_XXSM 'This builds the sparse matrix that has a 0 if no change over the prior or has a 1 if there was a change.
Set_SparseMatrix_Counts_XXSMC

'Empty the lists and varibles used
ClearMyDateLists 'Empty the lists that were used
CurrentTable = Empty
SparseRefTable = Empty
All_or_Events = Empty

If continueTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf continueTime = False Then
    'time error, do nothing
Else
    'untraped error
End If

End Sub
