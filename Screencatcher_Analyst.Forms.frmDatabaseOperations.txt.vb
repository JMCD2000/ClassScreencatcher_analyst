Option Compare Database
Option Explicit


Private Sub cmdAppendDateCoulmns_Click()
' This will append the date columns into the data tables
Dim contiuneTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdAppendDateCoulmns_Click"
contiuneTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder

'Call the AlterTablesAndColumns Module _
and run through the Sub
AppendDateFieldsToDataTables

If contiuneTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf contiuneTime = False Then
    'time error, do nothing
Else
    'untraped error
End If

End Sub


Private Sub cmdRunAll_Click()
' This will run all the updates with one click
Dim contiuneTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdRunAll_Click"
contiuneTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder

If contiuneTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf contiuneTime = False Then
    'time error, do nothing
Else
    'untraped error
End If

End Sub

Private Sub cmdSet_TABLE_Combined_Screenings_Click()
'This is to load or reload the screenings into the Combined_Screenings Table
Dim contiuneTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdSet_TABLE_Combined_Screenings_Click"
contiuneTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder

'Call the SetCombinedScreenings Module _
and run through all of the Subs
SetCurrentWorkingTable_CS 'Run the Current Working Table sub
SetFirstScreensAndEvents_CS 'This is to load or reload the screenings into the Combined_Screenings Table
SetNonShipEventScrns_CS 'This loads the screens that are not actual trials
SetLateAdds_TrialCards_CS 'This changes the screening from <Not Found> to <POST Trial>
SetClosedXX_TrialCards_CS 'This changes the screening from <Not Found> to <X/X>
SetTrialCardSplits_CS 'This changes the screening from <Not Found> to <SPLIT>

'Empty the lists and varibles used
ClearMyDateLists 'Empty the lists that were used
CurrentTable = Empty

If contiuneTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf contiuneTime = False Then
    'time error, do nothing
Else
    'untraped error
End If

End Sub


Private Sub cmdSet_TABLE_Screenings_Only_Click()
'This is to load or reload the screenings into the Screenings_Only Table
Dim contiuneTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdSet_TABLE_Screenings_Only_Click"
contiuneTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder

'Call the SetCombinedScreenings Module _
and run through all of the Subs
SetCurrentWorkingTable_SO 'Run the Current Working Table sub
SetFirstScreensAndEvents_SO 'This is to load or reload the screenings into the Combined_Screenings Table
SetNonShipEventScrns_SO 'This loads the screens that are not actual trials
SetLateAdds_TrialCards_SO 'This changes the screening from <Not Found> to <POST Trial>
SetClosedXX_TrialCards_SO 'This changes the screening from <Not Found> to <X/X>
SetTrialCardSplits_SO 'This changes the screening from <Not Found> to <SPLIT>

'Empty the lists and varibles used
ClearMyDateLists 'Empty the lists that were used
CurrentTable = Empty

If contiuneTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf contiuneTime = False Then
    'time error, do nothing
Else
    'untraped error
End If

End Sub


Private Sub cmdSet_TABLE_XX_Screen_Only_Click()
'This is to load or reload the screenings into the XX_Screen_Only Table
Dim contiuneTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdSet_TABLE_XX_Screen_Only_Click"
contiuneTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder

'Call the SetCombinedScreenings Module _
and run through all of the Subs
SetCurrentWorkingTable_XXSO 'Run the Current Working Table sub
SetFirstScreensAndEvents_XXSO 'This is to load or reload the screenings into the Combined_Screenings Table
SetNonShipEventScrns_XXSO 'This loads the screens that are not actual trials
SetLateAdds_TrialCards_XXSO 'This changes the screening from <Not Found> to <POST Trial>
SetClosedXX_TrialCards_XXSO 'This changes the screening from <Not Found> to <X/X>
SetTrialCardSplits_XXSO 'This changes the screening from <Not Found> to <SPLIT>

'Empty the lists and varibles used
ClearMyDateLists 'Empty the lists that were used
CurrentTable = Empty

If contiuneTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf contiuneTime = False Then
    'time error, do nothing
Else
    'untraped error
End If

End Sub


Private Sub cmdSet_TABLE_TC_Screen_Agg_Click()
'This is to load or reload the screenings into the TC_Screen_Agg Table
Dim contiuneTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdSet_TABLE_TC_Screen_Agg_Click"
contiuneTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder

'Call the SetCombinedScreenings Module _
and run through all of the Subs
SetCurrentWorkingTable_TCSA 'Run the Current Working Table sub
SetFirstScreensAndEvents_TCSA 'This is to load or reload the screenings into the Combined_Screenings Table
Build_and_Set_Aggregated_Screen_TCSA

'Empty the lists and varibles used
ClearMyDateLists 'Empty the lists that were used
CurrentTable = Empty

If contiuneTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf contiuneTime = False Then
    'time error, do nothing
Else
    'untraped error
End If

End Sub


Private Sub cmdSetSparseMatrix_Click()
'This is to load the Sparse Matrix tables
Dim contiuneTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdSetSparseMatrix_Click"
contiuneTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder

'Call the SetCombinedScreenings Module _
and run through all of the Subs

'This is to load or reload the screening sparse matrix into the <TABLE>_SparseMatrix
SetCurrentWorkingTable_SOSM "Screenings_Only_SparseMatrix", "Screenings_Only"
SetFirstScreensAndEvents_SOSM
Build_and_Set_Aggregated_Screen_SOSM 'This builds the sparse matrix that has a 0 if no change over the prior or has a 1 if there was a change.

CurrentTable = Empty
SparseRefTable = Empty

'This is to load or reload the screening sparse matrix into the <TABLE>_SparseMatrix
SetCurrentWorkingTable_SOSM "Combined_Screenings_SparseMatrix", "Combined_Screenings"
SetFirstScreensAndEvents_SOSM
Build_and_Set_Aggregated_Screen_SOSM 'This builds the sparse matrix that has a 0 if no change over the prior or has a 1 if there was a change.

'Empty the lists and varibles used
ClearMyDateLists 'Empty the lists that were used
CurrentTable = Empty
SparseRefTable = Empty

If contiuneTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf contiuneTime = False Then
    'time error, do nothing
Else
    'untraped error
End If

End Sub


Private Sub cmdSetXXSparseScreenMatrix_Click()
'This is to load the Screen Sparse Matrix table
Dim contiuneTime As Boolean
Dim thisProcedure As String

thisProcedure = "cmdSetXXSparseScreenMatrix_Click"
contiuneTime = startTime(thisProcedure)

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder

'Call the SetCombinedScreenings Module _
and run through all of the Subs

'This is to load or reload the screening sparse matrix into the <TABLE>_SparseMatrix
SetCurrentWorkingTable_XXSM
SetFirstScreensAndEvents_XXSM
Build_and_Set_Aggregated_Screen_XXSM 'This builds the sparse matrix that has a 0 if no change over the prior or has a 1 if there was a change.

'Empty the lists and varibles used
ClearMyDateLists 'Empty the lists that were used
CurrentTable = Empty
SparseRefTable = Empty

If contiuneTime = True Then
    'stop time
    endTime (thisProcedure)
ElseIf contiuneTime = False Then
    'time error, do nothing
Else
    'untraped error
End If

End Sub
