Option Compare Database
Option Explicit


Public Sub SetCurrentWorkingTable_CS()
'This sub is setting the current working _
table that is used in all the SQL statements.

CurrentTable = "Combined_Screenings"

End Sub


Public Sub SetFirstScreensAndEvents_CS()
'This is to load or reload the screenings into the Combined_Screenings Table _
1st text <Not Found> is entered in every field _
2nd Final is loaded into final(far most right date) and into [First_Screening] _
3rd FCT is loaded in to FCT date and update of [First_Screening] _
4th AT is loaded in to AT date and update of [First_Screening] _
5th BT is loaded in to BT date and update of [First_Screening]

'Run the Current Working Table sub
'    SetCurrentWorkingTable_CS
'Run the list builder
'    AddingToMyDateLists

'Set all date columns to <Not Found>
    Dim myDateVarList As Variant
    Dim notFound As String 'The <Not Found> is not used in TSM or elsewhere, becomes a visual that something was missed
    notFound = "Not Found"

    For Each myDateVarList In allColumnsList
        CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
        & "SET " _
        & "" & CurrentTable & ".[" & myDateVarList & "] = '" & notFound & "';"
        ' Debug.Print "done with column table: " & myDateVarList & "."
    Next myDateVarList
    
    myDateVarList = Empty

'Set TC data and first screens as place holder values
    Dim emptyID As String 'The dash is not used in TSM or elsewhere, becomes a visual that something was missed
    emptyID = "-"
    
    Dim emptyEvent As String 'The double E is not used in TSM or elsewhere, becomes a visual that something was missed
    emptyEvent = "EE"
    
    Dim emptySts_A_T As String 'The dash slash dash is not used in TSM or elsewhere, becomes a visual that something was missed
    emptySts_A_T = "-/-"

    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".Trial_ID = '" & emptyID & "', " _
    & "" & CurrentTable & ".Event = '" & emptyEvent & "', " _
    & "" & CurrentTable & ".Final_Sts_A_T = '" & emptySts_A_T & "', " _
    & "" & CurrentTable & ".First_Screening = '" & notFound & "';"
    ' Debug.Print vbCrLf & "Completed setting place holder values in columns Trial_ID, Event, Final_Sts_A_T, First_Screening."
    
' Debug.Print vbCrLf & "Completed the" & CurrentTable & "table data set with place holder values Update Query."

'Set Final Screening as Final and as First
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".First_Screening = [" & beanFinal & "].[TC_Screening], " _
    & "" & CurrentTable & ".[" & columnFinal & "] = [" & beanFinal & "].[TC_Screening], " _
    & "" & CurrentTable & ".Final_Sts_A_T = [" & beanFinal & "].[Final_Sts_A_T], " _
    & "" & CurrentTable & ".Trial_ID = [" & beanFinal & "].[Trial_ID], " _
    & "" & CurrentTable & ".Event = [" & beanFinal & "].[Event];"
    ' Debug.Print vbCrLf & "Completed setting values in columns Trial_ID, Event, Final_Sts_A_T, First_Screening."

'Set FCT Event Screening and First Screening, FCT upload
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFCT & "] ON " & CurrentTable & ".Trial_Card = [" & beanFCT & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".First_Screening = [" & beanFCT & "].[TC_Screening], " _
    & "" & CurrentTable & ".[" & columnFCT & "] = [" & beanFCT & "].[TC_Screening];"
    ' Debug.Print vbCrLf & "Completed setting FCT Event Screening and First Screening"

'Set AT Event Screening and First Screening, AT upload
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanAT & "] ON " & CurrentTable & ".Trial_Card = [" & beanAT & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".First_Screening = [" & beanAT & "].[TC_Screening], " _
    & "" & CurrentTable & ".[" & columnAT & "] = [" & beanAT & "].[TC_Screening];"
    ' Debug.Print vbCrLf & "Completed setting AT Event Screening and First Screening"

'Set BT Event Screening and First Screening, BT upload
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanBT & "] ON " & CurrentTable & ".Trial_Card = [" & beanBT & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".First_Screening = [" & beanBT & "].[TC_Screening], " _
    & "" & CurrentTable & ".[" & columnBT & "] = [" & beanBT & "].[TC_Screening];"
    ' Debug.Print vbCrLf & "Completed setting BT Event Screening and First Screening"
    
'Set OWLD Event Screening, Transfer Book Bean Data
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanOWLD & "] ON " & CurrentTable & ".Trial_Card = [" & beanOWLD & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".[" & columnOWLD & "] = [" & beanOWLD & "].[TC_Screening];"
    ' Debug.Print vbCrLf & "Completed setting OWLD Event Screening"

'ClearMyDateLists 'Empty created list objects

' Debug.Print vbCrLf & "The Trials First Screenings and OWLD Update Query completed." & vbCrLf

End Sub


Public Sub SetNonShipEventScrns_CS()
'This loads the screens that are not actual trials

'Run the list builder
'AddingToMyDateLists

Dim myDateVarList As Variant

For Each myDateVarList In nonTrialsList
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " INNER JOIN [" & myDateVarList & "_" & curHullNum & "] ON " & CurrentTable & ".Trial_Card = [" & myDateVarList & "_" & curHullNum & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".[" & myDateVarList & "] = [" & myDateVarList & "_" & curHullNum & "].[TC_Screening];"
    ' Debug.Print "done with column table: " & myDateVarList & "."
Next myDateVarList

myDateVarList = Empty

'ClearMyDateLists 'Empty created list objects

' Debug.Print vbCrLf & "Completed the Non Trial Screenings Update Query." & vbCrLf

End Sub


Public Sub SetLateAdds_TrialCards_CS()
'This changes the screening from <Not Found> to <POST Trial> _
where the trial card was entered after the actual inspection.
Dim notFound As String
notFound = "Not Found"

'Run the list builder
'AddingToMyDateLists

Dim myDateVarList As Variant

'Run for BT Event
For Each myDateVarList In allColumnsList
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".[" & myDateVarList & "] = 'POST BT Trial'" _
    & "WHERE " _
    & "(((" & CurrentTable & ".[" & myDateVarList & "])='Not Found') " _
    & "AND (([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
    & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))='01') " _
    & "AND ((" & CurrentTable & ".Event)='BT'));"
    ' Debug.Print "done with column table: " & myDateVarList & " For BT Event Late Adds."
Next myDateVarList

myDateVarList = Empty

'Run for AT Event _
This is marking new AT trial cards as POST AT Trial, Incorrect marking _
the fix for this is to use a list that only starts at the AT Event
For Each myDateVarList In allColumnsList
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".[" & myDateVarList & "] = 'POST AT Trial'" _
    & "WHERE " _
    & "(((" & CurrentTable & ".[" & myDateVarList & "])='Not Found') " _
    & "AND (([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
    & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))='01') " _
    & "AND ((" & CurrentTable & ".Event)='AT'));"
    ' Debug.Print "done with column table: " & myDateVarList & " For AT Event Late Adds."
Next myDateVarList
'This is fixing the Incorrect marking on days prior to the AT _
there are list populated dates in this SQL Query!
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".[" & PreAT_DatesVarList(0) & "] = 'Not Found', " _
    & "" & CurrentTable & ".[" & PreAT_DatesVarList(1) & "] = 'Not Found'" _
    & "WHERE " _
    & "(((" & CurrentTable & ".[" & PreAT_DatesVarList(0) & "])='POST AT Trial') " _
    & "AND ((" & CurrentTable & ".[" & PreAT_DatesVarList(1) & "])='POST AT Trial') " _
    & "AND (([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
    & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))='01') " _
    & "AND ((" & CurrentTable & ".Event)='AT'));"
    ' Debug.Print vbCrLf & "Completed re-setting prior to AT Event Screenings to Not Found."

myDateVarList = Empty

'Run for FCT Event _
This is marking new FCT trial cards as POST FCT Trial, Incorrect marking _
the fix for this is to use a list that only starts at the FCT Event
For Each myDateVarList In allColumnsList
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".[" & myDateVarList & "] = 'POST FCT Trial'" _
    & "WHERE " _
    & "(((" & CurrentTable & ".[" & myDateVarList & "])='Not Found') " _
    & "AND (([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
    & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))='01') " _
    & "AND ((" & CurrentTable & ".Event)='FCT'));"
    ' Debug.Print "done with column table: " & myDateVarList & " For FCT Event Late Adds."
Next myDateVarList
'This is fixing the Incorrect marking on days prior to the FCT _
there are hard coded dates in this SQL Query!
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".[" & PreFCT_DatesVarList(0) & "] = 'Not Found', " _
    & "" & CurrentTable & ".[" & PreFCT_DatesVarList(1) & "] = 'Not Found', " _
    & "" & CurrentTable & ".[" & PreFCT_DatesVarList(2) & "] = 'Not Found', " _
    & "" & CurrentTable & ".[" & PreFCT_DatesVarList(3) & "] = 'Not Found', " _
    & "" & CurrentTable & ".[" & PreFCT_DatesVarList(4) & "] = 'Not Found', " _
    & "" & CurrentTable & ".[" & PreFCT_DatesVarList(5) & "] = 'Not Found', " _
    & "" & CurrentTable & ".[" & PreFCT_DatesVarList(6) & "] = 'Not Found', " _
    & "" & CurrentTable & ".[" & PreFCT_DatesVarList(7) & "] = 'Not Found', " _
    & "" & CurrentTable & ".[" & PreFCT_DatesVarList(8) & "] = 'Not Found'" _
    & "WHERE " _
    & "(((" & CurrentTable & ".[" & PreFCT_DatesVarList(0) & "])='POST FCT Trial') " _
    & "AND ((" & CurrentTable & ".[" & PreFCT_DatesVarList(1) & "])='POST FCT Trial') " _
    & "AND ((" & CurrentTable & ".[" & PreFCT_DatesVarList(2) & "])='POST FCT Trial') " _
    & "AND ((" & CurrentTable & ".[" & PreFCT_DatesVarList(3) & "])='POST FCT Trial') " _
    & "AND ((" & CurrentTable & ".[" & PreFCT_DatesVarList(4) & "])='POST FCT Trial') " _
    & "AND ((" & CurrentTable & ".[" & PreFCT_DatesVarList(5) & "])='POST FCT Trial') " _
    & "AND ((" & CurrentTable & ".[" & PreFCT_DatesVarList(6) & "])='POST FCT Trial') " _
    & "AND ((" & CurrentTable & ".[" & PreFCT_DatesVarList(7) & "])='POST FCT Trial') " _
    & "AND ((" & CurrentTable & ".[" & PreFCT_DatesVarList(8) & "])='POST FCT Trial') " _
    & "AND (([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
    & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))='01') " _
    & "AND ((" & CurrentTable & ".Event)='FCT'));"
    ' Debug.Print vbCrLf & "Completed re-setting prior to FCT Event Screenings to Not Found."
    
'myDateVarList = Empty

' Debug.Print vbCrLf & "The Late Add Trial Cards Update Query completed." & vbCrLf

End Sub


Public Sub SetClosedXX_TrialCards_CS()
'This changes the screening from <Not Found> to <X/X> _
where the trial card was reidentified at a later Event.

'Run the list builder
'AddingToMyDateLists

Dim myDateVarList As Variant

'Run for all the X/X cards where screen is Not Found
For Each myDateVarList In allColumnsList
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".[" & myDateVarList & "] = 'X/X'" _
    & "WHERE " _
    & "(((" & CurrentTable & ".[" & myDateVarList & "])='Not Found') " _
    & "AND (([" & CurrentTable & "]![Final_Sts_A_T])='X/X'));"
    ' Debug.Print "done with column table: " & myDateVarList & " For X/X Closures."
Next myDateVarList

myDateVarList = Empty

'ClearMyDateLists 'Empty created list objects

' Debug.Print vbCrLf & "The X/X Trial Cards Update Query completed." & vbCrLf

End Sub


Public Sub SetTrialCardSplits_CS()
'This changes the screening from <Not Found> to <SPLIT> _
where the trial card was entered after the actual inspection.

'Run the list builder
'AddingToMyDateLists

Dim myDateVarList As Variant

'Run for BT Event _
this does not catch
For Each myDateVarList In allColumnsList
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
    & "SET " & CurrentTable & ".[" & myDateVarList & "] = 'SPLIT'" _
    & "WHERE (((" & CurrentTable & ".[" & myDateVarList & "])='Not Found') " _
    & "AND (([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
    & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))<>'01') " _
    & "AND ((" & CurrentTable & ".Event)='BT'));"
    ' Debug.Print "done with column table: " & myDateVarList & " For BT Event Splits."
Next myDateVarList

myDateVarList = Empty

'Run for AT Event
For Each myDateVarList In allColumnsList
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
    & "SET " & CurrentTable & ".[" & myDateVarList & "] = 'SPLIT'" _
    & "WHERE (((" & CurrentTable & ".[" & myDateVarList & "])='Not Found') " _
    & "AND (([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
    & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))<>'01') " _
    & "AND ((" & CurrentTable & ".Event)='AT'));"
    ' Debug.Print "done with column table: " & myDateVarList & " For AT Event Splits."
Next myDateVarList

myDateVarList = Empty

'Run for FCT Event
For Each myDateVarList In allColumnsList
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
    & "SET " & CurrentTable & ".[" & myDateVarList & "] = 'SPLIT'" _
    & "WHERE (((" & CurrentTable & ".[" & myDateVarList & "])='Not Found') " _
    & "AND (([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
    & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))<>'01') " _
    & "AND ((" & CurrentTable & ".Event)='FCT'));"
    ' Debug.Print "done with column table: " & myDateVarList & " For FCT Event Splits."
Next myDateVarList

myDateVarList = Empty

'ClearMyDateLists 'Empty created list objects

' Debug.Print vbCrLf & "The Split Trial Cards Update Query completed." & vbCrLf

End Sub


Public Sub needs_work_CS()
'This concats the screen values into a single field to show screen transitions.
Dim firstScreenVar As String ' This holds the first screen
Dim currentScreenVar As String ' This holds the last column looked at value to see if current column is different
Dim aggregatedScreenings As String ' This holds the screens as they are collected

'Run the list builder
'AddingToMyDateLists

Dim myDateVarList As Variant

'Run for all the trial cards where screen is populated
For Each myDateVarList In allColumnsList
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN " & CurrentTable & " ON " & CurrentTable & ".Trial_Card = " & CurrentTable & ".Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".First_Screening = Left([" & CurrentTable & "].[First_Screening],(Len([" & CurrentTable & "].[First_Screening])-4));"
    ' Debug.Print "done with column table: " & myDateVarList & " For Agg of First Screenings."
Next myDateVarList

myDateVarList = Empty

'ClearMyDateLists 'Empty created list objects

' Debug.Print vbCrLf & "The X/X Trial Cards Update Query completed." & vbCrLf

End Sub


