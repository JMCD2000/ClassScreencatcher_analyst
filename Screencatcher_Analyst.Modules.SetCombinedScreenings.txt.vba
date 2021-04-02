Option Compare Database
'The use of NotFound until the trial card is actually written at Event just looks messy, would like to change this
Option Explicit


Public Sub SetFirstScreensAndEvents_CS()
'This is to load or reload the screenings into the Combined_Screenings Table

'Set all date columns to <Not Found>
Dim myDateVarList As Variant
Dim notFound As String 'The <Not Found> is not used in TSM or elsewhere, becomes a visual that something was missed
notFound = "Not Found"
    
    'Check for All Reports or only Events
    If All_or_Events = "All" Then
        For Each myDateVarList In allColumnsList
            CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
            & "SET " _
            & "" & CurrentTable & ".[" & myDateVarList & "] = '" & notFound & "';"
            ' Debug.Print "done with column table: " & myDateVarList & "."
        Next myDateVarList
    ElseIf All_or_Events = "Events" Then
        For Each myDateVarList In trialsOnlyList
            CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
            & "SET " _
            & "" & CurrentTable & ".[" & myDateVarList & "] = '" & notFound & "';"
            ' Debug.Print "done with column table: " & myDateVarList & "."
        Next myDateVarList
    Else
        'Un trapped error
        'All_or_Events Global is empty or not expected value
        Debug.Print "Function SetFirstScreensAndEvents_CS() was passed empty or not expected value with GLOBAL All_or_Events:= " & All_or_Events & "."
    End If
    
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
    ' Debug.Print vbCrLf & "Completed the " & CurrentTable & " table data set with place holder values Update Query."

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
    
    'Set DEL Event Screening, DEL Milestone
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanDEL & "] ON " & CurrentTable & ".Trial_Card = [" & beanDEL & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".[" & columnDEL & "] = [" & beanDEL & "].[TC_Screening];"
    ' Debug.Print vbCrLf & "Completed setting DEL Event Screening

    ' Debug.Print vbCrLf & "The Trials First Screenings and OWLD Update Query completed." & vbCrLf

End Sub


Public Sub SetNonShipEventScrns_CS()
'This loads the screens that are not actual trials
'This is not needed for the Events only tables

Dim myDateVarList As Variant

    'Check for All Reports or only Events
    If All_or_Events = "All" Then
        For Each myDateVarList In nonTrialsList
            CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " INNER JOIN [" & myDateVarList & "_" & curHullNum & "] ON " & CurrentTable & ".Trial_Card = [" & myDateVarList & "_" & curHullNum & "].Trial_Card " _
            & "SET " _
            & "" & CurrentTable & ".[" & myDateVarList & "] = [" & myDateVarList & "_" & curHullNum & "].[TC_Screening];"
            ' Debug.Print "done with column table: " & myDateVarList & "."
        Next myDateVarList
    ElseIf All_or_Events = "Events" Then
        'This is not needed for the Events Tables
    Else
        'Un trapped error
        'All_or_Events Global is empty or not expected value
        Debug.Print "Function SetNonShipEventScrns_CS() was passed empty or not expected value with GLOBAL All_or_Events:= " & All_or_Events & "."
    End If
    
    myDateVarList = Empty
    
    ' Debug.Print vbCrLf & "Completed the Non Trial Screenings Update Query." & vbCrLf

End Sub


Public Sub SetLateAdds_TrialCards_CS()
'This changes the screening from <Not Found> to <POST Trial> _
where the trial card was entered after the actual inspection.
Dim notFound As String
notFound = "Not Found"

Dim myDateVarList As Variant

    'Check for All Reports or only Events
    If All_or_Events = "All" Then
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
        
        myDateVarList = Empty
        
        'This is fixing the Incorrect marking on days prior to the AT _
        these dates are in the "Pre..." lists.
        For Each myDateVarList In PreAT_DatesVarList
            CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
            & "SET " _
            & "" & CurrentTable & ".[" & myDateVarList & "] = 'Not Found'" _
            & "WHERE " _
            & "(((" & CurrentTable & ".[" & myDateVarList & "])='POST AT Trial') " _
            & "AND (([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
            & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))='01') " _
            & "AND ((" & CurrentTable & ".Event)='AT'));"
            ' Debug.Print "done with column table: " & myDateVarList & " For AT Event Late Adds."
        Next myDateVarList
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
        
        myDateVarList = Empty
        
        'This is fixing the Incorrect marking on days prior to the FCT _
        these dates are in the "Pre..." lists.
        For Each myDateVarList In PreFCT_DatesVarList
            CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
            & "SET " _
            & "" & CurrentTable & ".[" & myDateVarList & "] = 'Not Found'" _
            & "WHERE " _
            & "(((" & CurrentTable & ".[" & myDateVarList & "])='POST FCT Trial') " _
            & "AND (([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
            & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))='01') " _
            & "AND ((" & CurrentTable & ".Event)='FCT'));"
            ' Debug.Print "done with column table: " & myDateVarList & " For FCT Event Late Adds."
        Next myDateVarList
            ' Debug.Print vbCrLf & "Completed re-setting prior to FCT Event Screenings to Not Found."

    ElseIf All_or_Events = "Events" Then
        'Run for BT Event
        For Each myDateVarList In trialsOnlyList
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
        For Each myDateVarList In trialsOnlyList
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
        
        myDateVarList = Empty
        
        'Run for FCT Event _
        This is marking new FCT trial cards as POST FCT Trial, Incorrect marking _
        the fix for this is to use a list that only starts at the FCT Event
        For Each myDateVarList In trialsOnlyList
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
        
        myDateVarList = Empty
        
    Else
        'Un trapped error
        'All_or_Events Global is empty or not expected value
        Debug.Print "Function SetLateAdds_TrialCards_CS() was passed empty or not expected value with GLOBAL All_or_Events:= " & All_or_Events & "."
    End If
    
    myDateVarList = Empty

    ' Debug.Print vbCrLf & "The Late Add Trial Cards Update Query completed." & vbCrLf

End Sub


Public Sub SetClosedXX_TrialCards_CS()
'This changes the screening from <Not Found> to <X/X> _
where the trial card was reidentified at a later Event.

Dim myDateVarList As Variant

    'Check for All Reports or only Events
    If All_or_Events = "All" Then
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
    ElseIf All_or_Events = "Events" Then
        'Run for all the X/X cards where screen is Not Found
        For Each myDateVarList In trialsOnlyList
            CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
            & "SET " _
            & "" & CurrentTable & ".[" & myDateVarList & "] = 'X/X'" _
            & "WHERE " _
            & "(((" & CurrentTable & ".[" & myDateVarList & "])='Not Found') " _
            & "AND (([" & CurrentTable & "]![Final_Sts_A_T])='X/X'));"
            ' Debug.Print "done with column table: " & myDateVarList & " For X/X Closures."
        Next myDateVarList
    Else
        'Un trapped error
        'All_or_Events Global is empty or not expected value
        Debug.Print "Function SetClosedXX_TrialCards_CS() was passed empty or not expected value with GLOBAL All_or_Events:= " & All_or_Events & "."
    End If
    
    myDateVarList = Empty
    
    ' Debug.Print vbCrLf & "The X/X Trial Cards Update Query completed." & vbCrLf

End Sub


Public Sub SetTrialCardSplits_CS()
'This changes the screening from <Not Found> to <SPLIT> _
where the trial card was split after the actual inspection.

Dim myDateVarList As Variant

    'Check for All Reports or only Events
    If All_or_Events = "All" Then
        'Run for BT Event
        For Each myDateVarList In allColumnsList
            If allColumnsList.IndexOf(myDateVarList, 0) = allColumnsList_BT Then
                'Re-Set First Screen to Split AND Event date column to Split
                CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
                & "SET " & CurrentTable & ".First_Screening = 'SPLIT', " _
                & CurrentTable & ".[" & myDateVarList & "] = 'SPLIT'" _
                & "WHERE ((([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
                & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))<>'01') " _
                & "AND ((" & CurrentTable & ".Event)='BT'));"
            ElseIf allColumnsList.IndexOf(myDateVarList, 0) < allColumnsList_BT Then
                'Do Nothing, leave marked as "Not Found"
            Else
                CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
                & "SET " & CurrentTable & ".[" & myDateVarList & "] = 'SPLIT'" _
                & "WHERE (((" & CurrentTable & ".[" & myDateVarList & "])='Not Found') " _
                & "AND (([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
                & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))<>'01') " _
                & "AND ((" & CurrentTable & ".Event)='BT'));"
            End If
            ' Debug.Print "done with column table: " & myDateVarList & " For BT Event Splits."
        Next myDateVarList
        
        myDateVarList = Empty
        
        'Run for AT Event
        For Each myDateVarList In allColumnsList
            If allColumnsList.IndexOf(myDateVarList, 0) = allColumnsList_AT Then
                'Re-Set First Screen to Split AND Event date column to Split
                CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
                & "SET " & CurrentTable & ".First_Screening = 'SPLIT', " _
                & CurrentTable & ".[" & myDateVarList & "] = 'SPLIT'" _
                & "WHERE ((([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
                & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))<>'01') " _
                & "AND ((" & CurrentTable & ".Event)='AT'));"
            ElseIf allColumnsList.IndexOf(myDateVarList, 0) < allColumnsList_AT Then
                'Do Nothing, leave marked as "Not Found"
            Else
                CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
                & "SET " & CurrentTable & ".[" & myDateVarList & "] = 'SPLIT'" _
                & "WHERE (((" & CurrentTable & ".[" & myDateVarList & "])='Not Found') " _
                & "AND (([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
                & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))<>'01') " _
                & "AND ((" & CurrentTable & ".Event)='AT'));"
            End If
            ' Debug.Print "done with column table: " & myDateVarList & " For AT Event Splits."
        Next myDateVarList
        
        myDateVarList = Empty
        
        'Run for FCT Event
        For Each myDateVarList In allColumnsList
            If allColumnsList.IndexOf(myDateVarList, 0) = allColumnsList_FCT Then
                'Re-Set First Screen to Split AND Event date column to Split
                CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
                & "SET " & CurrentTable & ".First_Screening = 'SPLIT', " _
                & CurrentTable & ".[" & myDateVarList & "] = 'SPLIT'" _
                & "WHERE ((([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
                & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))<>'01') " _
                & "AND ((" & CurrentTable & ".Event)='FCT'));"
            ElseIf allColumnsList.IndexOf(myDateVarList, 0) < allColumnsList_FCT Then
                'Do Nothing, leave marked as "Not Found"
            Else
                CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
                & "SET " & CurrentTable & ".[" & myDateVarList & "] = 'SPLIT'" _
                & "WHERE (((" & CurrentTable & ".[" & myDateVarList & "])='Not Found') " _
                & "AND (([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
                & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))<>'01') " _
                & "AND ((" & CurrentTable & ".Event)='FCT'));"
            End If
            ' Debug.Print "done with column table: " & myDateVarList & " For FCT Event Splits."
        Next myDateVarList
    
    ElseIf All_or_Events = "Events" Then
        'Run for BT Event
        For Each myDateVarList In trialsOnlyList
            If trialsOnlyList.IndexOf(myDateVarList, 0) = trialsOnlyList_BT Then
                'Re-Set First Screen to Split AND Event date column to Split
                CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
                & "SET " & CurrentTable & ".First_Screening = 'SPLIT', " _
                & CurrentTable & ".[" & myDateVarList & "] = 'SPLIT'" _
                & "WHERE ((([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
                & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))<>'01') " _
                & "AND ((" & CurrentTable & ".Event)='BT'));"
            ElseIf trialsOnlyList.IndexOf(myDateVarList, 0) < trialsOnlyList_BT Then
                'Do Nothing, leave marked as "Not Found"
            Else
                CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
                & "SET " & CurrentTable & ".[" & myDateVarList & "] = 'SPLIT'" _
                & "WHERE (((" & CurrentTable & ".[" & myDateVarList & "])='Not Found') " _
                & "AND (([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
                & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))<>'01') " _
                & "AND ((" & CurrentTable & ".Event)='BT'));"
            End If
            ' Debug.Print "done with column table: " & myDateVarList & " For BT Event Splits."
        Next myDateVarList
        
        myDateVarList = Empty
        
        'Run for AT Event
        For Each myDateVarList In trialsOnlyList
            If trialsOnlyList.IndexOf(myDateVarList, 0) = trialsOnlyList_AT Then
                'Re-Set First Screen to Split AND Event date column to Split
                CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
                & "SET " & CurrentTable & ".First_Screening = 'SPLIT', " _
                & CurrentTable & ".[" & myDateVarList & "] = 'SPLIT'" _
                & "WHERE ((([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
                & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))<>'01') " _
                & "AND ((" & CurrentTable & ".Event)='AT'));"
            ElseIf trialsOnlyList.IndexOf(myDateVarList, 0) < trialsOnlyList_AT Then
                'Do Nothing, leave marked as "Not Found"
            Else
                CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
                & "SET " & CurrentTable & ".[" & myDateVarList & "] = 'SPLIT'" _
                & "WHERE (((" & CurrentTable & ".[" & myDateVarList & "])='Not Found') " _
                & "AND (([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
                & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))<>'01') " _
                & "AND ((" & CurrentTable & ".Event)='AT'));"
            End If
            ' Debug.Print "done with column table: " & myDateVarList & " For AT Event Splits."
        Next myDateVarList
        
        myDateVarList = Empty
        
        'Run for FCT Event
        For Each myDateVarList In trialsOnlyList
            If trialsOnlyList.IndexOf(myDateVarList, 0) = trialsOnlyList_FCT Then
                'Re-Set First Screen to Split AND Event date column to Split
                CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
                & "SET " & CurrentTable & ".First_Screening = 'SPLIT', " _
                & CurrentTable & ".[" & myDateVarList & "] = 'SPLIT'" _
                & "WHERE ((([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
                & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))<>'01') " _
                & "AND ((" & CurrentTable & ".Event)='FCT'));"
            ElseIf trialsOnlyList.IndexOf(myDateVarList, 0) < trialsOnlyList_FCT Then
                'Do Nothing, leave marked as "Not Found"
            Else
                CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
                & "SET " & CurrentTable & ".[" & myDateVarList & "] = 'SPLIT'" _
                & "WHERE (((" & CurrentTable & ".[" & myDateVarList & "])='Not Found') " _
                & "AND (([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
                & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))<>'01') " _
                & "AND ((" & CurrentTable & ".Event)='FCT'));"
            End If
            ' Debug.Print "done with column table: " & myDateVarList & " For FCT Event Splits."
        Next myDateVarList
    
    Else
        'Un trapped error
        'All_or_Events Global is empty or not expected value
        Debug.Print "Function SetTrialCardSplits_CS() was passed empty or not expected value with GLOBAL All_or_Events:= " & All_or_Events & "."
    End If
    
    myDateVarList = Empty
    
'    'Check for All Reports or only Events
'    If All_or_Events = "All" Then
'        '#Fix BUG here by going back and resetting the [TABLE].[First_Screening] to "SPLIT"
'        CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
'        & "SET " & CurrentTable & ".First_Screening = 'SPLIT'" _
'        & "WHERE (((" & CurrentTable & ".[" & allColumnsList(0) & "])='SPLIT') " _
'        & "AND (([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
'        & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))<>'01'));"
'    ElseIf All_or_Events = "Events" Then
'        '#Fix BUG here by going back and resetting the [TABLE].[First_Screening] to "SPLIT"
'        CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
'        & "SET " & CurrentTable & ".First_Screening = 'SPLIT'" _
'        & "WHERE (((" & CurrentTable & ".[" & trialsOnlyList(0) & "])='SPLIT') " _
'        & "AND (([" & CurrentTable & "]![Final_Sts_A_T])<>'X/X') " _
'        & "AND ((Right([" & CurrentTable & "]![Trial_Card],2))<>'01'));"
'    Else
'        'Un trapped error
'        'All_or_Events Global is empty or not expected value
'        Debug.Print "Function SetTrialCardSplits_CS() was passed empty or not expected value with GLOBAL All_or_Events:= " & All_or_Events & "."
'    End If
    
    ' Debug.Print vbCrLf & "The Split Trial Cards Update Query completed." & vbCrLf

End Sub


Public Sub SetTrialCardsMissingFromReports_CS(startEvent As Long, endEvent As Long)
'This changes the screening from <Not Found> to the prior value _
where the trial card was missing from a report source.
Dim rsWriteData As DAO.Recordset 'This is the table recordset
Dim i As Long ' Used as the allColumnsList(i) itterable
'Dim myTableName As String
'Dim startEvent As Long
'Dim endEvent As Long
Dim myTC_Number As String
Dim mySparseVar As String
Dim DateColumnVal As String 'Current Column header date value string, CurrrentColumn(0)
Dim PriorDateColumnVal As String 'Prior Column header date value string, CurrrentColumn(-1)
Dim Neg2DateColumnVal As String 'Neg 2 Column header date value string, CurrrentColumn(-2)
Dim Neg3DateColumnVal As String 'Neg 3 Column header date value string, CurrrentColumn(-3)
Dim CurState As String ' This is the current screening being evaluated
Dim PriorState As String ' CurState minus 1
Dim Neg2State As String ' CurState minus 2
Dim Neg3State As String ' CurState minus 3

'Call the SetListAndVars Module
'AddingToMyDateLists 'Run the list builder
'Call the SetListAndVars_Summary Module
'AddingToMySummaryDateLists 'Run the list builder

'myTableName = "All_Combined_Screenings"
'startEvent = allColumnsList_AT
'endEvent = allColumnsList_FCT

'CurrentDb.OpenRecordset(Name:="All_Z_Summary", Type:=dbOpenDynaset, Options:=dbInconsistent, LockEdit:=dbOptimistic)
Set rsWriteData = CurrentDb.OpenRecordset(CurrentTable, dbOpenDynaset, dbInconsistent, dbOptimistic)
'Move cursor to first row, this will be used to itterate through all the rows in order
rsWriteData.MoveFirst
'myTC_Number = rsWriteData.Fields(Trial_Card)
'Debug.Print ("myTC_Number: " & myTC_Number)
    

    While (Not rsWriteData.EOF)
        myTC_Number = rsWriteData.Fields("Trial_Card")
        'Debug.Print ("myTC_Number: " & myTC_Number)

       'Run for BT Event to the AT Event, but not the AT Event
       'Run for AT Event to the FCT Event, but not the FCT Event
       'Run for FCT Event to the OWLD Event, but not the OWLD Event
        For i = startEvent To (endEvent - 1)
            'Debug.Print "i = " & i
            DateColumnVal = allColumnsList(i)
            'Debug.Print "DateColumnVal = " & DateColumnVal
            
            If i = startEvent Then
                'This is the first date column in the current range
                'and it is the Event Column
                'i = 0
                CurState = rsWriteData(DateColumnVal)
                mySparseVar = CurState
            ElseIf i = (startEvent + 1) Then
                'This is the second date column in the current range
                'and it has the Event Column
                'i = 1
                PriorDateColumnVal = allColumnsList(i - 1)
                CurState = rsWriteData(DateColumnVal)
                PriorState = rsWriteData(PriorDateColumnVal)
                If CurState = "Not Found" Then
                    'TC Value is missing in current report
                    'Need to test for missing TC's from report to report
                    If PriorState = "Not Found" Then
                        'I am only looking back one reporting window
                        'TC Value is missing two reports in a row
                        mySparseVar = "Not Found"
                    ElseIf PriorState <> "Not Found" Then
                        'The TC is not missing from the prior report
                        mySparseVar = PriorState
                    Else
                        'Untrapped error
                        mySparseVar = "991A" ' This is a flag value that something is not correct
                    End If
                ElseIf CurState <> "Not Found" Then
                        'The TC is not missing from the current report
                        mySparseVar = CurState
                Else
                    'Untrapped error
                    mySparseVar = "991B" ' This is a flag value that something is not correct
                End If
            ElseIf i = (startEvent + 2) Then
                'This is the third date column in the current range
                'and it has the Event Column where i = 2
                'i > 1
                PriorDateColumnVal = allColumnsList(i - 1)
                Neg2DateColumnVal = allColumnsList(i - 2)
                CurState = rsWriteData(DateColumnVal)
                PriorState = rsWriteData(PriorDateColumnVal)
                Neg2State = rsWriteData(Neg2DateColumnVal)
                If CurState = "Not Found" Then
                    'TC Value is missing in current report
                    'Need to test for missing TC's from report to report
                    If PriorState = "Not Found" Then
                        'TC Value is missing two reports in a row
                        If Neg2State = "Not Found" Then
                            'I am only looking back two reporting windows
                            'TC Value is missing three reports in a row
                            mySparseVar = "Not Found"
                        Else
                            'Untrapped error
                            mySparseVar = "992A" ' This is a flag value that something is not correct
                        End If
                    ElseIf PriorState <> "Not Found" Then
                        'The TC is not missing from the prior report
                        mySparseVar = PriorState
                    Else
                        'Untrapped error
                        mySparseVar = "992B" ' This is a flag value that something is not correct
                    End If
                ElseIf CurState <> "Not Found" Then
                        'The TC is not missing from the current report
                        mySparseVar = CurState
                Else
                    'Untrapped error
                    mySparseVar = "992C" ' This is a flag value that something is not correct
                End If
            ElseIf i > (startEvent + 2) Then
                'This is the Fourth date column in the current range
                'and it has the Event Column where i = 3
                'i > 2
                PriorDateColumnVal = allColumnsList(i - 1)
                Neg2DateColumnVal = allColumnsList(i - 2)
                Neg3DateColumnVal = allColumnsList(i - 3)
                CurState = rsWriteData(DateColumnVal)
                PriorState = rsWriteData(PriorDateColumnVal)
                Neg2State = rsWriteData(Neg2DateColumnVal)
                Neg3State = rsWriteData(Neg3DateColumnVal)
                If CurState = "Not Found" Then
                    'TC Value is missing in current report
                    'Need to test for missing TC's from report to report
                    If PriorState = "Not Found" Then
                        'TC Value is missing two reports in a row
                        If Neg2State = "Not Found" Then
                            'TC Value is missing three reports in a row
                            If Neg3State = "Not Found" Then
                                'TC Value is missing four reports in a row
                                'I am only looking back three reporting windows
                                mySparseVar = "Not Found"
                            ElseIf Neg3State <> "Not Found" Then
                                'The TC is not missing from the Neg3State report
                                mySparseVar = Neg3State
                            Else
                                'Untrapped error
                                mySparseVar = "993A" ' This is a flag value that something is not correct
                            End If
                        ElseIf Neg2State <> "Not Found" Then
                            'The TC is not missing from the Neg2State report
                            mySparseVar = Neg2State
                        Else
                            'Untrapped error
                            mySparseVar = "993B" ' This is a flag value that something is not correct
                        End If
                    ElseIf PriorState <> "Not Found" Then
                        'The TC is not missing from the prior report
                        mySparseVar = PriorState
                    Else
                        'Untrapped error
                        mySparseVar = "993C" ' This is a flag value that something is not correct
                    End If
                ElseIf CurState <> "Not Found" Then
                        'The TC is not missing from the current report
                        mySparseVar = CurState
                Else
                    'Untrapped error
                    mySparseVar = "993D" ' This is a flag value that something is not correct
                End If
            Else
                'Untrapped error
            End If
            
            'Debug.Print "CurState = " & CurState
            'Debug.Print "mySparseVar = " & mySparseVar
            
            If mySparseVar = CurState Then
                'Do Nothing, the value is already in the table
            Else
                'Run the Update SQL
                'CurrentDb.Execute "UPDATE DISTINCTROW " & myTableName & " " _
                & "SET " & myTableName & ".[" & DateColumnVal & "] = '" & mySparseVar & "' " _
                & "WHERE ((" & myTableName & ".[Trial_Card])='" & myTC_Number & "');"
                
                rsWriteData.Edit
                rsWriteData(DateColumnVal) = mySparseVar
                rsWriteData.Update
                
            End If
            
            
            PriorDateColumnVal = Empty
            Neg2DateColumnVal = Empty
            Neg3DateColumnVal = Empty
            mySparseVar = Empty
            CurState = Empty
            PriorState = Empty
            Neg2State = Empty
            Neg3State = Empty
            
         Next i
         
         rsWriteData.MoveNext
    Wend
    
rsWriteData.Close
Set rsWriteData = Nothing

    ' Debug.Print vbCrLf & "The Split Trial Cards Update Query completed." & vbCrLf

End Sub
