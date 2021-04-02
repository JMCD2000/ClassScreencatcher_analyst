Option Compare Database
'The use of NotFound until the trial card is actually written at Event just looks messy, would like to change this
Option Explicit


Public Sub SetFirstScreensAndEvents_XXSO()
'This is to load or reload the screenings into the Screen_Only Table

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
        Debug.Print "Function SetFirstScreensAndEvents_XXSO() was passed empty or not expected value with GLOBAL All_or_Events:= " & All_or_Events & "."
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
    
    ' Debug.Print vbCrLf & "Completed the" & CurrentTable & "table data set with place holder values Update Query."

    'Set Final Screening as Final and as First
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFinal & "] ON " & CurrentTable & ".Trial_Card = [" & beanFinal & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".First_Screening = [" & beanFinal & "].[Screening], " _
    & "" & CurrentTable & ".[" & columnFinal & "] = [" & beanFinal & "].[Screening], " _
    & "" & CurrentTable & ".Final_Sts_A_T = [" & beanFinal & "].[Final_Sts_A_T], " _
    & "" & CurrentTable & ".Trial_ID = [" & beanFinal & "].[Trial_ID], " _
    & "" & CurrentTable & ".Event = [" & beanFinal & "].[Event];"
    ' Debug.Print vbCrLf & "Completed setting values in columns Trial_ID, Event, Final_Sts_A_T, First_Screening."

    'Set FCT Event Screening and First Screening, FCT upload
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanFCT & "] ON " & CurrentTable & ".Trial_Card = [" & beanFCT & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".First_Screening = [" & beanFCT & "].[Screening], " _
    & "" & CurrentTable & ".[" & columnFCT & "] = [" & beanFCT & "].[Screening];"
    ' Debug.Print vbCrLf & "Completed setting FCT Event Screening and First Screening"

    'Set AT Event Screening and First Screening, AT upload
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanAT & "] ON " & CurrentTable & ".Trial_Card = [" & beanAT & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".First_Screening = [" & beanAT & "].[Screening], " _
    & "" & CurrentTable & ".[" & columnAT & "] = [" & beanAT & "].[Screening];"
    ' Debug.Print vbCrLf & "Completed setting AT Event Screening and First Screening"

    'Set BT Event Screening and First Screening, BT upload
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanBT & "] ON " & CurrentTable & ".Trial_Card = [" & beanBT & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".First_Screening = [" & beanBT & "].[Screening], " _
    & "" & CurrentTable & ".[" & columnBT & "] = [" & beanBT & "].[Screening];"
    ' Debug.Print vbCrLf & "Completed setting BT Event Screening and First Screening"
    
    'Set OWLD Event Screening, Transfer Book Bean Data
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanOWLD & "] ON " & CurrentTable & ".Trial_Card = [" & beanOWLD & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".[" & columnOWLD & "] = [" & beanOWLD & "].[Screening];"
    ' Debug.Print vbCrLf & "Completed setting OWLD Event Screening"

    'Set DEL Event Screening, DEL Milestone
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & beanDEL & "] ON " & CurrentTable & ".Trial_Card = [" & beanDEL & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".[" & columnDEL & "] = [" & beanDEL & "].[Screening];"
    ' Debug.Print vbCrLf & "Completed setting DEL Event Screening
    
    ' Debug.Print vbCrLf & "The Trials First Screenings and OWLD Update Query completed." & vbCrLf

End Sub


Public Sub SetNonShipEventScrns_XXSO()
'This loads the screens that are not actual trials
'This is not needed for the Events only tables

Dim myDateVarList As Variant

    For Each myDateVarList In nonTrialsList
        CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " INNER JOIN [" & myDateVarList & "_" & curHullNum & "] ON " & CurrentTable & ".Trial_Card = [" & myDateVarList & "_" & curHullNum & "].Trial_Card " _
        & "SET " _
        & "" & CurrentTable & ".[" & myDateVarList & "] = [" & myDateVarList & "_" & curHullNum & "].[Screening];"
        ' Debug.Print "done with column table: " & myDateVarList & "."
    Next myDateVarList
    
    myDateVarList = Empty
    
    ' Debug.Print vbCrLf & "Completed the Non Trial Screenings Update Query." & vbCrLf

End Sub


Public Sub SetLateAdds_TrialCards_XXSO()
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
        there are list populated dates in this SQL Query!
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
        Debug.Print "Function SetLateAdds_TrialCards_XXSO() was passed empty or not expected value with GLOBAL All_or_Events:= " & All_or_Events & "."
    End If
    
    myDateVarList = Empty
    
    ' Debug.Print vbCrLf & "The Late Add Trial Cards Update Query completed." & vbCrLf

End Sub


Public Sub SetClosedXX_TrialCards_XXSO()
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
        Debug.Print "Function SetClosedXX_TrialCards_XXSO() was passed empty or not expected value with GLOBAL All_or_Events:= " & All_or_Events & "."
    End If
            
    myDateVarList = Empty
        
    ' Debug.Print vbCrLf & "The X/X Trial Cards Update Query completed." & vbCrLf

End Sub


Public Sub SetTrialCardSplits_XXSO()
'This changes the screening from <Not Found> to <SPLIT> _
where the trial card was entered after the actual inspection.

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
        Debug.Print "Function SetTrialCardSplits_XXSO() was passed empty or not expected value with GLOBAL All_or_Events:= " & All_or_Events & "."
    End If
    
    myDateVarList = Empty
    
    'Check for All Reports or only Events
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
'        Debug.Print "Function SetTrialCardSplits_XXSO() was passed empty or not expected value with GLOBAL All_or_Events:= " & All_or_Events & "."
'    End If
    
    ' Debug.Print vbCrLf & "The Split Trial Cards Update Query completed." & vbCrLf

End Sub
