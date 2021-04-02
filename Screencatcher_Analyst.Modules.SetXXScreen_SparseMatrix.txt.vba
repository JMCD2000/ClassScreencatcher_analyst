Option Compare Database
'This has issues with the way that splits are handeled and the way that cards are handeled prior to the Event
'The use of SPLIT until the actual split occures causes counting issues because the screening changes from SPLIT to the actual screening
'The use of NotFound until the trial card is actually written at Event just looks messy, would like to change this
Option Explicit


Public Sub SetFirstScreensAndEvents_XXSM()
'This is to load or reload the screening sparse matrix into the <TABLE>_SparseMatrix

'Set data columns to <Not Found>
Dim myDateVarList As Variant
Dim notFound As String 'The <Not Found> is not used in TSM or elsewhere, becomes a visual that something was missed
Dim notCounted As Integer 'The 999 is not an expected count for rescreens and becomes a visual that something was missed
Dim emptyID As String 'The dash is not used in TSM or elsewhere, becomes a visual that something was missed
Dim emptyEvent As String 'The double E is not used in TSM or elsewhere, becomes a visual that something was missed
Dim emptySts_A_T As String 'The dash slash dash is not used in TSM or elsewhere, becomes a visual that something was missed

    'Set Trial ID, Event and Final Status values as place holder values
    notFound = "Not Found"
    notCounted = 999
    emptyID = "-"
    emptyEvent = "EE"
    emptySts_A_T = "-/-"

    'Check for All Reports or only Events
    If All_or_Events = "All" Then
        For Each myDateVarList In allColumnsList
            CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & SparseRefTable & "] ON " & CurrentTable & ".Trial_Card = [" & SparseRefTable & "].Trial_Card " _
            & "SET " _
            & "" & CurrentTable & ".[" & myDateVarList & "] = '" & notFound & "';"
            ' Debug.Print "done with column table: " & myDateVarList & "."
        Next myDateVarList
    ElseIf All_or_Events = "Events" Then
        For Each myDateVarList In trialsOnlyList
            CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & SparseRefTable & "] ON " & CurrentTable & ".Trial_Card = [" & SparseRefTable & "].Trial_Card " _
            & "SET " _
            & "" & CurrentTable & ".[" & myDateVarList & "] = '" & notFound & "';"
            ' Debug.Print "done with column table: " & myDateVarList & "."
        Next myDateVarList
    Else
        'Un trapped error
        'All_or_Events Global is empty or not expected value
        Debug.Print "Function SetFirstScreensAndEvents_XXSM() was passed empty or not expected value with GLOBAL All_or_Events:= " & All_or_Events & "."
    End If
        
    myDateVarList = Empty
    
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & SparseRefTable & "] ON " & CurrentTable & ".Trial_Card = [" & SparseRefTable & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".Trial_ID = '" & emptyID & "', " _
    & "" & CurrentTable & ".Event = '" & emptyEvent & "', " _
    & "" & CurrentTable & ".Final_Sts_A_T = '" & emptySts_A_T & "', " _
    & "" & CurrentTable & ".Rescreen_counts_BT_to_DEL = '" & notCounted & "', " _
    & "" & CurrentTable & ".Rescreen_counts_AT_to_DEL = '" & notCounted & "', " _
    & "" & CurrentTable & ".Rescreen_counts_BT_to_OWLD = '" & notCounted & "', " _
    & "" & CurrentTable & ".Rescreen_counts_AT_to_OWLD = '" & notCounted & "', " _
    & "" & CurrentTable & ".Rescreen_counts_FCT_to_OWLD = '" & notCounted & "', " _
    & "" & CurrentTable & ".Rescreen_counts_BT_to_Final = '" & notCounted & "', " _
    & "" & CurrentTable & ".Rescreen_counts_AT_to_Final = '" & notCounted & "', " _
    & "" & CurrentTable & ".Rescreen_counts_FCT_to_Final = '" & notCounted & "';"
    ' Debug.Print "Completed the" & CurrentTable & "table data set with place holder values Update Query."
    
    'Set Trial ID, Event and Final Status values
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & SparseRefTable & "] ON " & CurrentTable & ".Trial_Card = [" & SparseRefTable & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".Trial_ID = [" & SparseRefTable & "].[Trial_ID], " _
    & "" & CurrentTable & ".Event = [" & SparseRefTable & "].[Event], " _
    & "" & CurrentTable & ".Final_Sts_A_T = [" & SparseRefTable & "].[Final_Sts_A_T];"
    ' Debug.Print "Completed setting values in columns Trial_ID, Event, Final_Sts_A_T."
    
    'Debug.Print vbCrLf & "The " & CurrentTable & " Table is now ready for Sparse Matrix Calculations." & vbCrLf

End Sub


Public Sub Build_and_Set_Aggregated_Screen_XXSM()
'This builds the sparse matrix that has a 0 if no change over the prior or has a 1 if there was a change.
Dim myDateVarList As Variant ' Used to cycle through each date column as an iterable
Dim myIndex As Integer ' This is to get the prior screen
'These are used on the recordset
Dim priorScreenVar As String ' This holds the first screen from the date columns
Dim currentScreenVar As String ' This holds the last column looked at value to see if current column is different
Dim mySparseVar As String ' This holds the
Dim aggregatedScreensVar As String ' This holds the screens as they are collected
Dim curTrial_Card As String ' This is for the update back to the TC_Screen_Agg Table
'Open a recordset to loop through each Record in the recordset
Dim dbs_Read As DAO.Database
Dim rst As DAO.Recordset ' or Recordset2?
Dim mySQLstring As String
'Open a recordset to update the TC_Screen_Agg table
Dim dbs_Write As DAO.Database
    
    'simple varible referenced typed SQL statement
    mySQLstring = "SELECT " & SparseRefTable & ".* FROM " & SparseRefTable & " INNER JOIN " & CurrentTable & " ON " & SparseRefTable & ".Trial_Card = " & CurrentTable & ".Trial_Card;"

    'Open a pointer to current database
    Set dbs_Read = CurrentDb()
    Set dbs_Write = CurrentDb()
    'Create the recordset with my SQL string
    Set rst = dbs_Read.OpenRecordset(mySQLstring)

    'Check for All Reports or only Events
    If All_or_Events = "All" Then
        Do While Not rst.EOF
            'Step down each record row
            myIndex = 0 'This is the first date column in the allColumnsList, reset to 0 on each row
            curTrial_Card = rst![Trial_Card]
            'Debug.Print ("Cur TC: " & curTrial_Card)
            
            For Each myDateVarList In allColumnsList
                'Step accross each date column within the current record row
                'Debug.Print ("TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                'Default value for the sparse matrix reset at each column itteration
                mySparseVar = Empty
                        
                'Assign the prior screen from the first date column
                If myDateVarList = allColumnsList(0) Then
                    'I am on the first date column, set prior screen = [First_Screening]
                    priorScreenVar = rst![First_Screening]
                    currentScreenVar = rst("" & myDateVarList & "") ' rst![ & myDateVarList & ]
'                    If priorScreenVar = "Not Found" _
'                        Or priorScreenVar = "POST BT Trial" _
'                        Or priorScreenVar = "POST AT Trial" _
'                        Or priorScreenVar = "POST FCT Trial" _
'                        Or priorScreenVar = "SPLIT" _
'                        Or priorScreenVar = "X/X" _
'                        Or priorScreenVar = "" _
'                        Or priorScreenVar = Empty _
'                        Then
                    If priorScreenVar = "Not Found" _
                        Or priorScreenVar = "POST BT Trial" _
                        Or priorScreenVar = "POST AT Trial" _
                        Or priorScreenVar = "POST FCT Trial" _
                        Or priorScreenVar = "X/X" _
                        Or priorScreenVar = "" _
                        Or priorScreenVar = Empty _
                        Then
                        'Do not add this to the aggragate screening
                        'Pass
                    Else
                        aggregatedScreensVar = rst![First_Screening]
                    End If
                    'Leave myIndex set to zero here so it is lagging next time arround
                    'Debug.Print ("First date column of record row: " & myDateVarList)
                ElseIf myDateVarList = allColumnsList.Item(allColumnsList.Count - 1) Then
                    'I am at the end of the date columns
                    priorScreenVar = rst("" & allColumnsList(myIndex) & "")
                    currentScreenVar = rst("" & myDateVarList & "")
                    'Do nothing to myIndex because this is the last column
                    'Debug.Print ("Last date column of record row: " & myDateVarList)
                Else
                    'I am in the middle of the date columns, set priorScreenVar=x-1 currentScreenVar=x
                    priorScreenVar = rst("" & allColumnsList(myIndex) & "")
                    currentScreenVar = rst("" & myDateVarList & "")
                    'Debug.Print ("Values being compaired, Prior : Current " & priorScreenVar & " : " & currentScreenVar)
                    'Increment the myIndex counter +1
                    myIndex = myIndex + 1
                End If
            
                'Debug.Print ("currentScreenVar =" & currentScreenVar & "=")
                'Debug.Print ("priorScreenVar =" & priorScreenVar & "=")
                    
'                If currentScreenVar = "Not Found" _
'                    Or currentScreenVar = "POST BT Trial" _
'                    Or currentScreenVar = "POST AT Trial" _
'                    Or currentScreenVar = "POST FCT Trial" _
'                    Or currentScreenVar = "SPLIT" _
'                    Or currentScreenVar = "X/X" _
'                    Or currentScreenVar = "" _
'                    Or currentScreenVar = Empty _
'                    Then
                If currentScreenVar = "Not Found" _
                    Or currentScreenVar = "POST BT Trial" _
                    Or currentScreenVar = "POST AT Trial" _
                    Or currentScreenVar = "POST FCT Trial" _
                    Or currentScreenVar = "X/X" _
                    Or currentScreenVar = "" _
                    Or currentScreenVar = Empty _
                    Then
                    'Do not compare current screening against prior screening
                    mySparseVar = "~~~"
                ElseIf currentScreenVar = priorScreenVar Then
                    'The screening has not changed, do not mark this screening
                    'Debug.Print ("currentScreenVar = priorScreenVar : " & currentScreenVar & " EQUALS " & priorScreenVar)
                    mySparseVar = "~~~"
                ElseIf currentScreenVar <> priorScreenVar Then
                    'The screening has changed and mark this screening
                    'Debug.Print ("currentScreenVar = priorScreenVar : " & currentScreenVar & " EQUALS " & priorScreenVar)
'                    If priorScreenVar = "Not Found" _
'                        Or priorScreenVar = "POST BT Trial" _
'                        Or priorScreenVar = "POST AT Trial" _
'                        Or priorScreenVar = "POST FCT Trial" _
'                        Or priorScreenVar = "SPLIT" _
'                        Or priorScreenVar = "X/X" _
'                        Or priorScreenVar = "" _
'                        Or priorScreenVar = Empty _
'                        Then
                    If priorScreenVar = "Not Found" _
                        Or priorScreenVar = "POST BT Trial" _
                        Or priorScreenVar = "POST AT Trial" _
                        Or priorScreenVar = "POST FCT Trial" _
                        Or priorScreenVar = "X/X" _
                        Or priorScreenVar = "" _
                        Or priorScreenVar = Empty _
                        Then
                        mySparseVar = currentScreenVar
                    Else
                        mySparseVar = priorScreenVar & "/" & currentScreenVar
                    End If
                Else
                    mySparseVar = 99 ' This is a flag value that something is not correct
                End If
                
                'Debug.Print ("UPDATE DISTINCTROW " & CurrentTable & " " _
                & "SET " _
                & "" & CurrentTable & ".[" & myDateVarList & "] = '" & mySparseVar & "' " _
                & "WHERE " _
                & "" & CurrentTable & ".Trial_Card = '" & curTrial_Card & "';")
                
                'Debug.Print ("Cur TC: " & curTrial_Card)
                'Update the Sparse Matrix table with the Value of 0 or 1
'                CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " " _
'                & "SET " _
'                & "" & CurrentTable & ".[" & myDateVarList & "] = '" & mySparseVar & "' " _
'                & "WHERE " _
'                & "" & CurrentTable & ".Trial_Card = '" & curTrial_Card & "';"
                
                dbs_Write.Execute "UPDATE DISTINCTROW " & CurrentTable & " " _
                & "SET " _
                & "" & CurrentTable & ".[" & myDateVarList & "] = '" & mySparseVar & "' " _
                & "WHERE " _
                & "" & CurrentTable & ".Trial_Card = '" & curTrial_Card & "';"
                
                
            Next myDateVarList
            
            'When last column of row record is reached control is returned to this loop
         
            myDateVarList = Empty
            'Debug.Print ("Next row in Table")
            
            rst.MoveNext
            
        Loop

    ElseIf All_or_Events = "Events" Then
        Do While Not rst.EOF
            'Step down each record row
            myIndex = 0 'This is the first date column in the allColumnsList, reset to 0 on each row
            curTrial_Card = rst![Trial_Card]
            'Debug.Print ("Cur TC: " & curTrial_Card)
            
            For Each myDateVarList In trialsOnlyList
                'Step accross each date column within the current record row
                'Debug.Print ("TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                'Default value for the sparse matrix reset at each column itteration
                mySparseVar = Empty
                        
                'Assign the prior screen from the first date column
                If myDateVarList = trialsOnlyList(0) Then
                    'I am on the first date column, set prior screen = [First_Screening]
                    priorScreenVar = rst![First_Screening]
                    currentScreenVar = rst("" & myDateVarList & "") ' rst![ & myDateVarList & ]
'                    If priorScreenVar = "Not Found" _
'                        Or priorScreenVar = "POST BT Trial" _
'                        Or priorScreenVar = "POST AT Trial" _
'                        Or priorScreenVar = "POST FCT Trial" _
'                        Or priorScreenVar = "SPLIT" _
'                        Or priorScreenVar = "X/X" _
'                        Or priorScreenVar = "" _
'                        Or priorScreenVar = Empty _
'                        Then
                    If priorScreenVar = "Not Found" _
                        Or priorScreenVar = "POST BT Trial" _
                        Or priorScreenVar = "POST AT Trial" _
                        Or priorScreenVar = "POST FCT Trial" _
                        Or priorScreenVar = "X/X" _
                        Or priorScreenVar = "" _
                        Or priorScreenVar = Empty _
                        Then
                        'Do not add this to the aggragate screening
                        'Pass
                    Else
                        aggregatedScreensVar = rst![First_Screening]
                    End If
                    'Leave myIndex set to zero here so it is lagging next time arround
                    'Debug.Print ("First date column of record row: " & myDateVarList)
                ElseIf myDateVarList = trialsOnlyList.Item(trialsOnlyList.Count - 1) Then
                    'I am at the end of the date columns
                    priorScreenVar = rst("" & trialsOnlyList(myIndex) & "")
                    currentScreenVar = rst("" & myDateVarList & "")
                    'Do nothing to myIndex because this is the last column
                    'Debug.Print ("Last date column of record row: " & myDateVarList)
                Else
                    'I am in the middle of the date columns, set priorScreenVar=x-1 currentScreenVar=x
                    priorScreenVar = rst("" & trialsOnlyList(myIndex) & "")
                    currentScreenVar = rst("" & myDateVarList & "")
                    'Debug.Print ("Values being compaired, Prior : Current " & priorScreenVar & " : " & currentScreenVar)
                    'Increment the myIndex counter +1
                    myIndex = myIndex + 1
                End If
            
                'Debug.Print ("currentScreenVar =" & currentScreenVar & "=")
                'Debug.Print ("priorScreenVar =" & priorScreenVar & "=")
                    
'                If currentScreenVar = "Not Found" _
'                    Or currentScreenVar = "POST BT Trial" _
'                    Or currentScreenVar = "POST AT Trial" _
'                    Or currentScreenVar = "POST FCT Trial" _
'                    Or currentScreenVar = "SPLIT" _
'                    Or currentScreenVar = "X/X" _
'                    Or currentScreenVar = "" _
'                    Or currentScreenVar = Empty _
'                    Then
                If currentScreenVar = "Not Found" _
                    Or currentScreenVar = "POST BT Trial" _
                    Or currentScreenVar = "POST AT Trial" _
                    Or currentScreenVar = "POST FCT Trial" _
                    Or currentScreenVar = "X/X" _
                    Or currentScreenVar = "" _
                    Or currentScreenVar = Empty _
                    Then
                    'Do not compare current screening against prior screening
                    mySparseVar = "~~~"
                ElseIf currentScreenVar = priorScreenVar Then
                    'The screening has not changed, do not mark this screening
                    'Debug.Print ("currentScreenVar = priorScreenVar : " & currentScreenVar & " EQUALS " & priorScreenVar)
                    mySparseVar = "~~~"
                ElseIf currentScreenVar <> priorScreenVar Then
                    'The screening has changed and mark this screening
                    'Debug.Print ("currentScreenVar = priorScreenVar : " & currentScreenVar & " EQUALS " & priorScreenVar)
'                    If priorScreenVar = "Not Found" _
'                        Or priorScreenVar = "POST BT Trial" _
'                        Or priorScreenVar = "POST AT Trial" _
'                        Or priorScreenVar = "POST FCT Trial" _
'                        Or priorScreenVar = "SPLIT" _
'                        Or priorScreenVar = "X/X" _
'                        Or priorScreenVar = "" _
'                        Or priorScreenVar = Empty _
'                        Then
                    If priorScreenVar = "Not Found" _
                        Or priorScreenVar = "POST BT Trial" _
                        Or priorScreenVar = "POST AT Trial" _
                        Or priorScreenVar = "POST FCT Trial" _
                        Or priorScreenVar = "X/X" _
                        Or priorScreenVar = "" _
                        Or priorScreenVar = Empty _
                        Then
                        mySparseVar = currentScreenVar
                    Else
                        mySparseVar = priorScreenVar & "/" & currentScreenVar
                    End If
                Else
                    mySparseVar = 99 ' This is a flag value that something is not correct
                End If
                
                'Debug.Print ("Cur TC: " & curTrial_Card)
                'Update the Sparse Matrix table with the Value of 0 or 1
'                CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " " _
'                & "SET " _
'                & "" & CurrentTable & ".[" & myDateVarList & "] = '" & mySparseVar & "' " _
'                & "WHERE " _
'                & "" & CurrentTable & ".Trial_Card = '" & curTrial_Card & "';"
                
                dbs_Write.Execute "UPDATE DISTINCTROW " & CurrentTable & " " _
                & "SET " _
                & "" & CurrentTable & ".[" & myDateVarList & "] = '" & mySparseVar & "' " _
                & "WHERE " _
                & "" & CurrentTable & ".Trial_Card = '" & curTrial_Card & "';"
                
            Next myDateVarList
            
            'When last column of row record is reached control is returned to this loop
         
            myDateVarList = Empty
            'Debug.Print ("Next row in Table")
            
            rst.MoveNext
            
        Loop
    
    Else
        'Un trapped error
    End If
    
    rst.Close
    dbs_Read.Close
    dbs_Write.Close
   
    'Debug.Print vbCrLf & "The Trial Cards Sparse Matrix Update Query completed." & vbCrLf

End Sub
