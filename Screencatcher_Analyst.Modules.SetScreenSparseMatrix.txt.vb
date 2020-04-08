Option Compare Database
Option Explicit


Public Function SetCurrentWorkingTable_SOSM(ByVal MatrixTBL As String, ByVal ReferenceTBL As String)
'This sub is setting the current working _
table that is used in all the SQL statements.

CurrentTable = MatrixTBL
SparseRefTable = ReferenceTBL

End Function


Public Sub SetFirstScreensAndEvents_SOSM()
'This is to load or reload the screening sparse matrix into the <TABLE>_SparseMatrix

'Set data columns to <Not Found>
    Dim myDateVarList As Variant
    Dim notFound As String 'The <Not Found> is not used in TSM or elsewhere, becomes a visual that something was missed
    notFound = "98"

    For Each myDateVarList In allColumnsList
        CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & SparseRefTable & "] ON " & CurrentTable & ".Trial_Card = [" & SparseRefTable & "].Trial_Card " _
        & "SET " _
        & "" & CurrentTable & ".[" & myDateVarList & "] = '" & notFound & "';"
        ' Debug.Print "done with column table: " & myDateVarList & "."
    Next myDateVarList
    
    myDateVarList = Empty

'Set Trial ID, Event and Final Status values as place holder values
    Dim emptyID As String 'The dash is not used in TSM or elsewhere, becomes a visual that something was missed
    emptyID = "-"
    
    Dim emptyEvent As String 'The double E is not used in TSM or elsewhere, becomes a visual that something was missed
    emptyEvent = "EE"
    
    Dim emptySts_A_T As String 'The dash slash dash is not used in TSM or elsewhere, becomes a visual that something was missed
    emptySts_A_T = "-/-"

    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & SparseRefTable & "] ON " & CurrentTable & ".Trial_Card = [" & SparseRefTable & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".Trial_ID = '" & emptyID & "', " _
    & "" & CurrentTable & ".Event = '" & emptyEvent & "', " _
    & "" & CurrentTable & ".Final_Sts_A_T = '" & emptySts_A_T & "';"
    ' Debug.Print "Completed the" & CurrentTable & "table data set with place holder values Update Query."

'Set Trial ID, Event and Final Status values
    CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [" & SparseRefTable & "] ON " & CurrentTable & ".Trial_Card = [" & SparseRefTable & "].Trial_Card " _
    & "SET " _
    & "" & CurrentTable & ".Trial_ID = [" & SparseRefTable & "].[Trial_ID], " _
    & "" & CurrentTable & ".Event = [" & SparseRefTable & "].[Event], " _
    & "" & CurrentTable & ".Final_Sts_A_T = [" & SparseRefTable & "].[Final_Sts_A_T];"
    ' Debug.Print "Completed setting values in columns Trial_ID, Event, Final_Sts_A_T."

'ClearMyDateLists 'Empty created list objects

Debug.Print vbCrLf & "The " & CurrentTable & " Table is now ready for Sparse Matrix Calculations." & vbCrLf

End Sub


Public Sub Build_and_Set_Aggregated_Screen_SOSM()
'This builds the sparse matrix that has a 0 if no change over the prior _
 or has a 1 if there was a change.
Dim myDateVarList As Variant ' Used to cycle through each date column as an iterable
Dim myIndex As Integer ' This is to get the prior screen
'These are used on the recordset
Dim priorScreenVar As String ' This holds the first screen from the date columns
Dim currentScreenVar As String ' This holds the last column looked at value to see if current column is different
Dim mySparseVar As Integer ' This holds the
Dim aggregatedScreensVar As String ' This holds the screens as they are collected
Dim curTrial_Card As String ' This is for the update back to the TC_Screen_Agg Table
'Open a recordset to loop through each Record in the recordset
Dim dbs As DAO.Database
Dim rst As DAO.Recordset ' or Recordset2?
Dim mySQLstring As String
'Open a recordset to update the TC_Screen_Agg table
Dim dbs_Sparse As DAO.Database

mySQLstring = "SELECT " & SparseRefTable & ".* FROM " & SparseRefTable & " INNER JOIN " & CurrentTable & " ON " & SparseRefTable & ".Trial_Card = " & CurrentTable & ".Trial_Card;"

'Open a pointer to current database
Set dbs = CurrentDb()
Set dbs_Sparse = CurrentDb()

'Create the recordset with my SQL string
Set rst = dbs.OpenRecordset(mySQLstring)

Do While Not rst.EOF
    'Step down each record row
    myIndex = 0 'This is the first date column in the allColumnsList, reset to 0 on each row
    curTrial_Card = rst![Trial_Card]
    'Debug.Print ("Cur TC: " & curTrial_Card)
    
    For Each myDateVarList In allColumnsList
        'Step accross each date column within the current record row
        'Debug.Print ("TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
        'Default value for the sparse matrix reset at each column itteration
        mySparseVar = 0
                
        'Assign the prior screen from the first date column
        If myDateVarList = allColumnsList(0) Then
            'I am on the first date column, set prior screen = [First_Screening]
            priorScreenVar = rst![First_Screening]
            currentScreenVar = rst("" & myDateVarList & "") ' rst![ & myDateVarList & ]
            aggregatedScreensVar = rst![First_Screening]
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
            
        If currentScreenVar = "Not Found" _
            Or currentScreenVar = "POST BT Trial" _
            Or currentScreenVar = "POST AT Trial" _
            Or currentScreenVar = "POST FCT Trial" _
            Or currentScreenVar = "SPLIT" _
            Or currentScreenVar = "X/X" _
            Or currentScreenVar = "" _
            Or currentScreenVar = Empty _
            Then
            'Do not compare current screening against prior screening
            mySparseVar = 0
        ElseIf currentScreenVar = priorScreenVar Then
            'The screening has not changed, do not mark this screening
            'Debug.Print ("currentScreenVar = priorScreenVar : " & currentScreenVar & " EQUALS " & priorScreenVar)
            mySparseVar = 0
        ElseIf currentScreenVar <> priorScreenVar Then
            'The screening has changed and mark this screening
            'Debug.Print ("currentScreenVar = priorScreenVar : " & currentScreenVar & " EQUALS " & priorScreenVar)
            mySparseVar = 1
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
        CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " " _
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

rst.Close
dbs.Close
dbs_Sparse.Close

'ClearMyDateLists 'Empty created list objects

Debug.Print vbCrLf & "The Trial Cards Sparse Matrix Update Query completed." & vbCrLf

End Sub
