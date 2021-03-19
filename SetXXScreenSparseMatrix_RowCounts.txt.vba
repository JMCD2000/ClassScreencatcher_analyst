Attribute VB_Name = "SetXXScreenSparseMatrix_RowCounts"
Option Compare Database
'This has issues with the way that splits are handeled and the way that cards are handeled prior to the Event
'The use of SPLIT until the actual split occures causes counting issues because the screening changes from SPLIT to the actual screening
'The use of NotFound until the trial card is actually written at Event just looks messy, would like to change this
Option Explicit


Public Sub Set_SparseMatrix_Counts_XXSMC()
'This counts the sparse matrix by row in the event to owld date range.
Dim myDateVarList As Variant ' Used to cycle through each date column as an iterable
'These are used on the recordset
Dim curDateVal As Integer ' This holds the current date column value
Dim cur_aCL_Index As Integer ' This holds the current index of the column name in the allColumnsList
Dim ScreenCountsVar_BT As Integer ' This holds the count of the screens as they are collected accross the row
Dim ScreenCountsVar_AT As Integer ' This holds the count of the screens as they are collected accross the row
Dim ScreenCountsVar_FCT As Integer ' This holds the count of the screens as they are collected accross the row
Dim curTrial_Card As String ' This is for the update back to the TC_Screen_Agg Table
'Open a recordset to loop through each Record in the recordset
Dim dbs_Read As DAO.Database
Dim rst As DAO.Recordset ' or Recordset2?
Dim mySQLstring As String
'Open a recordset to update the TC_Screen_Agg table
Dim dbs_Write As DAO.Database
Dim mySQL_update As String ' Updating SQL String

    'simple varible referenced typed SQL statement
    ' "SELECT " & SparseRefTable & ".* FROM " & SparseRefTable & " INNER JOIN " & CurrentTable & " ON " & SparseRefTable & ".Trial_Card = " & CurrentTable & ".Trial_Card;"
    mySQLstring = "SELECT " & CurrentTable & ".* FROM " & CurrentTable & ";"

    'Open a pointer to current database
    Set dbs_Read = CurrentDb()
    Set dbs_Write = CurrentDb()
    
    'Create the recordset with my SQL string
    Set rst = dbs_Read.OpenRecordset(mySQLstring)
    rst.MoveFirst
    
    'Check for All Reports or only Events
    If All_or_Events = "All" Then
        Do While Not rst.EOF
            'Step down each record row
            curTrial_Card = rst![Trial_Card]
            ScreenCountsVar_BT = 0
            ScreenCountsVar_AT = 0
            ScreenCountsVar_FCT = 0
            'Debug.Print ("Cur TC: " & curTrial_Card)
            
            For Each myDateVarList In allColumnsList
                'Step accross each date column within the current record row
                'Debug.Print ("TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                
                'Convert the string values to a numeric value with test
                If rst.Fields(myDateVarList).Value = "~~~" Then
                    curDateVal = 0 ' This means there was no screening change, don't count
                ElseIf rst.Fields(myDateVarList).Value = "99" Then
                    curDateVal = 0 ' This means there was a data error, don't count
                ElseIf rst.Fields(myDateVarList).Value = "" Then
                    curDateVal = 0 ' This means there was a data error, don't count
                ElseIf rst.Fields(myDateVarList).Value = Empty Then
                    curDateVal = 0 ' This means there was a data error, don't count
                ElseIf rst.Fields(myDateVarList).Value <> "~~~" Then
                    curDateVal = 1 ' This means there was a screening change
                Else
                    curDateVal = 0 ' This means there was a data error, don't count
                End If
                
                cur_aCL_Index = allColumnsList.IndexOf(myDateVarList, 0)
                
                'Check if the current date value is in BT Event range and add value
                If cur_aCL_Index < allColumnsList_BT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_BT = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = allColumnsList_BT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_BT = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < (allColumnsList_OWLD + 1) Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_BT = curDateVal + ScreenCountsVar_BT
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > allColumnsList_OWLD Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_BT = curDateVal + ScreenCountsVar_BT
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
            
                'Check if the current date value is in AT Event range and add value
                If cur_aCL_Index < allColumnsList_AT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_AT = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = allColumnsList_AT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_AT = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < (allColumnsList_OWLD + 1) Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_AT = curDateVal + ScreenCountsVar_AT
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > allColumnsList_OWLD Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_AT = curDateVal + ScreenCountsVar_AT
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
                
                'Check if the current date value is in FCT Event range and add value
                If cur_aCL_Index < allColumnsList_FCT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_FCT = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = allColumnsList_FCT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_FCT = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < (allColumnsList_OWLD + 1) Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_FCT = curDateVal + ScreenCountsVar_FCT
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > allColumnsList_OWLD Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_FCT = curDateVal + ScreenCountsVar_FCT
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
                
            Next myDateVarList
            
            'Debug.Print ("Cur TC: " & curTrial_Card)
            'Update the Sparse Matrix Event Ranges with the Values
            mySQL_update = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_counts_BT_to_OWLD] = " & ScreenCountsVar_BT & ", " & CurrentTable & ".[Rescreen_counts_AT_to_OWLD] = " & ScreenCountsVar_AT & ", " & CurrentTable & ".[Rescreen_counts_FCT_to_OWLD] = " & ScreenCountsVar_FCT & " WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
                
            'Debug.Print ("mySQL_update = " & mySQL_update)
            dbs_Write.Execute mySQL_update
                
            'When last column of the row record is reached, goto next row record
         
            myDateVarList = Empty
            'Debug.Print ("Next row in Table")
            
            rst.MoveNext
            
        Loop
    
    ElseIf All_or_Events = "Events" Then
        Do While Not rst.EOF
            'Step down each record row
            curTrial_Card = rst![Trial_Card]
            ScreenCountsVar_BT = 0
            ScreenCountsVar_AT = 0
            ScreenCountsVar_FCT = 0
            'Debug.Print ("Cur TC: " & curTrial_Card)
            
            For Each myDateVarList In trialsOnlyList
                'Step accross each date column within the current record row
                'Debug.Print ("TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                'curDateVal = rst![myDateVarList]
                
                'Convert the string values to a numeric value with test
                If rst.Fields(myDateVarList).Value = "~~~" Then
                    curDateVal = 0 ' This means there was no screening change, don't count
                ElseIf rst.Fields(myDateVarList).Value = "99" Then
                    curDateVal = 0 ' This means there was a data error, don't count
                ElseIf rst.Fields(myDateVarList).Value = "" Then
                    curDateVal = 0 ' This means there was a data error, don't count
                ElseIf rst.Fields(myDateVarList).Value = Empty Then
                    curDateVal = 0 ' This means there was a data error, don't count
                ElseIf rst.Fields(myDateVarList).Value <> "~~~" Then
                    curDateVal = 1 ' This means there was a screening change
                Else
                    curDateVal = 0 ' This means there was a data error, don't count
                End If
                
                cur_aCL_Index = trialsOnlyList.IndexOf(myDateVarList, 0)
                
                'Check if the current date value is in BT Event range and add value
                If cur_aCL_Index < trialsOnlyList_BT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_BT = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = trialsOnlyList_BT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_BT = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < (trialsOnlyList_OWLD + 1) Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_BT = curDateVal + ScreenCountsVar_BT
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > trialsOnlyList_OWLD Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_BT = curDateVal + ScreenCountsVar_BT
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
            
                'Check if the current date value is in AT Event range and add value
                If cur_aCL_Index < trialsOnlyList_AT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_AT = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = trialsOnlyList_AT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_AT = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < (trialsOnlyList_OWLD + 1) Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_AT = curDateVal + ScreenCountsVar_AT
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > trialsOnlyList_OWLD Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_AT = curDateVal + ScreenCountsVar_AT
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
                
                'Check if the current date value is in FCT Event range and add value
                If cur_aCL_Index < trialsOnlyList_FCT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_FCT = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = trialsOnlyList_FCT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_FCT = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < (trialsOnlyList_OWLD + 1) Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_FCT = curDateVal + ScreenCountsVar_FCT
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > trialsOnlyList_OWLD Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_FCT = curDateVal + ScreenCountsVar_FCT
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
                
            Next myDateVarList
            
            'Debug.Print ("Cur TC: " & curTrial_Card)
            'Update the Sparse Matrix Event Ranges with the Values
            mySQL_update = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_counts_BT_to_OWLD] = " & ScreenCountsVar_BT & ", " & CurrentTable & ".[Rescreen_counts_AT_to_OWLD] = " & ScreenCountsVar_AT & ", " & CurrentTable & ".[Rescreen_counts_FCT_to_OWLD] = " & ScreenCountsVar_FCT & " WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
                
            'Debug.Print ("mySQL_update = " & mySQL_update)
            dbs_Write.Execute mySQL_update
                
            'When last column of the row record is reached, goto next row record
         
            myDateVarList = Empty
            'Debug.Print ("Next row in Table")
            
            rst.MoveNext
            
        Loop
    
    Else
        'Un trapped error
        'All_or_Events Global is empty or not expected value
        Debug.Print "Function Set_SparseMatrix_Counts_XXSMC() was passed empty or not expected value with GLOBAL All_or_Events:= " & All_or_Events & "."
    End If
    
    rst.Close
    dbs_Read.Close
    dbs_Write.Close
    
    ' Debug.Print vbCrLf & "The Trial Cards Sparse Matrix Update Query completed." & vbCrLf

End Sub

Public Sub Set_SparseMatrix_Counts_XXSMC_DEL()
'This counts the sparse matrix by row in the event to owld date range.
Dim myDateVarList As Variant ' Used to cycle through each date column as an iterable
'These are used on the recordset
Dim curDateVal As Integer ' This holds the current date column value
Dim cur_aCL_Index As Integer ' This holds the current index of the column name in the allColumnsList
Dim ScreenCountsVar_BT_DEL As Integer ' This holds the count of the screens as they are collected accross the row for milestone DEL
Dim ScreenCountsVar_AT_DEL As Integer ' This holds the count of the screens as they are collected accross the row for milestone DEL
Dim curTrial_Card As String ' This is for the update back to the TC_Screen_Agg Table
'Open a recordset to loop through each Record in the recordset
Dim dbs_Read As DAO.Database
Dim rst As DAO.Recordset ' or Recordset2?
Dim mySQLstring As String
'Open a recordset to update the TC_Screen_Agg table
Dim dbs_Write As DAO.Database
Dim mySQL_update As String ' Updating SQL String

    'simple varible referenced typed SQL statement
    ' "SELECT " & SparseRefTable & ".* FROM " & SparseRefTable & " INNER JOIN " & CurrentTable & " ON " & SparseRefTable & ".Trial_Card = " & CurrentTable & ".Trial_Card;"
    mySQLstring = "SELECT " & CurrentTable & ".* FROM " & CurrentTable & ";"

    'Open a pointer to current database
    Set dbs_Read = CurrentDb()
    Set dbs_Write = CurrentDb()
    
    'Create the recordset with my SQL string
    Set rst = dbs_Read.OpenRecordset(mySQLstring)
    rst.MoveFirst
    
    'Check for All Reports or only Events
    If All_or_Events = "All" Then
        Do While Not rst.EOF
            'Step down each record row
            curTrial_Card = rst![Trial_Card]
            ScreenCountsVar_BT_DEL = 0
            ScreenCountsVar_AT_DEL = 0
            'Debug.Print ("Cur TC: " & curTrial_Card)
            
            For Each myDateVarList In allColumnsList
                'Step accross each date column within the current record row
                'Debug.Print ("TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                
                'Convert the string values to a numeric value with test
                If rst.Fields(myDateVarList).Value = "~~~" Then
                    curDateVal = 0 ' This means there was no screening change, don't count
                ElseIf rst.Fields(myDateVarList).Value = "99" Then
                    curDateVal = 0 ' This means there was a data error, don't count
                ElseIf rst.Fields(myDateVarList).Value = "" Then
                    curDateVal = 0 ' This means there was a data error, don't count
                ElseIf rst.Fields(myDateVarList).Value = Empty Then
                    curDateVal = 0 ' This means there was a data error, don't count
                ElseIf rst.Fields(myDateVarList).Value <> "~~~" Then
                    curDateVal = 1 ' This means there was a screening change
                Else
                    curDateVal = 0 ' This means there was a data error, don't count
                End If
                
                cur_aCL_Index = allColumnsList.IndexOf(myDateVarList, 0)
                
                'Check if the current date value is in BT Event range and add value
                If cur_aCL_Index < allColumnsList_BT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_BT = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = allColumnsList_BT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_BT_DEL = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < (allColumnsList_DEL + 1) Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_BT_DEL = curDateVal + ScreenCountsVar_BT_DEL
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > allColumnsList_DEL Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_BT = curDateVal + ScreenCountsVar_BT
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
            
                'Check if the current date value is in AT Event range and add value
                If cur_aCL_Index < allColumnsList_AT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_AT = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = allColumnsList_AT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_AT_DEL = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < (allColumnsList_DEL + 1) Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_AT_DEL = curDateVal + ScreenCountsVar_AT_DEL
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > allColumnsList_DEL Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_AT = curDateVal + ScreenCountsVar_AT
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
                
            Next myDateVarList
            
            'Debug.Print ("Cur TC: " & curTrial_Card)
            'Update the Sparse Matrix Event Ranges with the Values
            mySQL_update = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_counts_BT_to_DEL] = " & ScreenCountsVar_BT_DEL & ", " & CurrentTable & ".[Rescreen_counts_AT_to_DEL] = " & ScreenCountsVar_AT_DEL & " WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
                
            'Debug.Print ("mySQL_update = " & mySQL_update)
            dbs_Write.Execute mySQL_update
                
            'When last column of the row record is reached, goto next row record
         
            myDateVarList = Empty
            'Debug.Print ("Next row in Table")
            
            rst.MoveNext
            
        Loop
    
    ElseIf All_or_Events = "Events" Then
        Do While Not rst.EOF
            'Step down each record row
            curTrial_Card = rst![Trial_Card]
            ScreenCountsVar_BT_DEL = 0
            ScreenCountsVar_AT_DEL = 0
            'Debug.Print ("Cur TC: " & curTrial_Card)
            
            For Each myDateVarList In trialsOnlyList
                'Step accross each date column within the current record row
                'Debug.Print ("TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                'curDateVal = rst![myDateVarList]
                
                'Convert the string values to a numeric value with test
                If rst.Fields(myDateVarList).Value = "~~~" Then
                    curDateVal = 0 ' This means there was no screening change, don't count
                ElseIf rst.Fields(myDateVarList).Value = "99" Then
                    curDateVal = 0 ' This means there was a data error, don't count
                ElseIf rst.Fields(myDateVarList).Value = "" Then
                    curDateVal = 0 ' This means there was a data error, don't count
                ElseIf rst.Fields(myDateVarList).Value = Empty Then
                    curDateVal = 0 ' This means there was a data error, don't count
                ElseIf rst.Fields(myDateVarList).Value <> "~~~" Then
                    curDateVal = 1 ' This means there was a screening change
                Else
                    curDateVal = 0 ' This means there was a data error, don't count
                End If
                
                cur_aCL_Index = trialsOnlyList.IndexOf(myDateVarList, 0)
                
                'Check if the current date value is in BT Event range and add value
                If cur_aCL_Index < trialsOnlyList_BT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_BT = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = trialsOnlyList_BT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_BT_DEL = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < (trialsOnlyList_DEL + 1) Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_BT_DEL = curDateVal + ScreenCountsVar_BT_DEL
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > trialsOnlyList_DEL Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_BT = curDateVal + ScreenCountsVar_BT
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
            
                'Check if the current date value is in AT Event range and add value
                If cur_aCL_Index < trialsOnlyList_AT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_AT = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = trialsOnlyList_AT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_AT_DEL = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < (trialsOnlyList_DEL + 1) Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_AT_DEL = curDateVal + ScreenCountsVar_AT_DEL
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > trialsOnlyList_DEL Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_AT = curDateVal + ScreenCountsVar_AT
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
                
            Next myDateVarList
            
            'Debug.Print ("Cur TC: " & curTrial_Card)
            'Update the Sparse Matrix Event Ranges with the Values
            mySQL_update = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_counts_BT_to_DEL] = " & ScreenCountsVar_BT_DEL & ", " & CurrentTable & ".[Rescreen_counts_AT_to_DEL] = " & ScreenCountsVar_AT_DEL & " WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
                
            'Debug.Print ("mySQL_update = " & mySQL_update)
            dbs_Write.Execute mySQL_update
                
            'When last column of the row record is reached, goto next row record
         
            myDateVarList = Empty
            'Debug.Print ("Next row in Table")
            
            rst.MoveNext
            
        Loop
    
    Else
        'Un trapped error
        'All_or_Events Global is empty or not expected value
        Debug.Print "Function Set_SparseMatrix_Counts_XXSMC_DEL() was passed empty or not expected value with GLOBAL All_or_Events:= " & All_or_Events & "."
    End If
    
    rst.Close
    dbs_Read.Close
    dbs_Write.Close
    
    ' Debug.Print vbCrLf & "The Trial Cards Sparse Matrix Update Query completed." & vbCrLf

End Sub

