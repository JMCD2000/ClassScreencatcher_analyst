Option Compare Database

Option Explicit


Public Sub Set_SparseMatrix_Counts_XXSMC()
'This counts the sparse matrix by row in the event to DEL/OWLD/FINAL date ranges.
Dim myDateVarList As Variant ' Used to cycle through each date column as an iterable
'These are used on the recordset
Dim curDateVal As Integer ' This holds the current date column value
Dim cur_aCL_Index As Integer ' This holds the current index of the column name in the allColumnsList
Dim ScreenCountsVar_BT_DEL As Integer ' This holds the count of the screens as they are collected accross the row for milestone DEL
Dim ScreenCountsVar_AT_DEL As Integer ' This holds the count of the screens as they are collected accross the row for milestone DEL
Dim ScreenCountsVar_BT_OWLD As Integer ' This holds the count of the screens as they are collected accross the row to OWLD
Dim ScreenCountsVar_AT_OWLD As Integer ' This holds the count of the screens as they are collected accross the row to OWLD
Dim ScreenCountsVar_FCT_OWLD As Integer ' This holds the count of the screens as they are collected accross the row to OWLD
Dim ScreenCountsVar_BT_Final As Integer ' This holds the count of the screens as they are collected accross the row to Final
Dim ScreenCountsVar_AT_Final As Integer ' This holds the count of the screens as they are collected accross the row to Final
Dim ScreenCountsVar_FCT_Final As Integer ' This holds the count of the screens as they are collected accross the row to Final
Dim curTrial_Card As String ' This is for the update back to the TC_Screen_Agg Table
'Open a recordset to loop through each Record in the recordset
Dim dbs_Read As DAO.Database
Dim rst As DAO.Recordset ' or Recordset2?
Dim mySQLstring As String
'Open a recordset to update the TC_Screen_Agg table
Dim dbs_Write As DAO.Database
Dim mySQL_update_BT As String ' Updating SQL String
Dim mySQL_update_AT As String ' Updating SQL String
Dim mySQL_update_FCT As String ' Updating SQL String

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
            ScreenCountsVar_BT_OWLD = 0
            ScreenCountsVar_AT_OWLD = 0
            ScreenCountsVar_FCT_OWLD = 0
            ScreenCountsVar_BT_Final = 0
            ScreenCountsVar_AT_Final = 0
            ScreenCountsVar_FCT_Final = 0
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
                
                'BT Count to DEL
                'Check if the current date value is in BT Event range and add value
                If cur_aCL_Index < allColumnsList_BT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_BT_DEL = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = allColumnsList_BT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_BT_DEL = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < allColumnsList_DEL Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_BT_DEL = curDateVal + ScreenCountsVar_BT_DEL
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = allColumnsList_DEL Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_BT_DEL = curDateVal + ScreenCountsVar_BT_DEL
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > allColumnsList_DEL Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_BT_DEL = curDateVal + ScreenCountsVar_BT_DEL
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
                
                'BT Count to OWLD
                'Check if the current date value is in BT Event range and add value
                If cur_aCL_Index < allColumnsList_BT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_BT_OWLD = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = allColumnsList_BT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_BT_OWLD = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < allColumnsList_OWLD Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_BT_OWLD = curDateVal + ScreenCountsVar_BT_OWLD
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = allColumnsList_OWLD Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_BT_OWLD = curDateVal + ScreenCountsVar_BT_OWLD
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > allColumnsList_OWLD Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_BT_OWLD = curDateVal + ScreenCountsVar_BT_OWLD
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
                
                'BT Count to Final
                'Check if the current date value is in BT Event range and add value
                If cur_aCL_Index < allColumnsList_BT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_BT_Final = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = allColumnsList_BT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_BT_Final = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < allColumnsList_Final Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_BT_Final = curDateVal + ScreenCountsVar_BT_Final
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = allColumnsList_Final Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_BT_Final = curDateVal + ScreenCountsVar_BT_Final
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > allColumnsList_Final Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_BT_Final = curDateVal + ScreenCountsVar_BT_Final
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
            
                'AT Count to DEL
                'Check if the current date value is in AT Event range and add value
                If cur_aCL_Index < allColumnsList_AT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_AT_DEL = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = allColumnsList_AT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_AT_DEL = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < allColumnsList_DEL Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_AT_DEL = curDateVal + ScreenCountsVar_AT_DEL
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = allColumnsList_DEL Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_AT_DEL = curDateVal + ScreenCountsVar_AT_DEL
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > allColumnsList_DEL Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_AT_DEL = curDateVal + ScreenCountsVar_AT_DEL
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
                
                'AT Count to OWLD
                'Check if the current date value is in AT Event range and add value
                If cur_aCL_Index < allColumnsList_AT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_AT_OWLD = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = allColumnsList_AT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_AT_OWLD = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < allColumnsList_OWLD Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_AT_OWLD = curDateVal + ScreenCountsVar_AT_OWLD
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = allColumnsList_OWLD Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_AT_OWLD = curDateVal + ScreenCountsVar_AT_OWLD
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > allColumnsList_OWLD Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_AT_OWLD = curDateVal + ScreenCountsVar_AT_OWLD
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
                
                'AT Count to Final
                'Check if the current date value is in AT Event range and add value
                If cur_aCL_Index < allColumnsList_AT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_AT_Final = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = allColumnsList_AT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_AT_Final = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < allColumnsList_Final Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_AT_Final = curDateVal + ScreenCountsVar_AT_Final
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = allColumnsList_Final Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_AT_Final = curDateVal + ScreenCountsVar_AT_Final
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > allColumnsList_Final Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_AT_Final = curDateVal + ScreenCountsVar_AT_Final
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
                
                'FCT Count to OWLD
                'Check if the current date value is in FCT Event range and add value
                If cur_aCL_Index < allColumnsList_FCT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_FCT_OWLD = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = allColumnsList_FCT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_FCT_OWLD = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < allColumnsList_OWLD Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_FCT_OWLD = curDateVal + ScreenCountsVar_FCT_OWLD
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = allColumnsList_OWLD Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_FCT_OWLD = curDateVal + ScreenCountsVar_FCT_OWLD
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > allColumnsList_OWLD Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_FCT_OWLD = curDateVal + ScreenCountsVar_FCT_OWLD
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
                
                'FCT Count to Final
                'Check if the current date value is in FCT Event range and add value
                If cur_aCL_Index < allColumnsList_FCT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_FCT_Final = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = allColumnsList_FCT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_FCT_Final = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < allColumnsList_Final Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_FCT_Final = curDateVal + ScreenCountsVar_FCT_Final
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = allColumnsList_Final Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_FCT_Final = curDateVal + ScreenCountsVar_FCT_Final
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > allColumnsList_Final Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_FCT_Final = curDateVal + ScreenCountsVar_FCT_Final
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
                
            Next myDateVarList
            
            'Debug.Print ("Cur TC: " & curTrial_Card)
            'Update the Sparse Matrix BT Event Ranges with the Values
            mySQL_update_BT = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_counts_BT_to_DEL] = " & ScreenCountsVar_BT_DEL & ", " & CurrentTable & ".[Rescreen_counts_BT_to_OWLD] = " & ScreenCountsVar_BT_OWLD & ", " & CurrentTable & ".[Rescreen_counts_BT_to_Final] = " & ScreenCountsVar_BT_Final & " WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
            
            'Update the Sparse Matrix AT Event Ranges with the Values
            mySQL_update_AT = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_counts_AT_to_DEL] = " & ScreenCountsVar_AT_DEL & ", " & CurrentTable & ".[Rescreen_counts_AT_to_OWLD] = " & ScreenCountsVar_AT_OWLD & ", " & CurrentTable & ".[Rescreen_counts_AT_to_Final] = " & ScreenCountsVar_AT_Final & " WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
            
            'Update the Sparse Matrix FCT Event Ranges with the Values
            mySQL_update_FCT = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_counts_FCT_to_OWLD] = " & ScreenCountsVar_FCT_OWLD & ", " & CurrentTable & ".[Rescreen_counts_FCT_to_Final] = " & ScreenCountsVar_FCT_Final & " WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
            
            'Debug.Print ("mySQL_update = " & mySQL_update)
            dbs_Write.Execute mySQL_update_BT
            dbs_Write.Execute mySQL_update_AT
            dbs_Write.Execute mySQL_update_FCT
            
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
            ScreenCountsVar_BT_OWLD = 0
            ScreenCountsVar_AT_OWLD = 0
            ScreenCountsVar_FCT_OWLD = 0
            ScreenCountsVar_BT_Final = 0
            ScreenCountsVar_AT_Final = 0
            ScreenCountsVar_FCT_Final = 0
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
                
                'BT Count to DEL
                'Check if the current date value is in BT Event range and add value
                
                If cur_aCL_Index < trialsOnlyList_BT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_BT_DEL = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = trialsOnlyList_BT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_BT_DEL = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < trialsOnlyList_DEL Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_BT_DEL = curDateVal + ScreenCountsVar_BT_DEL
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = trialsOnlyList_DEL Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_BT_DEL = curDateVal + ScreenCountsVar_BT_DEL
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > trialsOnlyList_DEL Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_BT_DEL = curDateVal + ScreenCountsVar_BT_DEL
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
                
                'BT Count to OWLD
                'Check if the current date value is in BT Event range and add value
                If cur_aCL_Index < trialsOnlyList_BT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_BT_OWLD = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = trialsOnlyList_BT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_BT_OWLD = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < trialsOnlyList_OWLD Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_BT_OWLD = curDateVal + ScreenCountsVar_BT_OWLD
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = trialsOnlyList_OWLD Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_BT_OWLD = curDateVal + ScreenCountsVar_BT_OWLD
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > trialsOnlyList_OWLD Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_BT_OWLD = curDateVal + ScreenCountsVar_BT_OWLD
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
                
                'BT Count to Final
                'Check if the current date value is in BT Event range and add value
                If cur_aCL_Index < trialsOnlyList_BT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_BT_Final = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = trialsOnlyList_BT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_BT_Final = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < trialsOnlyList_Final Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_BT_Final = curDateVal + ScreenCountsVar_BT_Final
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = trialsOnlyList_Final Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_BT_Final = curDateVal + ScreenCountsVar_BT_Final
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > trialsOnlyList_Final Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_BT_Final = curDateVal + ScreenCountsVar_BT_Final
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
            
                'AT Count to DEL
                'Check if the current date value is in AT Event range and add value
                If cur_aCL_Index < trialsOnlyList_AT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_AT_DEL = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = trialsOnlyList_AT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_AT_DEL = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < trialsOnlyList_DEL Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_AT_DEL = curDateVal + ScreenCountsVar_AT_DEL
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = trialsOnlyList_DEL Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_AT_DEL = curDateVal + ScreenCountsVar_AT_DEL
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > trialsOnlyList_DEL Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_AT_DEL = curDateVal + ScreenCountsVar_AT_DEL
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
                
                'AT Count to OWLD
                'Check if the current date value is in AT Event range and add value
                If cur_aCL_Index < trialsOnlyList_AT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_AT_OWLD = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = trialsOnlyList_AT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_AT_OWLD = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < trialsOnlyList_OWLD Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_AT_OWLD = curDateVal + ScreenCountsVar_AT_OWLD
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = trialsOnlyList_OWLD Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_AT_OWLD = curDateVal + ScreenCountsVar_AT_OWLD
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > trialsOnlyList_OWLD Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_AT_OWLD = curDateVal + ScreenCountsVar_AT_OWLD
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
                
                'AT Count to Final
                'Check if the current date value is in AT Event range and add value
                If cur_aCL_Index < trialsOnlyList_AT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_AT_Final = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = trialsOnlyList_AT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_AT_Final = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < trialsOnlyList_Final Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_AT_Final = curDateVal + ScreenCountsVar_AT_Final
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = trialsOnlyList_Final Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_AT_Final = curDateVal + ScreenCountsVar_AT_Final
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > trialsOnlyList_Final Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_AT_Final = curDateVal + ScreenCountsVar_AT_Final
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
                
                'FCT Count to OWLD
                'Check if the current date value is in FCT Event range and add value
                If cur_aCL_Index < trialsOnlyList_FCT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_FCT_OWLD = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = trialsOnlyList_FCT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_FCT_OWLD = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < trialsOnlyList_OWLD Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_FCT_OWLD = curDateVal + ScreenCountsVar_FCT_OWLD
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = trialsOnlyList_OWLD Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_FCT_OWLD = curDateVal + ScreenCountsVar_FCT_OWLD
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > trialsOnlyList_OWLD Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_FCT_OWLD = curDateVal + ScreenCountsVar_FCT_OWLD
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
                
                'FCT Count to Final
                'Check if the current date value is in FCT Event range and add value
                If cur_aCL_Index < trialsOnlyList_FCT Then
                    'I am before the Event I am counting, do nothing
                    'ScreenCountsVar_FCT_Final = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = trialsOnlyList_FCT Then
                    'I am on the first Event column I am counting
                    ScreenCountsVar_FCT_Final = curDateVal
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index < trialsOnlyList_Final Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_FCT_Final = curDateVal + ScreenCountsVar_FCT_Final
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index = trialsOnlyList_Final Then
                    'I am in the Event Range of the date columns
                    ScreenCountsVar_FCT_Final = curDateVal + ScreenCountsVar_FCT_Final
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                ElseIf cur_aCL_Index > trialsOnlyList_Final Then
                    'I am after the Event Range of the date columns, do nothing
                    'ScreenCountsVar_FCT_Final = curDateVal + ScreenCountsVar_FCT_Final
                    'Leave myIndex alone here so other Event ranges can be checked
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                End If
                
            Next myDateVarList
            
            'Debug.Print ("Cur TC: " & curTrial_Card)
            'Update the Sparse Matrix BT Event Ranges with the Values
            mySQL_update_BT = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_counts_BT_to_DEL] = " & ScreenCountsVar_BT_DEL & ", " & CurrentTable & ".[Rescreen_counts_BT_to_OWLD] = " & ScreenCountsVar_BT_OWLD & ", " & CurrentTable & ".[Rescreen_counts_BT_to_Final] = " & ScreenCountsVar_BT_Final & " WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
            
            'Update the Sparse Matrix AT Event Ranges with the Values
            mySQL_update_AT = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_counts_AT_to_DEL] = " & ScreenCountsVar_AT_DEL & ", " & CurrentTable & ".[Rescreen_counts_AT_to_OWLD] = " & ScreenCountsVar_AT_OWLD & ", " & CurrentTable & ".[Rescreen_counts_AT_to_Final] = " & ScreenCountsVar_AT_Final & " WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
            
            'Update the Sparse Matrix FCT Event Ranges with the Values
            mySQL_update_FCT = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_counts_FCT_to_OWLD] = " & ScreenCountsVar_FCT_OWLD & ", " & CurrentTable & ".[Rescreen_counts_FCT_to_Final] = " & ScreenCountsVar_FCT_Final & " WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
            
            'Debug.Print ("mySQL_update = " & mySQL_update)
            dbs_Write.Execute mySQL_update_BT
            dbs_Write.Execute mySQL_update_AT
            dbs_Write.Execute mySQL_update_FCT
                
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
