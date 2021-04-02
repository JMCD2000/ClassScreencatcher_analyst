Option Compare Database
Option Explicit

Public Sub Build_and_Set_Aggregated_Screen_and_Counts_TCSA()
'This counts the sparse matrix by row in the event to DEL/OWLD/FINAL date ranges.
Dim myDateVarList As Variant ' Used to cycle through each date column as an iterable
'These are used on the recordset
'Dim curDateVal As Integer ' This holds the current date column value
Dim cur_aCL_Index As Integer ' This holds the current index of the column name in the allColumnsList
'Dim prior_aCL_Index As Integer ' This holds the current index minus one of the column name in the allColumnsList

Dim AggScreenVar_BT_DEL As String ' This holds the aggregate of the screens as they are collected accross the row for milestone DEL
Dim AggScreenVar_AT_DEL As String ' This holds the aggregate of the screens as they are collected accross the row for milestone DEL
Dim AggScreenVar_BT_OWLD As String ' This holds the aggregate of the screens as they are collected accross the row to OWLD
Dim AggScreenVar_AT_OWLD As String ' This holds the aggregate of the screens as they are collected accross the row to OWLD
Dim AggScreenVar_FCT_OWLD As String ' This holds the aggregate of the screens as they are collected accross the row to OWLD
Dim AggScreenVar_BT_Final As String ' This holds the aggregate of the screens as they are collected accross the row to Final
Dim AggScreenVar_AT_Final As String ' This holds the aggregate of the screens as they are collected accross the row to Final
Dim AggScreenVar_FCT_Final As String ' This holds the aggregate of the screens as they are collected accross the row to Final

Dim ScreenCountsVar_BT_DEL As Integer ' This holds the count of the screens as they are collected accross the row for milestone DEL
Dim ScreenCountsVar_AT_DEL As Integer ' This holds the count of the screens as they are collected accross the row for milestone DEL
Dim ScreenCountsVar_BT_OWLD As Integer ' This holds the count of the screens as they are collected accross the row to OWLD
Dim ScreenCountsVar_AT_OWLD As Integer ' This holds the count of the screens as they are collected accross the row to OWLD
Dim ScreenCountsVar_FCT_OWLD As Integer ' This holds the count of the screens as they are collected accross the row to OWLD
Dim ScreenCountsVar_BT_Final As Integer ' This holds the count of the screens as they are collected accross the row to Final
Dim ScreenCountsVar_AT_Final As Integer ' This holds the count of the screens as they are collected accross the row to Final
Dim ScreenCountsVar_FCT_Final As Integer ' This holds the count of the screens as they are collected accross the row to Final

'These are used on the recordset
Dim curTrial_Card As String ' This is for the update back to the TC_Screen_Agg Table

Dim priorScreenVar_BT_DEL As String ' This holds the prior screen value from the date column or first screen
Dim currentScreenVar_BT_DEL As String ' This holds the current date column value to see if current column is different
Dim pass_BT_DEL As Boolean  ' This skips the aggragating and counting block
Dim priorScreenVar_BT_OWLD As String ' This holds the prior screen value from the date column or first screen
Dim currentScreenVar_BT_OWLD As String ' This holds the current date column value to see if current column is different
Dim pass_BT_OWLD As Boolean  ' This skips the aggragating and counting block
Dim priorScreenVar_BT_Final As String ' This holds the prior screen value from the date column or first screen
Dim currentScreenVar_BT_Final As String ' This holds the current date column value to see if current column is different
Dim pass_BT_Final As Boolean  ' This skips the aggragating and counting block

Dim priorScreenVar_AT_DEL As String ' This holds the prior screen value from the date column or first screen
Dim currentScreenVar_AT_DEL As String ' This holds the current date column value to see if current column is different
Dim pass_AT_DEL As Boolean  ' This skips the aggragating and counting block
Dim priorScreenVar_AT_OWLD As String ' This holds the prior screen value from the date column or first screen
Dim currentScreenVar_AT_OWLD As String ' This holds the current date column value to see if current column is different
Dim pass_AT_OWLD As Boolean  ' This skips the aggragating and counting block
Dim priorScreenVar_AT_Final As String ' This holds the prior screen value from the date column or first screen
Dim currentScreenVar_AT_Final As String ' This holds the current date column value to see if current column is different
Dim pass_AT_Final As Boolean  ' This skips the aggragating and counting block

Dim priorScreenVar_FCT_DEL As String ' This holds the prior screen value from the date column or first screen
Dim currentScreenVar_FCT_DEL As String ' This holds the current date column value to see if current column is different
Dim pass_FCT_DEL As Boolean  ' This skips the aggragating and counting block
Dim priorScreenVar_FCT_OWLD As String ' This holds the prior screen value from the date column or first screen
Dim currentScreenVar_FCT_OWLD As String ' This holds the current date column value to see if current column is different
Dim pass_FCT_OWLD As Boolean  ' This skips the aggragating and counting block
Dim priorScreenVar_FCT_Final As String ' This holds the prior screen value from the date column or first screen
Dim currentScreenVar_FCT_Final As String ' This holds the current date column value to see if current column is different
Dim pass_FCT_Final As Boolean  ' This skips the aggragating and counting block

'Open a recordset to loop through each Record in the recordset
Dim dbs_Read As DAO.Database ' This is the source database table
Dim rst As DAO.Recordset ' or Recordset2?
Dim mySQLstring As String
'Open a recordset to update the TC_Screen_Agg table
Dim dbs_Write As DAO.Database ' This is the write to database table
Dim mySQL_update_BT_Counts As String ' Updating SQL String
Dim mySQL_update_AT_Counts As String ' Updating SQL String
Dim mySQL_update_FCT_Counts As String ' Updating SQL String
Dim mySQL_update_BT_Aggregate As String ' Updating SQL String
Dim mySQL_update_AT_Aggregate As String ' Updating SQL String
Dim mySQL_update_FCT_Aggregate As String ' Updating SQL String

    'Check for All Reports or only Events
    If All_or_Events = "All" Then
        'simple hard typed SQL statement
        mySQLstring = "SELECT All_XX_Screen_Only.* FROM All_XX_Screen_Only INNER JOIN All_TC_Screen_Agg ON All_XX_Screen_Only.Trial_Card = All_TC_Screen_Agg.Trial_Card;"
    ElseIf All_or_Events = "Events" Then
        'simple hard typed SQL statement
        mySQLstring = "SELECT Events_XX_Screen_Only.* FROM Events_XX_Screen_Only INNER JOIN Events_TC_Screen_Agg ON Events_XX_Screen_Only.Trial_Card = Events_TC_Screen_Agg.Trial_Card;"
    Else
        'Un trapped error
        'All_or_Events Global is empty or not expected value
        Debug.Print "Function Build_and_Set_Aggregated_Screen_Lifetime_TCSA() was passed empty or not expected value with GLOBAL All_or_Events:= " & All_or_Events & "."
    End If

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
            curTrial_Card = rst.Fields![Trial_Card]
            
            AggScreenVar_BT_DEL = Empty ' This holds the aggregate of the screens as they are collected accross the row for milestone DEL
            AggScreenVar_AT_DEL = Empty ' This holds the aggregate of the screens as they are collected accross the row for milestone DEL
            AggScreenVar_BT_OWLD = Empty ' This holds the aggregate of the screens as they are collected accross the row to OWLD
            AggScreenVar_AT_OWLD = Empty ' This holds the aggregate of the screens as they are collected accross the row to OWLD
            AggScreenVar_FCT_OWLD = Empty ' This holds the aggregate of the screens as they are collected accross the row to OWLD
            AggScreenVar_BT_Final = Empty ' This holds the aggregate of the screens as they are collected accross the row to Final
            AggScreenVar_AT_Final = Empty ' This holds the aggregate of the screens as they are collected accross the row to Final
            AggScreenVar_FCT_Final = Empty ' This holds the aggregate of the screens as they are collected accross the row to Final
            
            'Screen is a required field, cannot be empty
            ScreenCountsVar_BT_DEL = 1
            ScreenCountsVar_AT_DEL = 1
            ScreenCountsVar_BT_OWLD = 1
            ScreenCountsVar_AT_OWLD = 1
            ScreenCountsVar_FCT_OWLD = 1
            ScreenCountsVar_BT_Final = 1
            ScreenCountsVar_AT_Final = 1
            ScreenCountsVar_FCT_Final = 1
            'Debug.Print ("Cur TC: " & curTrial_Card)
            
            'ScreenCountsVar_BT = rst.Fields(myDateVarList).Value
            
            For Each myDateVarList In allColumnsList
                'Step accross each date column within the current record row
                'Debug.Print ("TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                
                'curDateVal = rst.Fields(myDateVarList).Value
                
                cur_aCL_Index = allColumnsList.IndexOf(myDateVarList, 0)
                'prior_aCL_Index = cur_aCL_Index - 1
                
                'BT Count to DEL
                'Check if the current date value is in BT Event range and add value
                If cur_aCL_Index < allColumnsList_BT Then
                    'I am before the Event I am counting, do nothing
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                    pass_BT_DEL = True
                ElseIf cur_aCL_Index = allColumnsList_BT Then
                    'I am on the first Event column I am counting
                    priorScreenVar_BT_DEL = rst![First_Screening]
                    currentScreenVar_BT_DEL = rst("" & myDateVarList & "") ' rst![ & myDateVarList & ]
                    AggScreenVar_BT_DEL = rst![First_Screening]
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                    pass_BT_DEL = False
                ElseIf cur_aCL_Index < allColumnsList_DEL Then
                    'I am in the Event Range of the date columns
                    priorScreenVar_BT_DEL = rst("" & allColumnsList(cur_aCL_Index - 1) & "")
                    currentScreenVar_BT_DEL = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_BT_DEL = False
                ElseIf cur_aCL_Index = allColumnsList_DEL Then
                    'I am on the last Event column of the date columns
                    priorScreenVar_BT_DEL = rst("" & allColumnsList(cur_aCL_Index - 1) & "")
                    currentScreenVar_BT_DEL = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_BT_DEL = False
                ElseIf cur_aCL_Index > allColumnsList_DEL Then
                    'I am after the Event Range of the date columns, do nothing
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                    pass_BT_DEL = True
                Else
                    Debug.Print ("ERROR with the range of the Event Range of record row. TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                    pass_BT_DEL = True
                End If
                
                If pass_BT_DEL = False Then
                    If currentScreenVar_BT_DEL = "Not Found" _
                        Or currentScreenVar_BT_DEL = "POST BT Trial" _
                        Or currentScreenVar_BT_DEL = "POST AT Trial" _
                        Or currentScreenVar_BT_DEL = "POST FCT Trial" _
                        Or currentScreenVar_BT_DEL = "SPLIT" _
                        Or currentScreenVar_BT_DEL = "X/X" _
                        Or currentScreenVar_BT_DEL = "" _
                        Or currentScreenVar_BT_DEL = Empty _
                        Then
                        'Do not add this to the aggragate screening
                        'Pass
                    ElseIf currentScreenVar_BT_DEL = priorScreenVar_BT_DEL Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("currentScreenVar_BT_DEL = priorScreenVar_BT_DEL : " & currentScreenVar_BT_DEL & " EQUALS " & priorScreenVar_BT_DEL)
                        'Pass
                    ElseIf Right(currentScreenVar_BT_DEL, 2) = Right(AggScreenVar_BT_DEL, 2) Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("Right(currentScreenVar_BT_DEL, 2): " & Right(currentScreenVar_BT_DEL, 2))
                        'Debug.Print ("Right(AggScreenVar_BT_DEL, 2): " & Right(AggScreenVar_BT_DEL, 2))
                        'Pass
                    ElseIf Right(currentScreenVar_BT_DEL, 2) <> Right(AggScreenVar_BT_DEL, 2) Then
                        'The screening has changed, aggragate this screening value
                        'Debug.Print ("currentScreenVar_BT_DEL: " & currentScreenVar_BT_DEL)
                        'Debug.Print ("Right(AggScreenVar_BT_DEL, 2): " & Right(AggScreenVar_BT_DEL, 2))
                        'Debug.Print ("Right(currentScreenVar_BT_DEL, 2): " & Right(currentScreenVar_BT_DEL, 2))
                        AggScreenVar_BT_DEL = AggScreenVar_BT_DEL & "/" & currentScreenVar_BT_DEL 'Or use the assigned rst![" & myDateVarList & "]
                        ScreenCountsVar_BT_DEL = ScreenCountsVar_BT_DEL + 1
                    End If
                ElseIf pass_BT_DEL = True Then
                    'Do not aggragate screenings or counts
                    'Pass
                End If
                
                'BT Count to OWLD
                'Check if the current date value is in BT Event range and add value
                If cur_aCL_Index < allColumnsList_BT Then
                    'I am before the Event I am counting, do nothing
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                    pass_BT_OWLD = True
                ElseIf cur_aCL_Index = allColumnsList_BT Then
                    'I am on the first Event column I am counting
                    priorScreenVar_BT_OWLD = rst![First_Screening]
                    currentScreenVar_BT_OWLD = rst("" & myDateVarList & "") ' rst![ & myDateVarList & ]
                    AggScreenVar_BT_OWLD = rst![First_Screening]
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                    pass_BT_OWLD = False
                ElseIf cur_aCL_Index < allColumnsList_OWLD Then
                    'I am in the Event Range of the date columns
                    priorScreenVar_BT_OWLD = rst("" & allColumnsList(cur_aCL_Index - 1) & "")
                    currentScreenVar_BT_OWLD = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_BT_OWLD = False
                ElseIf cur_aCL_Index = allColumnsList_OWLD Then
                    'I am on the last Event column of the date columns
                    priorScreenVar_BT_OWLD = rst("" & allColumnsList(cur_aCL_Index - 1) & "")
                    currentScreenVar_BT_OWLD = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_BT_OWLD = False
                ElseIf cur_aCL_Index > allColumnsList_OWLD Then
                    'I am after the Event Range of the date columns, do nothing
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                    pass_BT_OWLD = True
                Else
                    Debug.Print ("ERROR with the range of the Event Range of record row. TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                    pass_BT_OWLD = True
                End If
                
                If pass_BT_OWLD = False Then
                    If currentScreenVar_BT_OWLD = "Not Found" _
                        Or currentScreenVar_BT_OWLD = "POST BT Trial" _
                        Or currentScreenVar_BT_OWLD = "POST AT Trial" _
                        Or currentScreenVar_BT_OWLD = "POST FCT Trial" _
                        Or currentScreenVar_BT_OWLD = "SPLIT" _
                        Or currentScreenVar_BT_OWLD = "X/X" _
                        Or currentScreenVar_BT_OWLD = "" _
                        Or currentScreenVar_BT_OWLD = Empty _
                        Then
                        'Do not add this to the aggragate screening
                        'Pass
                    ElseIf currentScreenVar_BT_OWLD = priorScreenVar_BT_OWLD Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("currentScreenVar_BT_OWLD = priorScreenVar_BT_OWLD : " & currentScreenVar_BT_OWLD & " EQUALS " & priorScreenVar_BT_OWLD)
                        'Pass
                    ElseIf Right(currentScreenVar_BT_OWLD, 2) = Right(AggScreenVar_BT_OWLD, 2) Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("Right(currentScreenVar_BT_OWLD, 2): " & Right(currentScreenVar_BT_OWLD, 2))
                        'Debug.Print ("Right(AggScreenVar_BT_OWLD, 2): " & Right(AggScreenVar_BT_OWLD, 2))
                        'Pass
                    ElseIf Right(currentScreenVar_BT_OWLD, 2) <> Right(AggScreenVar_BT_OWLD, 2) Then
                        'The screening has changed, aggragate this screening value
                        'Debug.Print ("currentScreenVar_BT_OWLD: " & currentScreenVar_BT_OWLD)
                        'Debug.Print ("Right(AggScreenVar_BT_OWLD, 2): " & Right(AggScreenVar_BT_OWLD, 2))
                        'Debug.Print ("Right(currentScreenVar_BT_OWLD, 2): " & Right(currentScreenVar_BT_OWLD, 2))
                        AggScreenVar_BT_OWLD = AggScreenVar_BT_OWLD & "/" & currentScreenVar_BT_OWLD 'Or use the assigned rst![" & myDateVarList & "]
                        ScreenCountsVar_BT_OWLD = ScreenCountsVar_BT_OWLD + 1
                    End If
                ElseIf pass_BT_OWLD = True Then
                    'Do not aggragate screenings or counts
                    'Pass
                End If
                
                'BT Count to Final
                'Check if the current date value is in BT Event range and add value
                If cur_aCL_Index < allColumnsList_BT Then
                    'I am before the Event I am counting, do nothing
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                    pass_BT_Final = True
                ElseIf cur_aCL_Index = allColumnsList_BT Then
                    'I am on the first Event column I am counting
                    priorScreenVar_BT_Final = rst![First_Screening]
                    currentScreenVar_BT_Final = rst("" & myDateVarList & "") ' rst![ & myDateVarList & ]
                    AggScreenVar_BT_Final = rst![First_Screening]
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                    pass_BT_Final = False
                ElseIf cur_aCL_Index < allColumnsList_Final Then
                    'I am in the Event Range of the date columns
                    priorScreenVar_BT_Final = rst("" & allColumnsList(cur_aCL_Index - 1) & "")
                    currentScreenVar_BT_Final = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_BT_Final = False
                ElseIf cur_aCL_Index = allColumnsList_Final Then
                    'I am on the last Event column of the date columns
                    priorScreenVar_BT_Final = rst("" & allColumnsList(cur_aCL_Index - 1) & "")
                    currentScreenVar_BT_Final = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_BT_Final = False
                ElseIf cur_aCL_Index > allColumnsList_Final Then
                    'I am after the Event Range of the date columns, do nothing
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                    pass_BT_Final = True
                Else
                    Debug.Print ("ERROR with the range of the Event Range of record row. TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                    pass_BT_Final = True
                End If
                
                If pass_BT_Final = False Then
                    If currentScreenVar_BT_Final = "Not Found" _
                        Or currentScreenVar_BT_Final = "POST BT Trial" _
                        Or currentScreenVar_BT_Final = "POST AT Trial" _
                        Or currentScreenVar_BT_Final = "POST FCT Trial" _
                        Or currentScreenVar_BT_Final = "SPLIT" _
                        Or currentScreenVar_BT_Final = "X/X" _
                        Or currentScreenVar_BT_Final = "" _
                        Or currentScreenVar_BT_Final = Empty _
                        Then
                        'Do not add this to the aggragate screening
                        'Pass
                    ElseIf currentScreenVar_BT_Final = priorScreenVar_BT_Final Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("currentScreenVar_BT_Final = priorScreenVar_BT_Final : " & currentScreenVar_BT_Final & " EQUALS " & priorScreenVar_BT_Final)
                        'Pass
                    ElseIf Right(currentScreenVar_BT_Final, 2) = Right(AggScreenVar_BT_Final, 2) Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("Right(currentScreenVar_BT_Final, 2): " & Right(currentScreenVar_BT_Final, 2))
                        'Debug.Print ("Right(AggScreenVar_BT_Final, 2): " & Right(AggScreenVar_BT_Final, 2))
                        'Pass
                    ElseIf Right(currentScreenVar_BT_Final, 2) <> Right(AggScreenVar_BT_Final, 2) Then
                        'The screening has changed, aggragate this screening value
                        'Debug.Print ("currentScreenVar_BT_Final: " & currentScreenVar_BT_Final)
                        'Debug.Print ("Right(AggScreenVar_BT_Final, 2): " & Right(AggScreenVar_BT_Final, 2))
                        'Debug.Print ("Right(currentScreenVar_BT_Final, 2): " & Right(currentScreenVar_BT_Final, 2))
                        AggScreenVar_BT_Final = AggScreenVar_BT_Final & "/" & currentScreenVar_BT_Final 'Or use the assigned rst![" & myDateVarList & "]
                        ScreenCountsVar_BT_Final = ScreenCountsVar_BT_Final + 1
                    End If
                ElseIf pass_BT_Final = True Then
                    'Do not aggragate screenings or counts
                    'Pass
                End If
            
                'AT Count to DEL
                'Check if the current date value is in AT Event range and add value
                If cur_aCL_Index < allColumnsList_AT Then
                    'I am before the Event I am counting, do nothing
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                    pass_AT_DEL = True
                ElseIf cur_aCL_Index = allColumnsList_AT Then
                    'I am on the first Event column I am counting
                    priorScreenVar_AT_DEL = rst![First_Screening]
                    currentScreenVar_AT_DEL = rst("" & myDateVarList & "") ' rst![ & myDateVarList & ]
                    AggScreenVar_AT_DEL = rst![First_Screening]
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                    pass_AT_DEL = False
                ElseIf cur_aCL_Index < allColumnsList_DEL Then
                    'I am in the Event Range of the date columns
                    priorScreenVar_AT_DEL = rst("" & allColumnsList(cur_aCL_Index - 1) & "")
                    currentScreenVar_AT_DEL = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_AT_DEL = False
                ElseIf cur_aCL_Index = allColumnsList_DEL Then
                    'I am on the last Event column of the date columns
                    priorScreenVar_AT_DEL = rst("" & allColumnsList(cur_aCL_Index - 1) & "")
                    currentScreenVar_AT_DEL = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_AT_DEL = False
                ElseIf cur_aCL_Index > allColumnsList_DEL Then
                    'I am after the Event Range of the date columns, do nothing
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                    pass_AT_DEL = True
                Else
                    Debug.Print ("ERROR with the range of the Event Range of record row. TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                    pass_AT_DEL = True
                End If
                
                If pass_AT_DEL = False Then
                    If currentScreenVar_AT_DEL = "Not Found" _
                        Or currentScreenVar_AT_DEL = "POST BT Trial" _
                        Or currentScreenVar_AT_DEL = "POST AT Trial" _
                        Or currentScreenVar_AT_DEL = "POST FCT Trial" _
                        Or currentScreenVar_AT_DEL = "SPLIT" _
                        Or currentScreenVar_AT_DEL = "X/X" _
                        Or currentScreenVar_AT_DEL = "" _
                        Or currentScreenVar_AT_DEL = Empty _
                        Then
                        'Do not add this to the aggragate screening
                        'Pass
                    ElseIf currentScreenVar_AT_DEL = priorScreenVar_AT_DEL Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("currentScreenVar_AT_DEL = priorScreenVar_AT_DEL : " & currentScreenVar_AT_DEL & " EQUALS " & priorScreenVar_AT_DEL)
                        'Pass
                    ElseIf Right(currentScreenVar_AT_DEL, 2) = Right(AggScreenVar_AT_DEL, 2) Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("Right(currentScreenVar_AT_DEL, 2): " & Right(currentScreenVar_AT_DEL, 2))
                        'Debug.Print ("Right(AggScreenVar_AT_DEL, 2): " & Right(AggScreenVar_AT_DEL, 2))
                        'Pass
                    ElseIf Right(currentScreenVar_AT_DEL, 2) <> Right(AggScreenVar_AT_DEL, 2) Then
                        'The screening has changed, aggragate this screening value
                        'Debug.Print ("currentScreenVar_AT_DEL: " & currentScreenVar_AT_DEL)
                        'Debug.Print ("Right(AggScreenVar_AT_DEL, 2): " & Right(AggScreenVar_AT_DEL, 2))
                        'Debug.Print ("Right(currentScreenVar_AT_DEL, 2): " & Right(currentScreenVar_AT_DEL, 2))
                        AggScreenVar_AT_DEL = AggScreenVar_AT_DEL & "/" & currentScreenVar_AT_DEL 'Or use the assigned rst![" & myDateVarList & "]
                        ScreenCountsVar_AT_DEL = ScreenCountsVar_AT_DEL + 1
                    End If
                ElseIf pass_AT_DEL = True Then
                    'Do not aggragate screenings or counts
                    'Pass
                End If
                
                'AT Count to OWLD
                'Check if the current date value is in AT Event range and add value
                If cur_aCL_Index < allColumnsList_AT Then
                    'I am before the Event I am counting, do nothing
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                    pass_AT_OWLD = True
                ElseIf cur_aCL_Index = allColumnsList_AT Then
                    'I am on the first Event column I am counting
                    priorScreenVar_AT_OWLD = rst![First_Screening]
                    currentScreenVar_AT_OWLD = rst("" & myDateVarList & "") ' rst![ & myDateVarList & ]
                    AggScreenVar_AT_OWLD = rst![First_Screening]
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                    pass_AT_OWLD = False
                ElseIf cur_aCL_Index < allColumnsList_OWLD Then
                    'I am in the Event Range of the date columns
                    priorScreenVar_AT_OWLD = rst("" & allColumnsList(cur_aCL_Index - 1) & "")
                    currentScreenVar_AT_OWLD = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_AT_OWLD = False
                ElseIf cur_aCL_Index = allColumnsList_OWLD Then
                    'I am on the last Event column of the date columns
                    priorScreenVar_AT_OWLD = rst("" & allColumnsList(cur_aCL_Index - 1) & "")
                    currentScreenVar_AT_OWLD = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_AT_OWLD = False
                ElseIf cur_aCL_Index > allColumnsList_OWLD Then
                    'I am after the Event Range of the date columns, do nothing
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                    pass_AT_OWLD = True
                Else
                    Debug.Print ("ERROR with the range of the Event Range of record row. TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                    pass_AT_OWLD = True
                End If
                
                If pass_AT_OWLD = False Then
                    If currentScreenVar_AT_OWLD = "Not Found" _
                        Or currentScreenVar_AT_OWLD = "POST BT Trial" _
                        Or currentScreenVar_AT_OWLD = "POST AT Trial" _
                        Or currentScreenVar_AT_OWLD = "POST FCT Trial" _
                        Or currentScreenVar_AT_OWLD = "SPLIT" _
                        Or currentScreenVar_AT_OWLD = "X/X" _
                        Or currentScreenVar_AT_OWLD = "" _
                        Or currentScreenVar_AT_OWLD = Empty _
                        Then
                        'Do not add this to the aggragate screening
                        'Pass
                    ElseIf currentScreenVar_AT_OWLD = priorScreenVar_AT_OWLD Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("currentScreenVar_AT_OWLD = priorScreenVar_AT_OWLD : " & currentScreenVar_AT_OWLD & " EQUALS " & priorScreenVar_AT_OWLD)
                        'Pass
                    ElseIf Right(currentScreenVar_AT_OWLD, 2) = Right(AggScreenVar_AT_OWLD, 2) Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("Right(currentScreenVar_AT_OWLD, 2): " & Right(currentScreenVar_AT_OWLD, 2))
                        'Debug.Print ("Right(AggScreenVar_AT_OWLD, 2): " & Right(AggScreenVar_AT_OWLD, 2))
                        'Pass
                    ElseIf Right(currentScreenVar_AT_OWLD, 2) <> Right(AggScreenVar_AT_OWLD, 2) Then
                        'The screening has changed, aggragate this screening value
                        'Debug.Print ("currentScreenVar_AT_OWLD: " & currentScreenVar_AT_OWLD)
                        'Debug.Print ("Right(AggScreenVar_AT_OWLD, 2): " & Right(AggScreenVar_AT_OWLD, 2))
                        'Debug.Print ("Right(currentScreenVar_AT_OWLD, 2): " & Right(currentScreenVar_AT_OWLD, 2))
                        AggScreenVar_AT_OWLD = AggScreenVar_AT_OWLD & "/" & currentScreenVar_AT_OWLD 'Or use the assigned rst![" & myDateVarList & "]
                        ScreenCountsVar_AT_OWLD = ScreenCountsVar_AT_OWLD + 1
                    End If
                ElseIf pass_AT_OWLD = True Then
                    'Do not aggragate screenings or counts
                    'Pass
                End If
                
                'AT Count to Final
                'Check if the current date value is in AT Event range and add value
                If cur_aCL_Index < allColumnsList_AT Then
                    'I am before the Event I am counting, do nothing
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                    pass_AT_Final = True
                ElseIf cur_aCL_Index = allColumnsList_AT Then
                    'I am on the first Event column I am counting
                    priorScreenVar_AT_Final = rst![First_Screening]
                    currentScreenVar_AT_Final = rst("" & myDateVarList & "") ' rst![ & myDateVarList & ]
                    AggScreenVar_AT_Final = rst![First_Screening]
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                    pass_AT_Final = False
                ElseIf cur_aCL_Index < allColumnsList_Final Then
                    'I am in the Event Range of the date columns
                    priorScreenVar_AT_Final = rst("" & allColumnsList(cur_aCL_Index - 1) & "")
                    currentScreenVar_AT_Final = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_AT_Final = False
                ElseIf cur_aCL_Index = allColumnsList_Final Then
                    'I am on the last Event column of the date columns
                    priorScreenVar_AT_Final = rst("" & allColumnsList(cur_aCL_Index - 1) & "")
                    currentScreenVar_AT_Final = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_AT_Final = False
                ElseIf cur_aCL_Index > allColumnsList_Final Then
                    'I am after the Event Range of the date columns, do nothing
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                    pass_AT_Final = True
                Else
                    Debug.Print ("ERROR with the range of the Event Range of record row. TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                    pass_AT_Final = True
                End If
                
                If pass_AT_Final = False Then
                    If currentScreenVar_AT_Final = "Not Found" _
                        Or currentScreenVar_AT_Final = "POST BT Trial" _
                        Or currentScreenVar_AT_Final = "POST AT Trial" _
                        Or currentScreenVar_AT_Final = "POST FCT Trial" _
                        Or currentScreenVar_AT_Final = "SPLIT" _
                        Or currentScreenVar_AT_Final = "X/X" _
                        Or currentScreenVar_AT_Final = "" _
                        Or currentScreenVar_AT_Final = Empty _
                        Then
                        'Do not add this to the aggragate screening
                        'Pass
                    ElseIf currentScreenVar_AT_Final = priorScreenVar_AT_Final Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("currentScreenVar_AT_Final = priorScreenVar_AT_Final : " & currentScreenVar_AT_Final & " EQUALS " & priorScreenVar_AT_Final)
                        'Pass
                    ElseIf Right(currentScreenVar_AT_Final, 2) = Right(AggScreenVar_AT_Final, 2) Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("Right(currentScreenVar_AT_Final, 2): " & Right(currentScreenVar_AT_Final, 2))
                        'Debug.Print ("Right(AggScreenVar_AT_Final, 2): " & Right(AggScreenVar_AT_Final, 2))
                        'Pass
                    ElseIf Right(currentScreenVar_AT_Final, 2) <> Right(AggScreenVar_AT_Final, 2) Then
                        'The screening has changed, aggragate this screening value
                        'Debug.Print ("currentScreenVar_AT_Final: " & currentScreenVar_AT_Final)
                        'Debug.Print ("Right(AggScreenVar_AT_Final, 2): " & Right(AggScreenVar_AT_Final, 2))
                        'Debug.Print ("Right(currentScreenVar_AT_Final, 2): " & Right(currentScreenVar_AT_Final, 2))
                        AggScreenVar_AT_Final = AggScreenVar_AT_Final & "/" & currentScreenVar_AT_Final 'Or use the assigned rst![" & myDateVarList & "]
                        ScreenCountsVar_AT_Final = ScreenCountsVar_AT_Final + 1
                    End If
                ElseIf pass_AT_Final = True Then
                    'Do not aggragate screenings or counts
                    'Pass
                End If
                
                'FCT Count to OWLD
                'Check if the current date value is in FCT Event range and add value
                If cur_aCL_Index < allColumnsList_FCT Then
                    'I am before the Event I am counting, do nothing
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                    pass_FCT_OWLD = True
                ElseIf cur_aCL_Index = allColumnsList_FCT Then
                    'I am on the first Event column I am counting
                    priorScreenVar_FCT_OWLD = rst![First_Screening]
                    currentScreenVar_FCT_OWLD = rst("" & myDateVarList & "") ' rst![ & myDateVarList & ]
                    AggScreenVar_FCT_OWLD = rst![First_Screening]
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                    pass_FCT_OWLD = False
                ElseIf cur_aCL_Index < allColumnsList_OWLD Then
                    'I am in the Event Range of the date columns
                    priorScreenVar_FCT_OWLD = rst("" & allColumnsList(cur_aCL_Index - 1) & "")
                    currentScreenVar_FCT_OWLD = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_FCT_OWLD = False
                ElseIf cur_aCL_Index = allColumnsList_OWLD Then
                    'I am on the last Event column of the date columns
                    priorScreenVar_FCT_OWLD = rst("" & allColumnsList(cur_aCL_Index - 1) & "")
                    currentScreenVar_FCT_OWLD = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_FCT_OWLD = False
                ElseIf cur_aCL_Index > allColumnsList_OWLD Then
                    'I am after the Event Range of the date columns, do nothing
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                    pass_FCT_OWLD = True
                Else
                    Debug.Print ("ERROR with the range of the Event Range of record row. TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                    pass_FCT_OWLD = True
                End If
                
                If pass_FCT_OWLD = False Then
                    If currentScreenVar_FCT_OWLD = "Not Found" _
                        Or currentScreenVar_FCT_OWLD = "POST BT Trial" _
                        Or currentScreenVar_FCT_OWLD = "POST AT Trial" _
                        Or currentScreenVar_FCT_OWLD = "POST FCT Trial" _
                        Or currentScreenVar_FCT_OWLD = "SPLIT" _
                        Or currentScreenVar_FCT_OWLD = "X/X" _
                        Or currentScreenVar_FCT_OWLD = "" _
                        Or currentScreenVar_FCT_OWLD = Empty _
                        Then
                        'Do not add this to the aggragate screening
                        'Pass
                    ElseIf currentScreenVar_FCT_OWLD = priorScreenVar_FCT_OWLD Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("currentScreenVar_FCT_OWLD = priorScreenVar_FCT_OWLD : " & currentScreenVar_FCT_OWLD & " EQUALS " & priorScreenVar_FCT_OWLD)
                        'Pass
                    ElseIf Right(currentScreenVar_FCT_OWLD, 2) = Right(AggScreenVar_FCT_OWLD, 2) Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("Right(currentScreenVar_FCT_OWLD, 2): " & Right(currentScreenVar_FCT_OWLD, 2))
                        'Debug.Print ("Right(AggScreenVar_FCT_OWLD, 2): " & Right(AggScreenVar_FCT_OWLD, 2))
                        'Pass
                    ElseIf Right(currentScreenVar_FCT_OWLD, 2) <> Right(AggScreenVar_FCT_OWLD, 2) Then
                        'The screening has changed, aggragate this screening value
                        'Debug.Print ("currentScreenVar_FCT_OWLD: " & currentScreenVar_FCT_OWLD)
                        'Debug.Print ("Right(AggScreenVar_FCT_OWLD, 2): " & Right(AggScreenVar_FCT_OWLD, 2))
                        'Debug.Print ("Right(currentScreenVar_FCT_OWLD, 2): " & Right(currentScreenVar_FCT_OWLD, 2))
                        AggScreenVar_FCT_OWLD = AggScreenVar_FCT_OWLD & "/" & currentScreenVar_FCT_OWLD 'Or use the assigned rst![" & myDateVarList & "]
                        ScreenCountsVar_FCT_OWLD = ScreenCountsVar_FCT_OWLD + 1
                    End If
                ElseIf pass_FCT_OWLD = True Then
                    'Do not aggragate screenings or counts
                    'Pass
                End If
                
                'FCT Count to Final
                'Check if the current date value is in FCT Event range and add value
                If cur_aCL_Index < allColumnsList_FCT Then
                    'I am before the Event I am counting, do nothing
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                    pass_FCT_Final = True
                ElseIf cur_aCL_Index = allColumnsList_FCT Then
                    'I am on the first Event column I am counting
                    priorScreenVar_FCT_Final = rst![First_Screening]
                    currentScreenVar_FCT_Final = rst("" & myDateVarList & "") ' rst![ & myDateVarList & ]
                    AggScreenVar_FCT_Final = rst![First_Screening]
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                    pass_FCT_Final = False
                ElseIf cur_aCL_Index < allColumnsList_Final Then
                    'I am in the Event Range of the date columns
                    priorScreenVar_FCT_Final = rst("" & allColumnsList(cur_aCL_Index - 1) & "")
                    currentScreenVar_FCT_Final = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_FCT_Final = False
                ElseIf cur_aCL_Index = allColumnsList_Final Then
                    'I am on the last Event column of the date columns
                    priorScreenVar_FCT_Final = rst("" & allColumnsList(cur_aCL_Index - 1) & "")
                    currentScreenVar_FCT_Final = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_FCT_Final = False
                ElseIf cur_aCL_Index > allColumnsList_Final Then
                    'I am after the Event Range of the date columns, do nothing
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                    pass_FCT_Final = True
                Else
                    Debug.Print ("ERROR with the range of the Event Range of record row. TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                    pass_FCT_Final = True
                End If
                
                If pass_FCT_Final = False Then
                    If currentScreenVar_FCT_Final = "Not Found" _
                        Or currentScreenVar_FCT_Final = "POST BT Trial" _
                        Or currentScreenVar_FCT_Final = "POST AT Trial" _
                        Or currentScreenVar_FCT_Final = "POST FCT Trial" _
                        Or currentScreenVar_FCT_Final = "SPLIT" _
                        Or currentScreenVar_FCT_Final = "X/X" _
                        Or currentScreenVar_FCT_Final = "" _
                        Or currentScreenVar_FCT_Final = Empty _
                        Then
                        'Do not add this to the aggragate screening
                        'Pass
                    ElseIf currentScreenVar_FCT_Final = priorScreenVar_FCT_Final Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("currentScreenVar_FCT_Final = priorScreenVar_FCT_Final : " & currentScreenVar_FCT_Final & " EQUALS " & priorScreenVar_FCT_Final)
                        'Pass
                    ElseIf Right(currentScreenVar_FCT_Final, 2) = Right(AggScreenVar_FCT_Final, 2) Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("Right(currentScreenVar_FCT_Final, 2): " & Right(currentScreenVar_FCT_Final, 2))
                        'Debug.Print ("Right(AggScreenVar_FCT_Final, 2): " & Right(AggScreenVar_FCT_Final, 2))
                        'Pass
                    ElseIf Right(currentScreenVar_FCT_Final, 2) <> Right(AggScreenVar_FCT_Final, 2) Then
                        'The screening has changed, aggragate this screening value
                        'Debug.Print ("currentScreenVar_FCT_Final: " & currentScreenVar_FCT_Final)
                        'Debug.Print ("Right(AggScreenVar_FCT_Final, 2): " & Right(AggScreenVar_FCT_Final, 2))
                        'Debug.Print ("Right(currentScreenVar_FCT_Final, 2): " & Right(currentScreenVar_FCT_Final, 2))
                        AggScreenVar_FCT_Final = AggScreenVar_FCT_Final & "/" & currentScreenVar_FCT_Final 'Or use the assigned rst![" & myDateVarList & "]
                        ScreenCountsVar_FCT_Final = ScreenCountsVar_FCT_Final + 1
                    End If
                ElseIf pass_FCT_Final = True Then
                    'Do not aggragate screenings or counts
                    'Pass
                End If
                
            Next myDateVarList
            
            'Debug.Print ("Cur TC: " & curTrial_Card)
            'Update the Sparse Matrix BT Event Ranges with the Values
            mySQL_update_BT_Counts = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_counts_BT_to_DEL] = " & ScreenCountsVar_BT_DEL & ", " & CurrentTable & ".[Rescreen_counts_BT_to_OWLD] = " & ScreenCountsVar_BT_OWLD & ", " & CurrentTable & ".[Rescreen_counts_BT_to_Final] = " & ScreenCountsVar_BT_Final & " WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
            mySQL_update_BT_Aggregate = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_aggregates_BT_to_DEL] = '" & AggScreenVar_BT_DEL & "', " & CurrentTable & ".[Rescreen_aggregates_BT_to_OWLD] = '" & AggScreenVar_BT_OWLD & "', " & CurrentTable & ".[Rescreen_aggregates_BT_to_Final] = '" & AggScreenVar_BT_Final & "' WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
            'Debug.Print ("mySQL_update_BT_Counts = " & mySQL_update_BT_Counts)
            'Debug.Print ("mySQL_update_BT_Aggregate = " & mySQL_update_BT_Aggregate)
            
            'Update the Sparse Matrix AT Event Ranges with the Values
            mySQL_update_AT_Counts = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_counts_AT_to_DEL] = " & ScreenCountsVar_AT_DEL & ", " & CurrentTable & ".[Rescreen_counts_AT_to_OWLD] = " & ScreenCountsVar_AT_OWLD & ", " & CurrentTable & ".[Rescreen_counts_AT_to_Final] = " & ScreenCountsVar_AT_Final & " WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
            mySQL_update_AT_Aggregate = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_aggregates_AT_to_DEL] = '" & AggScreenVar_AT_DEL & "', " & CurrentTable & ".[Rescreen_aggregates_AT_to_OWLD] = '" & AggScreenVar_AT_OWLD & "', " & CurrentTable & ".[Rescreen_aggregates_AT_to_Final] = '" & AggScreenVar_AT_Final & "' WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
            'Debug.Print ("mySQL_update_AT_Counts = " & mySQL_update_BT_Counts)
            'Debug.Print ("mySQL_update_AT_Aggregate = " & mySQL_update_BT_Aggregate)
            
            'Update the Sparse Matrix FCT Event Ranges with the Values
            mySQL_update_FCT_Counts = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_counts_FCT_to_OWLD] = " & ScreenCountsVar_FCT_OWLD & ", " & CurrentTable & ".[Rescreen_counts_FCT_to_Final] = " & ScreenCountsVar_FCT_Final & " WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
            mySQL_update_FCT_Aggregate = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_aggregates_FCT_to_OWLD] = '" & AggScreenVar_FCT_OWLD & "', " & CurrentTable & ".[Rescreen_aggregates_FCT_to_Final] = '" & AggScreenVar_FCT_Final & "' WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
            'Debug.Print ("mySQL_update_FCT_Counts = " & mySQL_update_BT_Counts)
            'Debug.Print ("mySQL_update_FCT_Aggregate = " & mySQL_update_BT_Aggregate)
            
            dbs_Write.Execute mySQL_update_BT_Counts
            dbs_Write.Execute mySQL_update_BT_Aggregate
            dbs_Write.Execute mySQL_update_AT_Counts
            dbs_Write.Execute mySQL_update_AT_Aggregate
            dbs_Write.Execute mySQL_update_FCT_Counts
            dbs_Write.Execute mySQL_update_FCT_Aggregate
            
            'When last column of the row record is reached, goto next row record
         
            myDateVarList = Empty
            'Debug.Print ("Next row in Table")
            
            rst.MoveNext
            
        Loop
    
    ElseIf All_or_Events = "Events" Then
        Do While Not rst.EOF
            'Step down each record row
            curTrial_Card = rst.Fields![Trial_Card]
            
            AggScreenVar_BT_DEL = Empty ' This holds the aggregate of the screens as they are collected accross the row for milestone DEL
            AggScreenVar_AT_DEL = Empty ' This holds the aggregate of the screens as they are collected accross the row for milestone DEL
            AggScreenVar_BT_OWLD = Empty ' This holds the aggregate of the screens as they are collected accross the row to OWLD
            AggScreenVar_AT_OWLD = Empty ' This holds the aggregate of the screens as they are collected accross the row to OWLD
            AggScreenVar_FCT_OWLD = Empty ' This holds the aggregate of the screens as they are collected accross the row to OWLD
            AggScreenVar_BT_Final = Empty ' This holds the aggregate of the screens as they are collected accross the row to Final
            AggScreenVar_AT_Final = Empty ' This holds the aggregate of the screens as they are collected accross the row to Final
            AggScreenVar_FCT_Final = Empty ' This holds the aggregate of the screens as they are collected accross the row to Final
            
            'Screen is a required field, cannot be empty
            ScreenCountsVar_BT_DEL = 1
            ScreenCountsVar_AT_DEL = 1
            ScreenCountsVar_BT_OWLD = 1
            ScreenCountsVar_AT_OWLD = 1
            ScreenCountsVar_FCT_OWLD = 1
            ScreenCountsVar_BT_Final = 1
            ScreenCountsVar_AT_Final = 1
            ScreenCountsVar_FCT_Final = 1
            'Debug.Print ("Cur TC: " & curTrial_Card)
            
            'ScreenCountsVar_BT = rst.Fields(myDateVarList).Value
            
            For Each myDateVarList In trialsOnlyList
                'Step accross each date column within the current record row
                'Debug.Print ("TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                
                'curDateVal = rst.Fields(myDateVarList).Value
                
                cur_aCL_Index = trialsOnlyList.IndexOf(myDateVarList, 0)
                'prior_aCL_Index = cur_aCL_Index - 1
                
                'BT Count to DEL
                'Check if the current date value is in BT Event range and add value
                If cur_aCL_Index < trialsOnlyList_BT Then
                    'I am before the Event I am counting, do nothing
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                    pass_BT_DEL = True
                ElseIf cur_aCL_Index = trialsOnlyList_BT Then
                    'I am on the first Event column I am counting
                    priorScreenVar_BT_DEL = rst![First_Screening]
                    currentScreenVar_BT_DEL = rst("" & myDateVarList & "") ' rst![ & myDateVarList & ]
                    AggScreenVar_BT_DEL = rst![First_Screening]
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                    pass_BT_DEL = False
                ElseIf cur_aCL_Index < trialsOnlyList_DEL Then
                    'I am in the Event Range of the date columns
                    priorScreenVar_BT_DEL = rst("" & trialsOnlyList(cur_aCL_Index - 1) & "")
                    currentScreenVar_BT_DEL = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_BT_DEL = False
                ElseIf cur_aCL_Index = trialsOnlyList_DEL Then
                    'I am on the last Event column of the date columns
                    priorScreenVar_BT_DEL = rst("" & trialsOnlyList(cur_aCL_Index - 1) & "")
                    currentScreenVar_BT_DEL = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_BT_DEL = False
                ElseIf cur_aCL_Index > trialsOnlyList_DEL Then
                    'I am after the Event Range of the date columns, do nothing
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                    pass_BT_DEL = True
                Else
                    Debug.Print ("ERROR with the range of the Event Range of record row. TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                    pass_BT_DEL = True
                End If
                
                If pass_BT_DEL = False Then
                    If currentScreenVar_BT_DEL = "Not Found" _
                        Or currentScreenVar_BT_DEL = "POST BT Trial" _
                        Or currentScreenVar_BT_DEL = "POST AT Trial" _
                        Or currentScreenVar_BT_DEL = "POST FCT Trial" _
                        Or currentScreenVar_BT_DEL = "SPLIT" _
                        Or currentScreenVar_BT_DEL = "X/X" _
                        Or currentScreenVar_BT_DEL = "" _
                        Or currentScreenVar_BT_DEL = Empty _
                        Then
                        'Do not add this to the aggragate screening
                        'Pass
                    ElseIf currentScreenVar_BT_DEL = priorScreenVar_BT_DEL Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("currentScreenVar_BT_DEL = priorScreenVar_BT_DEL : " & currentScreenVar_BT_DEL & " EQUALS " & priorScreenVar_BT_DEL)
                        'Pass
                    ElseIf Right(currentScreenVar_BT_DEL, 2) = Right(AggScreenVar_BT_DEL, 2) Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("Right(currentScreenVar_BT_DEL, 2): " & Right(currentScreenVar_BT_DEL, 2))
                        'Debug.Print ("Right(AggScreenVar_BT_DEL, 2): " & Right(AggScreenVar_BT_DEL, 2))
                        'Pass
                    ElseIf Right(currentScreenVar_BT_DEL, 2) <> Right(AggScreenVar_BT_DEL, 2) Then
                        'The screening has changed, aggragate this screening value
                        'Debug.Print ("currentScreenVar_BT_DEL: " & currentScreenVar_BT_DEL)
                        'Debug.Print ("Right(AggScreenVar_BT_DEL, 2): " & Right(AggScreenVar_BT_DEL, 2))
                        'Debug.Print ("Right(currentScreenVar_BT_DEL, 2): " & Right(currentScreenVar_BT_DEL, 2))
                        AggScreenVar_BT_DEL = AggScreenVar_BT_DEL & "/" & currentScreenVar_BT_DEL 'Or use the assigned rst![" & myDateVarList & "]
                        ScreenCountsVar_BT_DEL = ScreenCountsVar_BT_DEL + 1
                    End If
                ElseIf pass_BT_DEL = True Then
                    'Do not aggragate screenings or counts
                    'Pass
                End If
                
                'BT Count to OWLD
                'Check if the current date value is in BT Event range and add value
                If cur_aCL_Index < trialsOnlyList_BT Then
                    'I am before the Event I am counting, do nothing
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                    pass_BT_OWLD = True
                ElseIf cur_aCL_Index = trialsOnlyList_BT Then
                    'I am on the first Event column I am counting
                    priorScreenVar_BT_OWLD = rst![First_Screening]
                    currentScreenVar_BT_OWLD = rst("" & myDateVarList & "") ' rst![ & myDateVarList & ]
                    AggScreenVar_BT_OWLD = rst![First_Screening]
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                    pass_BT_OWLD = False
                ElseIf cur_aCL_Index < trialsOnlyList_OWLD Then
                    'I am in the Event Range of the date columns
                    priorScreenVar_BT_OWLD = rst("" & trialsOnlyList(cur_aCL_Index - 1) & "")
                    currentScreenVar_BT_OWLD = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_BT_OWLD = False
                ElseIf cur_aCL_Index = trialsOnlyList_OWLD Then
                    'I am on the last Event column of the date columns
                    priorScreenVar_BT_OWLD = rst("" & trialsOnlyList(cur_aCL_Index - 1) & "")
                    currentScreenVar_BT_OWLD = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_BT_OWLD = False
                ElseIf cur_aCL_Index > trialsOnlyList_OWLD Then
                    'I am after the Event Range of the date columns, do nothing
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                    pass_BT_OWLD = True
                Else
                    Debug.Print ("ERROR with the range of the Event Range of record row. TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                    pass_BT_OWLD = True
                End If
                
                If pass_BT_OWLD = False Then
                    If currentScreenVar_BT_OWLD = "Not Found" _
                        Or currentScreenVar_BT_OWLD = "POST BT Trial" _
                        Or currentScreenVar_BT_OWLD = "POST AT Trial" _
                        Or currentScreenVar_BT_OWLD = "POST FCT Trial" _
                        Or currentScreenVar_BT_OWLD = "SPLIT" _
                        Or currentScreenVar_BT_OWLD = "X/X" _
                        Or currentScreenVar_BT_OWLD = "" _
                        Or currentScreenVar_BT_OWLD = Empty _
                        Then
                        'Do not add this to the aggragate screening
                        'Pass
                    ElseIf currentScreenVar_BT_OWLD = priorScreenVar_BT_OWLD Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("currentScreenVar_BT_OWLD = priorScreenVar_BT_OWLD : " & currentScreenVar_BT_OWLD & " EQUALS " & priorScreenVar_BT_OWLD)
                        'Pass
                    ElseIf Right(currentScreenVar_BT_OWLD, 2) = Right(AggScreenVar_BT_OWLD, 2) Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("Right(currentScreenVar_BT_OWLD, 2): " & Right(currentScreenVar_BT_OWLD, 2))
                        'Debug.Print ("Right(AggScreenVar_BT_OWLD, 2): " & Right(AggScreenVar_BT_OWLD, 2))
                        'Pass
                    ElseIf Right(currentScreenVar_BT_OWLD, 2) <> Right(AggScreenVar_BT_OWLD, 2) Then
                        'The screening has changed, aggragate this screening value
                        'Debug.Print ("currentScreenVar_BT_OWLD: " & currentScreenVar_BT_OWLD)
                        'Debug.Print ("Right(AggScreenVar_BT_OWLD, 2): " & Right(AggScreenVar_BT_OWLD, 2))
                        'Debug.Print ("Right(currentScreenVar_BT_OWLD, 2): " & Right(currentScreenVar_BT_OWLD, 2))
                        AggScreenVar_BT_OWLD = AggScreenVar_BT_OWLD & "/" & currentScreenVar_BT_OWLD 'Or use the assigned rst![" & myDateVarList & "]
                        ScreenCountsVar_BT_OWLD = ScreenCountsVar_BT_OWLD + 1
                    End If
                ElseIf pass_BT_OWLD = True Then
                    'Do not aggragate screenings or counts
                    'Pass
                End If
                
                'BT Count to Final
                'Check if the current date value is in BT Event range and add value
                If cur_aCL_Index < trialsOnlyList_BT Then
                    'I am before the Event I am counting, do nothing
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                    pass_BT_Final = True
                ElseIf cur_aCL_Index = trialsOnlyList_BT Then
                    'I am on the first Event column I am counting
                    priorScreenVar_BT_Final = rst![First_Screening]
                    currentScreenVar_BT_Final = rst("" & myDateVarList & "") ' rst![ & myDateVarList & ]
                    AggScreenVar_BT_Final = rst![First_Screening]
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                    pass_BT_Final = False
                ElseIf cur_aCL_Index < trialsOnlyList_Final Then
                    'I am in the Event Range of the date columns
                    priorScreenVar_BT_Final = rst("" & trialsOnlyList(cur_aCL_Index - 1) & "")
                    currentScreenVar_BT_Final = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_BT_Final = False
                ElseIf cur_aCL_Index = trialsOnlyList_Final Then
                    'I am on the last Event column of the date columns
                    priorScreenVar_BT_Final = rst("" & trialsOnlyList(cur_aCL_Index - 1) & "")
                    currentScreenVar_BT_Final = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_BT_Final = False
                ElseIf cur_aCL_Index > trialsOnlyList_Final Then
                    'I am after the Event Range of the date columns, do nothing
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                    pass_BT_Final = True
                Else
                    Debug.Print ("ERROR with the range of the Event Range of record row. TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                    pass_BT_Final = True
                End If
                
                If pass_BT_Final = False Then
                    If currentScreenVar_BT_Final = "Not Found" _
                        Or currentScreenVar_BT_Final = "POST BT Trial" _
                        Or currentScreenVar_BT_Final = "POST AT Trial" _
                        Or currentScreenVar_BT_Final = "POST FCT Trial" _
                        Or currentScreenVar_BT_Final = "SPLIT" _
                        Or currentScreenVar_BT_Final = "X/X" _
                        Or currentScreenVar_BT_Final = "" _
                        Or currentScreenVar_BT_Final = Empty _
                        Then
                        'Do not add this to the aggragate screening
                        'Pass
                    ElseIf currentScreenVar_BT_Final = priorScreenVar_BT_Final Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("currentScreenVar_BT_Final = priorScreenVar_BT_Final : " & currentScreenVar_BT_Final & " EQUALS " & priorScreenVar_BT_Final)
                        'Pass
                    ElseIf Right(currentScreenVar_BT_Final, 2) = Right(AggScreenVar_BT_Final, 2) Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("Right(currentScreenVar_BT_Final, 2): " & Right(currentScreenVar_BT_Final, 2))
                        'Debug.Print ("Right(AggScreenVar_BT_Final, 2): " & Right(AggScreenVar_BT_Final, 2))
                        'Pass
                    ElseIf Right(currentScreenVar_BT_Final, 2) <> Right(AggScreenVar_BT_Final, 2) Then
                        'The screening has changed, aggragate this screening value
                        'Debug.Print ("currentScreenVar_BT_Final: " & currentScreenVar_BT_Final)
                        'Debug.Print ("Right(AggScreenVar_BT_Final, 2): " & Right(AggScreenVar_BT_Final, 2))
                        'Debug.Print ("Right(currentScreenVar_BT_Final, 2): " & Right(currentScreenVar_BT_Final, 2))
                        AggScreenVar_BT_Final = AggScreenVar_BT_Final & "/" & currentScreenVar_BT_Final 'Or use the assigned rst![" & myDateVarList & "]
                        ScreenCountsVar_BT_Final = ScreenCountsVar_BT_Final + 1
                    End If
                ElseIf pass_BT_Final = True Then
                    'Do not aggragate screenings or counts
                    'Pass
                End If
            
                'AT Count to DEL
                'Check if the current date value is in AT Event range and add value
                If cur_aCL_Index < trialsOnlyList_AT Then
                    'I am before the Event I am counting, do nothing
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                    pass_AT_DEL = True
                ElseIf cur_aCL_Index = trialsOnlyList_AT Then
                    'I am on the first Event column I am counting
                    priorScreenVar_AT_DEL = rst![First_Screening]
                    currentScreenVar_AT_DEL = rst("" & myDateVarList & "") ' rst![ & myDateVarList & ]
                    AggScreenVar_AT_DEL = rst![First_Screening]
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                    pass_AT_DEL = False
                ElseIf cur_aCL_Index < trialsOnlyList_DEL Then
                    'I am in the Event Range of the date columns
                    priorScreenVar_AT_DEL = rst("" & trialsOnlyList(cur_aCL_Index - 1) & "")
                    currentScreenVar_AT_DEL = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_AT_DEL = False
                ElseIf cur_aCL_Index = trialsOnlyList_DEL Then
                    'I am on the last Event column of the date columns
                    priorScreenVar_AT_DEL = rst("" & trialsOnlyList(cur_aCL_Index - 1) & "")
                    currentScreenVar_AT_DEL = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_AT_DEL = False
                ElseIf cur_aCL_Index > trialsOnlyList_DEL Then
                    'I am after the Event Range of the date columns, do nothing
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                    pass_AT_DEL = True
                Else
                    Debug.Print ("ERROR with the range of the Event Range of record row. TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                    pass_AT_DEL = True
                End If
                
                If pass_AT_DEL = False Then
                    If currentScreenVar_AT_DEL = "Not Found" _
                        Or currentScreenVar_AT_DEL = "POST BT Trial" _
                        Or currentScreenVar_AT_DEL = "POST AT Trial" _
                        Or currentScreenVar_AT_DEL = "POST FCT Trial" _
                        Or currentScreenVar_AT_DEL = "SPLIT" _
                        Or currentScreenVar_AT_DEL = "X/X" _
                        Or currentScreenVar_AT_DEL = "" _
                        Or currentScreenVar_AT_DEL = Empty _
                        Then
                        'Do not add this to the aggragate screening
                        'Pass
                    ElseIf currentScreenVar_AT_DEL = priorScreenVar_AT_DEL Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("currentScreenVar_AT_DEL = priorScreenVar_AT_DEL : " & currentScreenVar_AT_DEL & " EQUALS " & priorScreenVar_AT_DEL)
                        'Pass
                    ElseIf Right(currentScreenVar_AT_DEL, 2) = Right(AggScreenVar_AT_DEL, 2) Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("Right(currentScreenVar_AT_DEL, 2): " & Right(currentScreenVar_AT_DEL, 2))
                        'Debug.Print ("Right(AggScreenVar_AT_DEL, 2): " & Right(AggScreenVar_AT_DEL, 2))
                        'Pass
                    ElseIf Right(currentScreenVar_AT_DEL, 2) <> Right(AggScreenVar_AT_DEL, 2) Then
                        'The screening has changed, aggragate this screening value
                        'Debug.Print ("currentScreenVar_AT_DEL: " & currentScreenVar_AT_DEL)
                        'Debug.Print ("Right(AggScreenVar_AT_DEL, 2): " & Right(AggScreenVar_AT_DEL, 2))
                        'Debug.Print ("Right(currentScreenVar_AT_DEL, 2): " & Right(currentScreenVar_AT_DEL, 2))
                        AggScreenVar_AT_DEL = AggScreenVar_AT_DEL & "/" & currentScreenVar_AT_DEL 'Or use the assigned rst![" & myDateVarList & "]
                        ScreenCountsVar_AT_DEL = ScreenCountsVar_AT_DEL + 1
                    End If
                ElseIf pass_AT_DEL = True Then
                    'Do not aggragate screenings or counts
                    'Pass
                End If
                
                'AT Count to OWLD
                'Check if the current date value is in AT Event range and add value
                If cur_aCL_Index < trialsOnlyList_AT Then
                    'I am before the Event I am counting, do nothing
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                    pass_AT_OWLD = True
                ElseIf cur_aCL_Index = trialsOnlyList_AT Then
                    'I am on the first Event column I am counting
                    priorScreenVar_AT_OWLD = rst![First_Screening]
                    currentScreenVar_AT_OWLD = rst("" & myDateVarList & "") ' rst![ & myDateVarList & ]
                    AggScreenVar_AT_OWLD = rst![First_Screening]
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                    pass_AT_OWLD = False
                ElseIf cur_aCL_Index < trialsOnlyList_OWLD Then
                    'I am in the Event Range of the date columns
                    priorScreenVar_AT_OWLD = rst("" & trialsOnlyList(cur_aCL_Index - 1) & "")
                    currentScreenVar_AT_OWLD = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_AT_OWLD = False
                ElseIf cur_aCL_Index = trialsOnlyList_OWLD Then
                    'I am on the last Event column of the date columns
                    priorScreenVar_AT_OWLD = rst("" & trialsOnlyList(cur_aCL_Index - 1) & "")
                    currentScreenVar_AT_OWLD = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_AT_OWLD = False
                ElseIf cur_aCL_Index > trialsOnlyList_OWLD Then
                    'I am after the Event Range of the date columns, do nothing
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                    pass_AT_OWLD = True
                Else
                    Debug.Print ("ERROR with the range of the Event Range of record row. TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                    pass_AT_OWLD = True
                End If
                
                If pass_AT_OWLD = False Then
                    If currentScreenVar_AT_OWLD = "Not Found" _
                        Or currentScreenVar_AT_OWLD = "POST BT Trial" _
                        Or currentScreenVar_AT_OWLD = "POST AT Trial" _
                        Or currentScreenVar_AT_OWLD = "POST FCT Trial" _
                        Or currentScreenVar_AT_OWLD = "SPLIT" _
                        Or currentScreenVar_AT_OWLD = "X/X" _
                        Or currentScreenVar_AT_OWLD = "" _
                        Or currentScreenVar_AT_OWLD = Empty _
                        Then
                        'Do not add this to the aggragate screening
                        'Pass
                    ElseIf currentScreenVar_AT_OWLD = priorScreenVar_AT_OWLD Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("currentScreenVar_AT_OWLD = priorScreenVar_AT_OWLD : " & currentScreenVar_AT_OWLD & " EQUALS " & priorScreenVar_AT_OWLD)
                        'Pass
                    ElseIf Right(currentScreenVar_AT_OWLD, 2) = Right(AggScreenVar_AT_OWLD, 2) Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("Right(currentScreenVar_AT_OWLD, 2): " & Right(currentScreenVar_AT_OWLD, 2))
                        'Debug.Print ("Right(AggScreenVar_AT_OWLD, 2): " & Right(AggScreenVar_AT_OWLD, 2))
                        'Pass
                    ElseIf Right(currentScreenVar_AT_OWLD, 2) <> Right(AggScreenVar_AT_OWLD, 2) Then
                        'The screening has changed, aggragate this screening value
                        'Debug.Print ("currentScreenVar_AT_OWLD: " & currentScreenVar_AT_OWLD)
                        'Debug.Print ("Right(AggScreenVar_AT_OWLD, 2): " & Right(AggScreenVar_AT_OWLD, 2))
                        'Debug.Print ("Right(currentScreenVar_AT_OWLD, 2): " & Right(currentScreenVar_AT_OWLD, 2))
                        AggScreenVar_AT_OWLD = AggScreenVar_AT_OWLD & "/" & currentScreenVar_AT_OWLD 'Or use the assigned rst![" & myDateVarList & "]
                        ScreenCountsVar_AT_OWLD = ScreenCountsVar_AT_OWLD + 1
                    End If
                ElseIf pass_AT_OWLD = True Then
                    'Do not aggragate screenings or counts
                    'Pass
                End If
                
                'AT Count to Final
                'Check if the current date value is in AT Event range and add value
                If cur_aCL_Index < trialsOnlyList_AT Then
                    'I am before the Event I am counting, do nothing
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                    pass_AT_Final = True
                ElseIf cur_aCL_Index = trialsOnlyList_AT Then
                    'I am on the first Event column I am counting
                    priorScreenVar_AT_Final = rst![First_Screening]
                    currentScreenVar_AT_Final = rst("" & myDateVarList & "") ' rst![ & myDateVarList & ]
                    AggScreenVar_AT_Final = rst![First_Screening]
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                    pass_AT_Final = False
                ElseIf cur_aCL_Index < trialsOnlyList_Final Then
                    'I am in the Event Range of the date columns
                    priorScreenVar_AT_Final = rst("" & trialsOnlyList(cur_aCL_Index - 1) & "")
                    currentScreenVar_AT_Final = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_AT_Final = False
                ElseIf cur_aCL_Index = trialsOnlyList_Final Then
                    'I am on the last Event column of the date columns
                    priorScreenVar_AT_Final = rst("" & trialsOnlyList(cur_aCL_Index - 1) & "")
                    currentScreenVar_AT_Final = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_AT_Final = False
                ElseIf cur_aCL_Index > trialsOnlyList_Final Then
                    'I am after the Event Range of the date columns, do nothing
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                    pass_AT_Final = True
                Else
                    Debug.Print ("ERROR with the range of the Event Range of record row. TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                    pass_AT_Final = True
                End If
                
                If pass_AT_Final = False Then
                    If currentScreenVar_AT_Final = "Not Found" _
                        Or currentScreenVar_AT_Final = "POST BT Trial" _
                        Or currentScreenVar_AT_Final = "POST AT Trial" _
                        Or currentScreenVar_AT_Final = "POST FCT Trial" _
                        Or currentScreenVar_AT_Final = "SPLIT" _
                        Or currentScreenVar_AT_Final = "X/X" _
                        Or currentScreenVar_AT_Final = "" _
                        Or currentScreenVar_AT_Final = Empty _
                        Then
                        'Do not add this to the aggragate screening
                        'Pass
                    ElseIf currentScreenVar_AT_Final = priorScreenVar_AT_Final Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("currentScreenVar_AT_Final = priorScreenVar_AT_Final : " & currentScreenVar_AT_Final & " EQUALS " & priorScreenVar_AT_Final)
                        'Pass
                    ElseIf Right(currentScreenVar_AT_Final, 2) = Right(AggScreenVar_AT_Final, 2) Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("Right(currentScreenVar_AT_Final, 2): " & Right(currentScreenVar_AT_Final, 2))
                        'Debug.Print ("Right(AggScreenVar_AT_Final, 2): " & Right(AggScreenVar_AT_Final, 2))
                        'Pass
                    ElseIf Right(currentScreenVar_AT_Final, 2) <> Right(AggScreenVar_AT_Final, 2) Then
                        'The screening has changed, aggragate this screening value
                        'Debug.Print ("currentScreenVar_AT_Final: " & currentScreenVar_AT_Final)
                        'Debug.Print ("Right(AggScreenVar_AT_Final, 2): " & Right(AggScreenVar_AT_Final, 2))
                        'Debug.Print ("Right(currentScreenVar_AT_Final, 2): " & Right(currentScreenVar_AT_Final, 2))
                        AggScreenVar_AT_Final = AggScreenVar_AT_Final & "/" & currentScreenVar_AT_Final 'Or use the assigned rst![" & myDateVarList & "]
                        ScreenCountsVar_AT_Final = ScreenCountsVar_AT_Final + 1
                    End If
                ElseIf pass_AT_Final = True Then
                    'Do not aggragate screenings or counts
                    'Pass
                End If
                
                'FCT Count to OWLD
                'Check if the current date value is in FCT Event range and add value
                If cur_aCL_Index < trialsOnlyList_FCT Then
                    'I am before the Event I am counting, do nothing
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                    pass_FCT_OWLD = True
                ElseIf cur_aCL_Index = trialsOnlyList_FCT Then
                    'I am on the first Event column I am counting
                    priorScreenVar_FCT_OWLD = rst![First_Screening]
                    currentScreenVar_FCT_OWLD = rst("" & myDateVarList & "") ' rst![ & myDateVarList & ]
                    AggScreenVar_FCT_OWLD = rst![First_Screening]
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                    pass_FCT_OWLD = False
                ElseIf cur_aCL_Index < trialsOnlyList_OWLD Then
                    'I am in the Event Range of the date columns
                    priorScreenVar_FCT_OWLD = rst("" & trialsOnlyList(cur_aCL_Index - 1) & "")
                    currentScreenVar_FCT_OWLD = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_FCT_OWLD = False
                ElseIf cur_aCL_Index = trialsOnlyList_OWLD Then
                    'I am on the last Event column of the date columns
                    priorScreenVar_FCT_OWLD = rst("" & trialsOnlyList(cur_aCL_Index - 1) & "")
                    currentScreenVar_FCT_OWLD = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_FCT_OWLD = False
                ElseIf cur_aCL_Index > trialsOnlyList_OWLD Then
                    'I am after the Event Range of the date columns, do nothing
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                    pass_FCT_OWLD = True
                Else
                    Debug.Print ("ERROR with the range of the Event Range of record row. TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                    pass_FCT_OWLD = True
                End If
                
                If pass_FCT_OWLD = False Then
                    If currentScreenVar_FCT_OWLD = "Not Found" _
                        Or currentScreenVar_FCT_OWLD = "POST BT Trial" _
                        Or currentScreenVar_FCT_OWLD = "POST AT Trial" _
                        Or currentScreenVar_FCT_OWLD = "POST FCT Trial" _
                        Or currentScreenVar_FCT_OWLD = "SPLIT" _
                        Or currentScreenVar_FCT_OWLD = "X/X" _
                        Or currentScreenVar_FCT_OWLD = "" _
                        Or currentScreenVar_FCT_OWLD = Empty _
                        Then
                        'Do not add this to the aggragate screening
                        'Pass
                    ElseIf currentScreenVar_FCT_OWLD = priorScreenVar_FCT_OWLD Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("currentScreenVar_FCT_OWLD = priorScreenVar_FCT_OWLD : " & currentScreenVar_FCT_OWLD & " EQUALS " & priorScreenVar_FCT_OWLD)
                        'Pass
                    ElseIf Right(currentScreenVar_FCT_OWLD, 2) = Right(AggScreenVar_FCT_OWLD, 2) Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("Right(currentScreenVar_FCT_OWLD, 2): " & Right(currentScreenVar_FCT_OWLD, 2))
                        'Debug.Print ("Right(AggScreenVar_FCT_OWLD, 2): " & Right(AggScreenVar_FCT_OWLD, 2))
                        'Pass
                    ElseIf Right(currentScreenVar_FCT_OWLD, 2) <> Right(AggScreenVar_FCT_OWLD, 2) Then
                        'The screening has changed, aggragate this screening value
                        'Debug.Print ("currentScreenVar_FCT_OWLD: " & currentScreenVar_FCT_OWLD)
                        'Debug.Print ("Right(AggScreenVar_FCT_OWLD, 2): " & Right(AggScreenVar_FCT_OWLD, 2))
                        'Debug.Print ("Right(currentScreenVar_FCT_OWLD, 2): " & Right(currentScreenVar_FCT_OWLD, 2))
                        AggScreenVar_FCT_OWLD = AggScreenVar_FCT_OWLD & "/" & currentScreenVar_FCT_OWLD 'Or use the assigned rst![" & myDateVarList & "]
                        ScreenCountsVar_FCT_OWLD = ScreenCountsVar_FCT_OWLD + 1
                    End If
                ElseIf pass_FCT_OWLD = True Then
                    'Do not aggragate screenings or counts
                    'Pass
                End If
                
                'FCT Count to Final
                'Check if the current date value is in FCT Event range and add value
                If cur_aCL_Index < trialsOnlyList_FCT Then
                    'I am before the Event I am counting, do nothing
                    'Debug.Print ("Column before Event Range of record row: " & myDateVarList)
                    pass_FCT_Final = True
                ElseIf cur_aCL_Index = trialsOnlyList_FCT Then
                    'I am on the first Event column I am counting
                    priorScreenVar_FCT_Final = rst![First_Screening]
                    currentScreenVar_FCT_Final = rst("" & myDateVarList & "") ' rst![ & myDateVarList & ]
                    AggScreenVar_FCT_Final = rst![First_Screening]
                    'Debug.Print ("Column starting Event Range of record row: " & myDateVarList)
                    pass_FCT_Final = False
                ElseIf cur_aCL_Index < trialsOnlyList_Final Then
                    'I am in the Event Range of the date columns
                    priorScreenVar_FCT_Final = rst("" & trialsOnlyList(cur_aCL_Index - 1) & "")
                    currentScreenVar_FCT_Final = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_FCT_Final = False
                ElseIf cur_aCL_Index = trialsOnlyList_Final Then
                    'I am on the last Event column of the date columns
                    priorScreenVar_FCT_Final = rst("" & trialsOnlyList(cur_aCL_Index - 1) & "")
                    currentScreenVar_FCT_Final = rst("" & myDateVarList & "")
                    'Debug.Print ("In the range of the Event Range of record row: " & myDateVarList)
                    pass_FCT_Final = False
                ElseIf cur_aCL_Index > trialsOnlyList_Final Then
                    'I am after the Event Range of the date columns, do nothing
                    'Debug.Print ("Out of the range of the Event Range of record row: " & myDateVarList)
                    pass_FCT_Final = True
                Else
                    Debug.Print ("ERROR with the range of the Event Range of record row. TC Num : Date Col = " & curTrial_Card & " : " & myDateVarList & "")
                    pass_FCT_Final = True
                End If
                
                If pass_FCT_Final = False Then
                    If currentScreenVar_FCT_Final = "Not Found" _
                        Or currentScreenVar_FCT_Final = "POST BT Trial" _
                        Or currentScreenVar_FCT_Final = "POST AT Trial" _
                        Or currentScreenVar_FCT_Final = "POST FCT Trial" _
                        Or currentScreenVar_FCT_Final = "SPLIT" _
                        Or currentScreenVar_FCT_Final = "X/X" _
                        Or currentScreenVar_FCT_Final = "" _
                        Or currentScreenVar_FCT_Final = Empty _
                        Then
                        'Do not add this to the aggragate screening
                        'Pass
                    ElseIf currentScreenVar_FCT_Final = priorScreenVar_FCT_Final Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("currentScreenVar_FCT_Final = priorScreenVar_FCT_Final : " & currentScreenVar_FCT_Final & " EQUALS " & priorScreenVar_FCT_Final)
                        'Pass
                    ElseIf Right(currentScreenVar_FCT_Final, 2) = Right(AggScreenVar_FCT_Final, 2) Then
                        'The screening has not changed, do not aggragate this screening
                        'Debug.Print ("Right(currentScreenVar_FCT_Final, 2): " & Right(currentScreenVar_FCT_Final, 2))
                        'Debug.Print ("Right(AggScreenVar_FCT_Final, 2): " & Right(AggScreenVar_FCT_Final, 2))
                        'Pass
                    ElseIf Right(currentScreenVar_FCT_Final, 2) <> Right(AggScreenVar_FCT_Final, 2) Then
                        'The screening has changed, aggragate this screening value
                        'Debug.Print ("currentScreenVar_FCT_Final: " & currentScreenVar_FCT_Final)
                        'Debug.Print ("Right(AggScreenVar_FCT_Final, 2): " & Right(AggScreenVar_FCT_Final, 2))
                        'Debug.Print ("Right(currentScreenVar_FCT_Final, 2): " & Right(currentScreenVar_FCT_Final, 2))
                        AggScreenVar_FCT_Final = AggScreenVar_FCT_Final & "/" & currentScreenVar_FCT_Final 'Or use the assigned rst![" & myDateVarList & "]
                        ScreenCountsVar_FCT_Final = ScreenCountsVar_FCT_Final + 1
                    End If
                ElseIf pass_FCT_Final = True Then
                    'Do not aggragate screenings or counts
                    'Pass
                End If
                
            Next myDateVarList
            
            'Debug.Print ("Cur TC: " & curTrial_Card)
            'Update the Sparse Matrix BT Event Ranges with the Values
            mySQL_update_BT_Counts = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_counts_BT_to_DEL] = " & ScreenCountsVar_BT_DEL & ", " & CurrentTable & ".[Rescreen_counts_BT_to_OWLD] = " & ScreenCountsVar_BT_OWLD & ", " & CurrentTable & ".[Rescreen_counts_BT_to_Final] = " & ScreenCountsVar_BT_Final & " WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
            mySQL_update_BT_Aggregate = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_aggregates_BT_to_DEL] = '" & AggScreenVar_BT_DEL & "', " & CurrentTable & ".[Rescreen_aggregates_BT_to_OWLD] = '" & AggScreenVar_BT_OWLD & "', " & CurrentTable & ".[Rescreen_aggregates_BT_to_Final] = '" & AggScreenVar_BT_Final & "' WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
            'Update the Sparse Matrix AT Event Ranges with the Values
            mySQL_update_AT_Counts = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_counts_AT_to_DEL] = " & ScreenCountsVar_AT_DEL & ", " & CurrentTable & ".[Rescreen_counts_AT_to_OWLD] = " & ScreenCountsVar_AT_OWLD & ", " & CurrentTable & ".[Rescreen_counts_AT_to_Final] = " & ScreenCountsVar_AT_Final & " WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
            mySQL_update_AT_Aggregate = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_aggregates_AT_to_DEL] = '" & AggScreenVar_AT_DEL & "', " & CurrentTable & ".[Rescreen_aggregates_AT_to_OWLD] = '" & AggScreenVar_AT_OWLD & "', " & CurrentTable & ".[Rescreen_aggregates_AT_to_Final] = '" & AggScreenVar_AT_Final & "' WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
            'Update the Sparse Matrix FCT Event Ranges with the Values
            mySQL_update_FCT_Counts = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_counts_FCT_to_OWLD] = " & ScreenCountsVar_FCT_OWLD & ", " & CurrentTable & ".[Rescreen_counts_FCT_to_Final] = " & ScreenCountsVar_FCT_Final & " WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
            mySQL_update_FCT_Aggregate = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Rescreen_aggregates_FCT_to_OWLD] = '" & AggScreenVar_FCT_OWLD & "', " & CurrentTable & ".[Rescreen_aggregates_FCT_to_Final] = '" & AggScreenVar_FCT_Final & "' WHERE " & CurrentTable & ".[Trial_Card] = '" & curTrial_Card & "';"
            'Debug.Print ("mySQL_update = " & mySQL_update)
            dbs_Write.Execute mySQL_update_BT_Counts
            dbs_Write.Execute mySQL_update_BT_Aggregate
            dbs_Write.Execute mySQL_update_AT_Counts
            dbs_Write.Execute mySQL_update_AT_Aggregate
            dbs_Write.Execute mySQL_update_FCT_Counts
            dbs_Write.Execute mySQL_update_FCT_Aggregate
            
            'When last column of the row record is reached, goto next row record
         
            myDateVarList = Empty
            'Debug.Print ("Next row in Table")
            
            rst.MoveNext
            
        Loop
    
    Else
        'Un trapped error
        'All_or_Events Global is empty or not expected value
        Debug.Print "Function Build_and_Set_Aggregated_Screen_and_Counts_TCSA() was passed empty or not expected value with GLOBAL All_or_Events:= " & All_or_Events & "."
    End If
    
    rst.Close
    dbs_Read.Close
    dbs_Write.Close
    
    ' Debug.Print vbCrLf & "The Trial Cards Sparse Matrix Update Query completed." & vbCrLf

End Sub
