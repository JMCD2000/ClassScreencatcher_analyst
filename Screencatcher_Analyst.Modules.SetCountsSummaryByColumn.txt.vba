Option Compare Database
Option Explicit


Sub Set_CS_SM_SummaryCounts()
'This function loops through the "All_Combined_Screenings_SparseMatrix" _
table and Sums the Rescreen Changes by Date Column.
Dim myReadSQLstr As String 'This queries the DB to get the Table RecordSet
Dim rsReadData As DAO.Recordset 'This is the table recordset
Dim rsWriteData As DAO.Recordset 'This is the table recordset
Dim i As Long ' Used as the allColumnsList(i) itterable
Dim j As Long ' Used as the aCL_EventIndexList(j) itterable
Dim myStop As Long ' Used as the upperbound aCL_EventIndexList(j) itterable
Dim myColumnCounter As Long
'Dim myTableVarList As Variant
Dim myEventFilterVar As Variant
Dim DateColumnVal As String
Dim myWhereClause As String
Dim myTableName As String
Dim myWriteRow As String

'Cycle through the Sparse Matrix tables
'For Each myTableVarList In All_SparseMatrixList
'myTableName = myTableVarList

myTableName = "All_Combined_Screenings_SparseMatrix" ' All_SparseMatrixList(0)

'CurrentDb.OpenRecordset(Name:="All_Z_Summary", Type:=dbOpenDynaset, Options:=dbInconsistent, LockEdit:=dbOptimistic)
Set rsWriteData = CurrentDb.OpenRecordset("All_Z_Summary", dbOpenDynaset, dbInconsistent, dbOptimistic)

'Filter the table
For Each myEventFilterVar In EventFilterList
    myWhereClause = myEventFilterVar
    
    'Set date range to step across
    For j = 0 To myStop = ((aCL_EventIndexList.Count) - 1)
        'Debug.Print "myStop = " & ((aCL_EventIndexList.Count) - 1)
        'Step accross the date columns
    ''' This could stop at "aCL_OWLD_index_pos" or at "aCL_Final_index_pos"
        For i = aCL_EventIndexList(j) To aCL_Final_index_pos
            
            DateColumnVal = allColumnsList(i)
            myColumnCounter = 0
            
            'Build SQL string
            'myReadSQLstr = "SELECT * FROM All_Combined_Screenings_SparseMatrix WHERE Final_Sts_A_T <> 'X/X' AND [2017/06/30] = '1' AND (" & myWhereClause & ");"
            myReadSQLstr = "SELECT * FROM " & myTableName & " WHERE Final_Sts_A_T <> 'X/X' AND (([" & DateColumnVal & "] = '1') OR([" & DateColumnVal & "] <> '~~~')) AND (" & myWhereClause & ");"
            'Debug.Print "myReadSQLstr: " & myReadSQLstr
            
            'Run SQL
            'CurrentDb.OpenRecordset(Name:=myReadSQLstr, Type:=dbOpenSnapshot, Options:=dbReadOnly, LockEdit:=)
            Set rsReadData = CurrentDb.OpenRecordset(myReadSQLstr, dbOpenSnapshot, dbReadOnly)
            
            If rsReadData.EOF Then
                myColumnCounter = 0
            Else
                rsReadData.MoveLast
                myColumnCounter = rsReadData.RecordCount
            End If
            'Debug.Print "rsReadData.RecordCount: " & rsReadData.RecordCount
            
            Select Case myEventFilterVar
            Case EventFilterList(0) ' 'This is BT New and Splits
                myWriteRow = "CS_SM_DeltaRescreen_BT_New"
            Case EventFilterList(1) ' 'This is BT Roll and Splits
                myWriteRow = "CS_SM_DeltaRescreen_BT_Roll"
            Case EventFilterList(2) ' 'This is AT New and Splits
                myWriteRow = "CS_SM_DeltaRescreen_AT_New"
            Case EventFilterList(3) ' 'This is AT Roll and Splits
                myWriteRow = "CS_SM_DeltaRescreen_AT_Roll"
            Case EventFilterList(4) ' 'This is FCT New and Splits
                myWriteRow = "CS_SM_DeltaRescreen_FCT_New"
            Case EventFilterList(5) ' 'This is FCT Roll and Splits
                myWriteRow = "CS_SM_DeltaRescreen_FCT_Roll"
            Case EventFilterList(6) ' 'This is All BT/AT/FCT New and Splits
                myWriteRow = "CS_SM_DeltaRescreen_ALL"
            Case EventFilterList(7) ' 'This is INSURV New, Roll and Splits
                myWriteRow = "CS_SM_DeltaRescreen_INSURV"
            Case Else
                'untrapped error, Pass
            End Select
            
            'Write the column sum
            ' With current Recordset, Find first record where Column [Table_Source] equals myWriteRow.Value
            rsWriteData.FindFirst "[Table_Source] = '" & myWriteRow & "'"
            rsWriteData.Edit
            rsWriteData.Fields(DateColumnVal) = myColumnCounter
            rsWriteData.Update
            
            myWriteRow = Empty
            DateColumnVal = Empty
            
            myReadSQLstr = Empty
            rsReadData.Close 'The column count has been written, close the record set
            
        Next i ' next date column
            
        i = Empty
            
    Next j ' next event index to set the date column range starting point
        
    j = Empty 'Done with the column Ranges
    
Next myEventFilterVar 'Next Event and Trial_ID SQL to Filter the current table
    
myEventFilterVar = Empty ' Done with the SQL Filter Statments
rsWriteData.Close 'The table "All_Z_Summary" has been updated with the table "All_Combined_Screenings_SparseMatrix"

'Next myTableVarList 'Cycle through the Sparse Matrix tables

Set rsWriteData = Nothing
Set rsReadData = Nothing

End Sub


Sub Set_All_SparseMatrix_SummaryCounts()
'This function loops through the SparseMatrix Tables and _
Sums the Rescreen Changes by Date Column.
Dim myReadSQLstr As String 'This queries the DB to get the Table RecordSet
Dim rsReadData As DAO.Recordset 'This is the table recordset
Dim rsWriteData As DAO.Recordset 'This is the table recordset
Dim i As Long ' Used as the allColumnsList(i) itterable
Dim j As Long ' Used as the aCL_EventIndexList(j) itterable
Dim myStop As Long ' Used as the upperbound aCL_EventIndexList(j) itterable
Dim myColumnCounter As Long
Dim myTableVarList As Variant
Dim myEventFilterVar As Variant
Dim DateColumnVal As String
Dim myWhereClause As String
Dim myTableName As String
Dim myWriteRow As String

'Cycle through the Sparse Matrix tables
For Each myTableVarList In All_SparseMatrixList
    myTableName = myTableVarList
    'myTableName = "All_Combined_Screenings_SparseMatrix" ' All_SparseMatrixList(0)
        
    'CurrentDb.OpenRecordset(Name:="All_Z_Summary", Type:=dbOpenDynaset, Options:=dbInconsistent, LockEdit:=dbOptimistic)
    Set rsWriteData = CurrentDb.OpenRecordset("All_Z_Summary", dbOpenDynaset, dbInconsistent, dbOptimistic)
    
    'Filter the table
    For Each myEventFilterVar In EventFilterList
        myWhereClause = myEventFilterVar
        
        'Set date range to step across
        For j = 0 To myStop = ((aCL_EventIndexList.Count) - 1)
            'Debug.Print "myStop = " & ((aCL_EventIndexList.Count) - 1)
            'Step accross the date columns
        ''' This could stop at "aCL_OWLD_index_pos" or at "aCL_Final_index_pos"
            For i = aCL_EventIndexList(j) To aCL_Final_index_pos
                
                DateColumnVal = allColumnsList(i)
                myColumnCounter = 0
                
                'Build SQL string
                'myReadSQLstr = "SELECT * FROM All_Combined_Screenings_SparseMatrix WHERE Final_Sts_A_T <> 'X/X' AND [2017/06/30] = '1' AND (" & myWhereClause & ");"
                myReadSQLstr = "SELECT * FROM " & myTableName & " WHERE Final_Sts_A_T <> 'X/X' AND (([" & DateColumnVal & "] = '1') OR([" & DateColumnVal & "] <> '~~~')) AND (" & myWhereClause & ");"
                'Debug.Print "myReadSQLstr: " & myReadSQLstr
                
                'Run SQL
                'CurrentDb.OpenRecordset(Name:=myReadSQLstr, Type:=dbOpenSnapshot, Options:=dbReadOnly, LockEdit:=)
                Set rsReadData = CurrentDb.OpenRecordset(myReadSQLstr, dbOpenSnapshot, dbReadOnly)
                
                If rsReadData.EOF Then
                    myColumnCounter = 0
                Else
                    rsReadData.MoveLast
                    myColumnCounter = rsReadData.RecordCount
                End If
                'Debug.Print "rsReadData.RecordCount: " & rsReadData.RecordCount
                
                If myTableVarList = "All_Combined_Screenings_SparseMatrix" Then
                    Select Case myEventFilterVar
                    Case EventFilterList(0) ' 'This is BT New and Splits
                        myWriteRow = "CS_SM_DeltaRescreen_BT_New"
                    Case EventFilterList(1) ' 'This is BT Roll and Splits
                        myWriteRow = "CS_SM_DeltaRescreen_BT_Roll"
                    Case EventFilterList(2) ' 'This is AT New and Splits
                        myWriteRow = "CS_SM_DeltaRescreen_AT_New"
                    Case EventFilterList(3) ' 'This is AT Roll and Splits
                        myWriteRow = "CS_SM_DeltaRescreen_AT_Roll"
                    Case EventFilterList(4) ' 'This is FCT New and Splits
                        myWriteRow = "CS_SM_DeltaRescreen_FCT_New"
                    Case EventFilterList(5) ' 'This is FCT Roll and Splits
                        myWriteRow = "CS_SM_DeltaRescreen_FCT_Roll"
                    Case EventFilterList(6) ' 'This is All BT/AT/FCT New and Splits
                        myWriteRow = "CS_SM_DeltaRescreen_ALL"
                    Case EventFilterList(7) ' 'This is INSURV New, Roll and Splits
                        myWriteRow = "CS_SM_DeltaRescreen_INSURV"
                    Case Else
                        'untrapped error, Pass
                    End Select
                
                ElseIf myTableVarList = "All_Screenings_Only_SparseMatrix" Then
                    Select Case myEventFilterVar
                    Case EventFilterList(0) ' 'This is BT New and Splits
                        myWriteRow = "SO_SM_DeltaRescreen_BT_New"
                    Case EventFilterList(1) ' 'This is BT Roll and Splits
                        myWriteRow = "SO_SM_DeltaRescreen_BT_Roll"
                    Case EventFilterList(2) ' 'This is AT New and Splits
                        myWriteRow = "SO_SM_DeltaRescreen_AT_New"
                    Case EventFilterList(3) ' 'This is AT Roll and Splits
                        myWriteRow = "SO_SM_DeltaRescreen_AT_Roll"
                    Case EventFilterList(4) ' 'This is FCT New and Splits
                        myWriteRow = "SO_SM_DeltaRescreen_FCT_New"
                    Case EventFilterList(5) ' 'This is FCT Roll and Splits
                        myWriteRow = "SO_SM_DeltaRescreen_FCT_Roll"
                    Case EventFilterList(6) ' 'This is All BT/AT/FCT New and Splits
                        myWriteRow = "SO_SM_DeltaRescreen_ALL"
                    Case EventFilterList(7) ' 'This is INSURV New, Roll and Splits
                        myWriteRow = "SO_SM_DeltaRescreen_INSURV"
                    Case Else
                        'untrapped error, Pass
                    End Select
                
                ElseIf myTableVarList = "All_XX_Screen_Only_SparseMatrix" Then
                    Select Case myEventFilterVar
                    Case EventFilterList(0) ' 'This is BT New and Splits
                        myWriteRow = "XXSO_SM_DeltaRescreen_BT_New"
                    Case EventFilterList(1) ' 'This is BT Roll and Splits
                        myWriteRow = "XXSO_SM_DeltaRescreen_BT_Roll"
                    Case EventFilterList(2) ' 'This is AT New and Splits
                        myWriteRow = "XXSO_SM_DeltaRescreen_AT_New"
                    Case EventFilterList(3) ' 'This is AT Roll and Splits
                        myWriteRow = "XXSO_SM_DeltaRescreen_AT_Roll"
                    Case EventFilterList(4) ' 'This is FCT New and Splits
                        myWriteRow = "XXSO_SM_DeltaRescreen_FCT_New"
                    Case EventFilterList(5) ' 'This is FCT Roll and Splits
                        myWriteRow = "XXSO_SM_DeltaRescreen_FCT_Roll"
                    Case EventFilterList(6) ' 'This is All BT/AT/FCT New and Splits
                        myWriteRow = "XXSO_SM_DeltaRescreen_ALL"
                    Case EventFilterList(7) ' 'This is INSURV New, Roll and Splits
                        myWriteRow = "XXSO_SM_DeltaRescreen_INSURV"
                    Case Else
                        'untrapped error, Pass
                    End Select
                Else
                    'Untrapped Error
                End If
                
                'Write the column sum
                ' With current Recordset, Find first record where Column [Table_Source] equals myWriteRow.Value
                rsWriteData.FindFirst "[Table_Source] = '" & myWriteRow & "'"
                rsWriteData.Edit
                rsWriteData.Fields(DateColumnVal) = myColumnCounter
                rsWriteData.Update
                
                myWriteRow = Empty
                DateColumnVal = Empty
                
                myReadSQLstr = Empty
                rsReadData.Close 'The column count has been written, close the record set
                
                'Debug.Print ("          next i : " & i)
            Next i ' next date column
                
            i = Empty
            'Debug.Print ("      next j : " & j)
        Next j ' next event index to set the date column range starting point
            
        j = Empty 'Done with the column Ranges
        'Debug.Print ("   next myEventFilterVar : " & myEventFilterVar)
    Next myEventFilterVar 'Next Event and Trial_ID SQL to Filter the current table
        
    myEventFilterVar = Empty ' Done with the SQL Filter Statments
    rsWriteData.Close 'The table "All_Z_Summary" has been updated with the table "All_Combined_Screenings_SparseMatrix"
    'Debug.Print ("next myTableVarList : " & myTableVarList)
Next myTableVarList 'Cycle through the Sparse Matrix tables

myTableVarList = Empty
Set rsWriteData = Nothing
Set rsReadData = Nothing

Debug.Print ("Finished : Sub Set_All_SparseMatrix_SummaryCounts()")

End Sub


Sub SetRunningDays()
'This function loops through the Dates and _
computes running days by Date Column
Dim rsWriteData As DAO.Recordset 'This is the table recordset being written to
Dim d As Variant ' d as in date
Dim j As Variant ' Used as the aCL_EventIndexList(j) itterable
Dim DateColumnDiff As Long ' current column - prior column
Dim dateIndex As Long ' This is the index to be able to do the date diff
Dim myMM_Cur As String ' This is the current column month value
Dim myDD_Cur As String ' This is the current column day value
Dim myYYYY_Cur As String ' This is the current column year value
Dim myMM_Prior As String ' This is the prior column month value
Dim myDD_Prior As String ' This is the prior column day value
Dim myYYYY_Prior As String ' This is the prior column year value
Dim varCurDateCol  As String ' Month/Day/Year for use in stringfunctions and slicing
Dim varPriorDateCol  As String ' Month/Day/Year for use in stringfunctions and slicing
Dim myCurDateCol  As Date ' Month/Day/Year for use in datefunctions
Dim myPriorDateCol  As Date ' Month/Day/Year for use in datefunctions
Dim DateColumnVal As String
Dim myWriteRow1 As String ' This is the BT/AT/FCT Events
Dim myWriteRow2 As String ' This is the ALL and INSURV Events
Dim useWriteRow2 As Boolean ' This is a selector to avoid table errors when myWriteRow2 is empty

'CurrentDb.OpenRecordset(Name:="All_Z_Summary", Type:=dbOpenDynaset, Options:=dbInconsistent, LockEdit:=dbOptimistic)
Set rsWriteData = CurrentDb.OpenRecordset("All_Z_Summary", dbOpenDynaset, dbInconsistent, dbOptimistic)

'Set date range to step across
For Each j In aCL_EventIndexList
    'Debug.Print ("j index : " & j)
    'Debug.Print ("allColumnsList(j) " & allColumnsList(j))
    
    dateIndex = 0 'This is set and reset outside of the date columns
    
    'Step accross the date columns
    ''' This could stop at "aCL_OWLD_index_pos" or at "aCL_Final_index_pos"
    For Each d In allColumnsList
        'Set the string value of the list to the date column varible
        DateColumnVal = d
            
        If d = allColumnsList(j) Then
            'If d equals the starting column, difference value is zero
            'd = 2017/06/30 or allColumnsList(0)
            DateColumnDiff = 0 'First column has no prior
            'Debug.Print ("DateColumnDiff: " & DateColumnDiff & " CurCol:" & DateColumnVal)
        ElseIf d < allColumnsList(j) Then
            'If d is less than or left of starting column, difference value is zero
            'Could make this a negitive number, but that has not been used as a metric
            'd < 2017/08/18 or allColumnsList(7) = myATvar
            DateColumnDiff = 0
            'Debug.Print ("DateColumnDiff: " & DateColumnDiff & " CurCol:" & DateColumnVal & " < " & myATvar)
        ElseIf d > allColumnsList(j) Then
            'If d is less than or left of starting column, difference value is zero
            'd > 2017/08/18 or allColumnsList(7) = myATvar
            'get column string date values
            varCurDateCol = allColumnsList(dateIndex) 'YYYY/MM/DD 2017/07/04
            varPriorDateCol = allColumnsList(j) 'YYYY/MM/DD 2017/06/30
            'parse out date components
            'current date column
            myMM_Cur = Mid(varCurDateCol, 6, 2) 'yyyy/mm/dd Current date column' = Format(Now, "mm/dd/yyyy")
            myDD_Cur = Mid(varCurDateCol, 9, 2) 'yyyy/mm/dd Current date column' = Format(Now, "mm/dd/yyyy")
            myYYYY_Cur = Mid(varCurDateCol, 1, 4) 'yyyy/mm/dd Current date column' = Format(Now, "mm/dd/yyyy")
            myCurDateCol = DateValue(myMM_Cur & "/" & myDD_Cur & "/" & myYYYY_Cur) 'Current date column' = Format(Now, "mm/dd/yyyy")
            'prior date column
            myMM_Prior = Mid(varPriorDateCol, 6, 2) 'yyyy/mm/dd Prior date column' = Format(Now, "mm/dd/yyyy")
            myDD_Prior = Mid(varPriorDateCol, 9, 2) 'yyyy/mm/dd Prior date column' = Format(Now, "mm/dd/yyyy")
            myYYYY_Prior = Mid(varPriorDateCol, 1, 4) 'yyyy/mm/dd Prior date column' = Format(Now, "mm/dd/yyyy")
            myPriorDateCol = DateValue(myMM_Prior & "/" & myDD_Prior & "/" & myYYYY_Prior) 'Prior date column' = Format(Now, "mm/dd/yyyy")
            'get date diff
            DateColumnDiff = DateDiff("d", myPriorDateCol, myCurDateCol)
            'Debug.Print ("DateColumnDiff: " & DateColumnDiff & " CurCol:" & myCurDateCol & " PriorCol:" & myPriorDateCol)
        Else
            'Untrapped Error, Pass
        End If
        
        '[Table_Source].Value = "Date_Diff_from_Prior" 'days delta between date columns
        '[Table_Source].Value = "DaysPost_ALL" 'increasing days from BT
        '[Table_Source].Value = "DaysPost_AT" 'increasing days from AT
        '[Table_Source].Value = "DaysPost_BT" 'increasing days from BT
        '[Table_Source].Value = "DaysPost_FCT" 'increasing days from FCT
        '[Table_Source].Value = "DaysPost_INSURV" 'increasing days from AT
        
        'Default away from a possible error
        useWriteRow2 = False
        
        Select Case j
            Case aCL_EventIndexList(0) ' This is BT and ALL
                '[Table_Source].Value = "DaysPost_BT"
                '[Table_Source].Value = "DaysPost_ALL"
                myWriteRow1 = "DaysPost_BT"
                useWriteRow2 = True
                myWriteRow2 = "DaysPost_ALL"
            Case aCL_EventIndexList(1) ' This is AT and INSURV
                '[Table_Source].Value = "DaysPost_AT"
                '[Table_Source].Value = "DaysPost_INSURV"
                myWriteRow1 = "DaysPost_AT"
                useWriteRow2 = True
                myWriteRow2 = "DaysPost_INSURV"
            Case aCL_EventIndexList(2) ' This is FCT
                '[Table_Source].Value = "DaysPost_FCT"
                myWriteRow1 = "DaysPost_FCT"
                useWriteRow2 = False
                myWriteRow2 = Empty
            Case Else
                'untrapped error, Pass
        End Select
        
        'Set the string value of the list to the date column varible
        DateColumnVal = allColumnsList(dateIndex)
        
        'Write cumulative days to row
        rsWriteData.FindFirst "[Table_Source] = '" & myWriteRow1 & "'"
        'Write the column sum
        rsWriteData.Edit
        rsWriteData.Fields(DateColumnVal) = DateColumnDiff
        rsWriteData.Update
        
        If useWriteRow2 = True Then
            'Write cumulative days to row
            rsWriteData.FindFirst "[Table_Source] = '" & myWriteRow2 & "'"
            'Write the column sum
            rsWriteData.Edit
            rsWriteData.Fields(DateColumnVal) = DateColumnDiff
            rsWriteData.Update
        Else
            'Assumed is False, Pass
        End If
        
        'Increment the date index
        dateIndex = dateIndex + 1
        
        'Empty the varibles
        myMM_Cur = Empty
        myDD_Cur = Empty
        myYYYY_Cur = Empty
        myCurDateCol = Empty
        myMM_Prior = Empty
        myDD_Prior = Empty
        myYYYY_Prior = Empty
        myPriorDateCol = Empty
        DateColumnDiff = Empty
        myWriteRow1 = Empty
        myWriteRow2 = Empty
        'Default away from a possible error
        'useWriteRow2 = False
        
        'Debug.Print ("      next d : " & d)
    Next d
    
    'Debug.Print ("     next j : " & j)
Next j

'Close the connection
rsWriteData.Close
'Empty the connection Object
Set rsWriteData = Nothing

Debug.Print ("Finished : Sub SetRunningDays()")

End Sub


Sub SetDateDifferences()
'This function loops through the Dates and _
computes difference in days by Date Column
Dim rsWriteData As DAO.Recordset 'This is the table recordset being written to
Dim d As Variant ' d as in date
Dim DateColumnDiff As Long ' current column - prior column
Dim dateIndex As Long ' This is the index to be able to do the date diff
Dim myMM_Cur As String ' This is the current column month value
Dim myDD_Cur As String ' This is the current column day value
Dim myYYYY_Cur As String ' This is the current column year value
Dim myMM_Prior As String ' This is the prior column month value
Dim myDD_Prior As String ' This is the prior column day value
Dim myYYYY_Prior As String ' This is the prior column year value
Dim varCurDateCol  As String ' Month/Day/Year for use in stringfunctions and slicing
Dim varPriorDateCol  As String ' Month/Day/Year for use in stringfunctions and slicing
Dim myCurDateCol  As Date ' Month/Day/Year for use in datefunctions
Dim myPriorDateCol  As Date ' Month/Day/Year for use in datefunctions
Dim DateColumnVal As String
Dim myWriteRow As String '[Table_Source].Value = "Date_Diff_from_Prior" 'days delta between date columns

'CurrentDb.OpenRecordset(Name:="All_Z_Summary", Type:=dbOpenDynaset, Options:=dbInconsistent, LockEdit:=dbOptimistic)
Set rsWriteData = CurrentDb.OpenRecordset("All_Z_Summary", dbOpenDynaset, dbInconsistent, dbOptimistic)

'Step accross all of the date columns to _
get the difference in days between last _
report and current report
dateIndex = 0 'This is the first column
For Each d In allColumnsList
    'Set the string value of the list to the date column varible
    DateColumnVal = d
        
    If d = allColumnsList(0) Then
        'If d equals the starting column, difference value is zero
        'd = 2017/06/30 or allColumnsList(0)
        DateColumnDiff = 0 'First column has no prior
        'Debug.Print ("DateColumnDiff: " & DateColumnDiff & " CurCol:" & DateColumnVal)
    ElseIf d < allColumnsList(0) Then
        'If d is less than or left of starting column, difference value is zero
        'Could make this a negitive number, but that has not been used as a metric
        'd < 2017/08/18 or allColumnsList(7) = myATvar
        DateColumnDiff = 0
        'Debug.Print ("DateColumnDiff: " & DateColumnDiff & " CurCol:" & DateColumnVal & " < " & myATvar)
    ElseIf d > allColumnsList(0) Then
        'If d is less than or left of starting column, difference value is zero
        'd > 2017/08/18 or allColumnsList(7) = myATvar
        'get column string date values
        varCurDateCol = allColumnsList(dateIndex) 'YYYY/MM/DD 2017/07/04
        varPriorDateCol = allColumnsList(dateIndex - 1) 'YYYY/MM/DD 2017/06/30
        'parse out date components
        'current date column
        myMM_Cur = Mid(varCurDateCol, 6, 2) 'yyyy/mm/dd Current date column' = Format(Now, "mm/dd/yyyy")
        myDD_Cur = Mid(varCurDateCol, 9, 2) 'yyyy/mm/dd Current date column' = Format(Now, "mm/dd/yyyy")
        myYYYY_Cur = Mid(varCurDateCol, 1, 4) 'yyyy/mm/dd Current date column' = Format(Now, "mm/dd/yyyy")
        myCurDateCol = DateValue(myMM_Cur & "/" & myDD_Cur & "/" & myYYYY_Cur) 'Current date column' = Format(Now, "mm/dd/yyyy")
        'prior date column
        myMM_Prior = Mid(varPriorDateCol, 6, 2) 'yyyy/mm/dd Prior date column' = Format(Now, "mm/dd/yyyy")
        myDD_Prior = Mid(varPriorDateCol, 9, 2) 'yyyy/mm/dd Prior date column' = Format(Now, "mm/dd/yyyy")
        myYYYY_Prior = Mid(varPriorDateCol, 1, 4) 'yyyy/mm/dd Prior date column' = Format(Now, "mm/dd/yyyy")
        myPriorDateCol = DateValue(myMM_Prior & "/" & myDD_Prior & "/" & myYYYY_Prior) 'Prior date column' = Format(Now, "mm/dd/yyyy")
        'get date diff
        DateColumnDiff = DateDiff("d", myPriorDateCol, myCurDateCol)
        'Debug.Print ("DateColumnDiff: " & DateColumnDiff & " CurCol:" & myCurDateCol & " PriorCol:" & myPriorDateCol)
    Else
        'Untrapped Error, Pass
    End If
    
    'Set the string value of the list to the date column varible
    DateColumnVal = allColumnsList(dateIndex)
    '[Table_Source].Value = "Date_Diff_from_Prior" 'days delta between date columns
    myWriteRow = "Date_Diff_from_Prior"
    
    'Write cumulative days to row
    rsWriteData.FindFirst "[Table_Source] = '" & myWriteRow & "'"
    'Write the column sum
    rsWriteData.Edit
    rsWriteData.Fields(DateColumnVal) = DateColumnDiff
    rsWriteData.Update
        
    'Increment the date index
    dateIndex = dateIndex + 1
    
    'Empty the varibles
    myMM_Cur = Empty
    myDD_Cur = Empty
    myYYYY_Cur = Empty
    myCurDateCol = Empty
    myMM_Prior = Empty
    myDD_Prior = Empty
    myYYYY_Prior = Empty
    myPriorDateCol = Empty
    DateColumnDiff = Empty
    myWriteRow = Empty
    
    'Debug.Print ("      next d : " & d)
    Next d
    
'Close the connection
rsWriteData.Close
'Empty the connection Object
Set rsWriteData = Nothing

'Debug.Print ("Finished : Sub SetDateDifferences()")

End Sub
