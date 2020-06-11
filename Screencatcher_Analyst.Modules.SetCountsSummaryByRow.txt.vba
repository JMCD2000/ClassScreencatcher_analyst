Option Compare Database
Option Explicit


Sub Set_CS_SM_ReScreenCounts()
'This function loops through the "All_Combined_Screenings_SparseMatrix" _
table and Sums the Rescreen Changes across the record row by event window.
Dim myReadSQLstr As String 'This queries the DB to get the Table RecordSet
Dim rsReadData As DAO.Recordset 'This is the table recordset
Dim rsWriteData As DAO.Recordset 'This is the table recordset
Dim i As Long ' Used as the allColumnsList(i) itterable
Dim myColumnCounter As Long
Dim DateColumnVal As String
Dim myTableName As String
Dim myWriteColumn As String
Dim myTC_Number As String 'this is used in the seek method

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder
'Call the SetListAndVars_Summary Module
AddingToMySummaryDateLists 'Run the list builder

'For Each myTableName In
myTableName = "All_Combined_Screenings_SparseMatrix" ' All_SparseMatrixList(0)

'CurrentDb.OpenRecordset(Name:="All_Z_Summary", Type:=dbOpenDynaset, Options:=dbInconsistent, LockEdit:=dbOptimistic)
Set rsWriteData = CurrentDb.OpenRecordset("All_Combined_Screenings_SparseMatrix", dbOpenDynaset, dbInconsistent, dbOptimistic)
'Move cursor to first row, this will be used to itterate through all the rows in order
rsWriteData.MoveFirst
'myTC_Number = rsWriteData.Fields(Trial_Card)
'Debug.Print ("myTC_Number: " & myTC_Number)

    While (Not rsWriteData.EOF)
        myColumnCounter = 0 'This collects the field count of the selected row
        myTC_Number = rsWriteData.Fields("Trial_Card")
        Debug.Print ("myTC_Number: " & myTC_Number)
        
        For i = aCL_EventIndexList(0) To aCL_OWLD_index_pos
            DateColumnVal = allColumnsList(i)
            
            'Build SQL string
            'myReadSQLstr = "SELECT * FROM All_Combined_Screenings_SparseMatrix WHERE Final_Sts_A_T <> 'X/X' AND [2017/06/30] = '1' AND (" & myWhereClause & ");"
            myReadSQLstr = "SELECT " & DateColumnVal & " FROM " & myTableName & " WHERE ([" & DateColumnVal & "] = '1') AND ([Trial_Card] = '" & myTC_Number & "');"
            'Debug.Print "myReadSQLstr: " & myReadSQLstr
            
            'Run SQL
            'CurrentDb.OpenRecordset(Name:=myReadSQLstr, Type:=dbOpenSnapshot, Options:=dbReadOnly, LockEdit:=)
            Set rsReadData = CurrentDb.OpenRecordset(myReadSQLstr, dbOpenSnapshot, dbReadOnly)
            
            If rsReadData.EOF Then
                'date column doesn't equal 1, no record returned
                'Do nothing, pass
            Else
                rsReadData.MoveLast
                myColumnCounter = myColumnCounter + 1
            End If
            
            'Debug.Print "rsReadData.RecordCount: " & rsReadData.RecordCount
            'Debug.Print "myColumnCounter: " & myColumnCounter
            
            myReadSQLstr = Empty
            rsReadData.Close 'The column has been counted, close the record set
            
        Next i ' next date column
            
        Debug.Print "myColumnCounter: " & myColumnCounter
        myWriteColumn = "Rescreen_counts_BT_to_OWLD"
        
        'Write the row sum
        rsWriteData.Edit
        rsWriteData.Fields(myWriteColumn) = myColumnCounter
        rsWriteData.Update
        
        myWriteColumn = Empty
        myColumnCounter = Empty
        
        i = Empty
        rsWriteData.MoveNext
            
    Wend 'While end loop
    
rsWriteData.Close
Set rsWriteData = Nothing
Set rsReadData = Nothing
Debug.Print ("Finished!")

End Sub


Sub Set_All_SparseMatrix_ReScreenCounts()
'This function loops through the "All_Combined_Screenings_SparseMatrix" _
table and Sums the Rescreen Changes across the record row by event window.
Dim myReadSQLstr As String 'This queries the DB to get the Table RecordSet
Dim rsReadData As DAO.Recordset 'This is the table recordset
Dim rsWriteData As DAO.Recordset 'This is the table recordset
Dim i As Long ' Used as the allColumnsList(i) itterable
Dim myColumnCounter As Long
Dim DateColumnVal As String
Dim myTableVarList As Variant 'Loop control
Dim myTableName As String ' current table
Dim myWriteColumn As String
Dim myTC_Number As String 'this is used in the seek method

'Call the SetListAndVars Module
AddingToMyDateLists 'Run the list builder
'Call the SetListAndVars_Summary Module
AddingToMySummaryDateLists 'Run the list builder

'Cycle through the Sparse Matrix tables
For Each myTableVarList In All_SparseMatrixList

    'For Each myTableName In
    myTableName = myTableVarList ' All_SparseMatrixList(0)
    
    'CurrentDb.OpenRecordset(Name:="All_Z_Summary", Type:=dbOpenDynaset, Options:=dbInconsistent, LockEdit:=dbOptimistic)
    Set rsWriteData = CurrentDb.OpenRecordset("All_Combined_Screenings_SparseMatrix", dbOpenDynaset, dbInconsistent, dbOptimistic)
    'Move cursor to first row, this will be used to itterate through all the rows in order
    rsWriteData.MoveFirst
    'myTC_Number = rsWriteData.Fields(Trial_Card)
    'Debug.Print ("myTC_Number: " & myTC_Number)
    
        While (Not rsWriteData.EOF)
            myColumnCounter = 0 'This collects the field count of the selected row
            myTC_Number = rsWriteData.Fields("Trial_Card")
            'Debug.Print ("myTC_Number: " & myTC_Number)
            
            For i = aCL_EventIndexList(0) To aCL_OWLD_index_pos
                DateColumnVal = allColumnsList(i)
                
                'Build SQL string
                'myReadSQLstr = "SELECT * FROM All_Combined_Screenings_SparseMatrix WHERE Final_Sts_A_T <> 'X/X' AND [2017/06/30] = '1' AND (" & myWhereClause & ");"
                myReadSQLstr = "SELECT " & DateColumnVal & " FROM " & myTableName & " WHERE ([" & DateColumnVal & "] = '1') AND ([Trial_Card] = '" & myTC_Number & "');"
                'Debug.Print "myReadSQLstr: " & myReadSQLstr
                
                'Run SQL
                'CurrentDb.OpenRecordset(Name:=myReadSQLstr, Type:=dbOpenSnapshot, Options:=dbReadOnly, LockEdit:=)
                Set rsReadData = CurrentDb.OpenRecordset(myReadSQLstr, dbOpenSnapshot, dbReadOnly)
                
                If rsReadData.EOF Then
                    'date column doesn't equal 1, no record returned
                    'Do nothing, pass
                Else
                    rsReadData.MoveLast
                    myColumnCounter = myColumnCounter + 1
                End If
                
                'Debug.Print "rsReadData.RecordCount: " & rsReadData.RecordCount
                'Debug.Print "myColumnCounter: " & myColumnCounter
                
                myReadSQLstr = Empty
                rsReadData.Close 'The column has been counted, close the record set
                
            Next i ' next date column
                
            Debug.Print "myColumnCounter: " & myColumnCounter
            myWriteColumn = "Rescreen_counts_BT_to_OWLD"
            
            'Write the row sum
            rsWriteData.Edit
            rsWriteData.Fields(myWriteColumn) = myColumnCounter
            rsWriteData.Update
            
            myWriteColumn = Empty
            myColumnCounter = Empty
            
            i = Empty
            rsWriteData.MoveNext
                
        Wend 'While end loop
        
    rsWriteData.Close
    Set rsWriteData = Nothing
    Set rsReadData = Nothing
    Debug.Print ("Finished " & myTableVarList & "!")
    
Next myTableVarList 'Cycle through the Sparse Matrix tables

myTableVarList = Empty

Debug.Print ("Finished!")

End Sub

Sub SetRecordReScreenCounts()
'This counts the rescreens across the record row _
by event window
Dim mySQLstr As String 'This Gets the Table RecordSet
Dim rs As DAO.Recordset 'This is the table recordset
Dim i As Long ' Used as the allColumnsList(i) itterable
Dim myTCcounter As Long
Dim myTableVarList As Variant

'Rescreen_counts_BT_to_OWLD
For Each myTableVarList In All_SparseMatrixList
    mySQLstr = "SELECT * FROM " & myTableVarList & ""
    
    'CurrentDb.OpenRecordset(Name:=mySQLstr, Type:=dbOpenDynaset, Options:=dbInconsistent, LockEdit:=dbOptimistic)
    Set rs = CurrentDb.OpenRecordset(mySQLstr, dbOpenDynaset, dbInconsistent, dbOptimistic)
    
    'move rs cursor to first record
    rs.MoveFirst
    
    'Rescreen_counts_BT_to_OWLD, ALL Events
    While (Not rs.EOF)
        myTCcounter = 0
        'Step across the date columns
        For i = aCL_BT_index_pos To aCL_OWLD_index_pos
            'Debug.Print ("i = " & i)
            'Debug.Print ("rs.field = " & rs.Fields("Trial_Card").Value)
            
            If rs.Fields(allColumnsList(i)) = "0" Then
                'zero's are found in "All_Combined_Screenings_SparseMatrix" and "All_Screenings_Only_SparseMatrix"
                'Do Nothing, Pass
            ElseIf rs.Fields(allColumnsList(i)) = "~~~" Then
                'tripple tilda's are found in "All_XX_Screen_Only_SparseMatrix"
                'Do Nothing, Pass
            ElseIf rs.Fields(allColumnsList(i)) = "1" Then
                myTCcounter = myTCcounter + 1
            ElseIf rs.Fields(allColumnsList(i)) <> "~~~" Then
                myTCcounter = myTCcounter + 1
                'This needs to be last or it will count everything
            Else
                'Invalid data
                'Do Nothing, Pass
            End If
            
            Debug.Print ("rs.Fields(allColumnsList(i)): " & allColumnsList(i).Value)
            
        Next i
        'write the row total
        rs.Edit
        rs.Fields("Rescreen_counts_BT_to_OWLD") = myTCcounter
        rs.Update
                
        Debug.Print ("rs.Fields(Trial_Card): " & rs.Fields("Trial_Card").Value)
        
        rs.MoveNext
        
    Wend 'While end loop
    
    'move rs cursor back to first record
    rs.MoveFirst
    
    'Rescreen_counts_AT_to_OWLD, INSURV events
    While (Not rs.EOF)
        myTCcounter = 0
        'Step across the date columns
        For i = aCL_AT_index_pos To aCL_OWLD_index_pos
            'Debug.Print ("i = " & i)
            'Debug.Print ("rs.field = " & rs.Fields("Trial_Card").Value)
            
            If rs.Fields(allColumnsList(i)) = "0" Then
                'zero's are found in "All_Combined_Screenings_SparseMatrix" and "All_Screenings_Only_SparseMatrix"
                'Do Nothing, Pass
            ElseIf rs.Fields(allColumnsList(i)) = "~~~" Then
                'tripple tilda's are found in "All_XX_Screen_Only_SparseMatrix"
                'Do Nothing, Pass
            ElseIf rs.Fields(allColumnsList(i)) = "1" Then
                myTCcounter = myTCcounter + 1
            ElseIf rs.Fields(allColumnsList(i)) <> "~~~" Then
                myTCcounter = myTCcounter + 1
                'This needs to be last or it will count everything
            Else
                'Invalid data
                'Do Nothing, Pass
            End If
            
            Debug.Print ("rs.Fields(allColumnsList(i)): " & rs.Fields(allColumnsList(i).Value))
        
        Next i
        'write the row total
        rs.Edit
        rs.Fields("Rescreen_counts_BT_to_OWLD") = myTCcounter
        rs.Update
                
        Debug.Print ("rs.Fields(Trial_Card): " & rs.Fields("Trial_Card").Value)
        
        rs.MoveNext
        
    Wend 'While end loop
    
    'move rs cursor back to first record
    rs.MoveFirst
    
    'Rescreen_counts_FCT_to_OWLD
    While (Not rs.EOF)
        myTCcounter = 0
        'Step across the date columns
        For i = aCL_FCT_index_pos To aCL_OWLD_index_pos
            'Debug.Print ("i = " & i)
            'Debug.Print ("rs.field = " & rs.Fields("Trial_Card").Value)
            
            If rs.Fields(allColumnsList(i)) = "0" Then
                'zero's are found in "All_Combined_Screenings_SparseMatrix" and "All_Screenings_Only_SparseMatrix"
                'Do Nothing, Pass
            ElseIf rs.Fields(allColumnsList(i)) = "~~~" Then
                'tripple tilda's are found in "All_XX_Screen_Only_SparseMatrix"
                'Do Nothing, Pass
            ElseIf rs.Fields(allColumnsList(i)) = "1" Then
                myTCcounter = myTCcounter + 1
            ElseIf rs.Fields(allColumnsList(i)) <> "~~~" Then
                myTCcounter = myTCcounter + 1
                'This needs to be last or it will count everything
            Else
                'Invalid data
                'Do Nothing, Pass
            End If
            
             Debug.Print ("rs.Fields(allColumnsList(i)): " & rs.Fields(allColumnsList(i).Value))
             
        Next i
        'write the row total
        rs.Edit
        rs.Fields("Rescreen_counts_BT_to_OWLD") = myTCcounter
        rs.Update
                
        Debug.Print ("rs.Fields(Trial_Card): " & rs.Fields("Trial_Card").Value)
        
        rs.MoveNext
        
    Wend 'While end loop
    
    rs.Close
    mySQLstr = Empty

Next myTableVarList

Set rs = Nothing

End Sub
