Option Compare Database
Option Explicit


Public Function IndexInArray(myArray As Variant, mySearch As Variant) As Variant
' This Function loops an passed in array to find and _
return the index of the item in the array
Dim i As Long
Dim upper As Long
upper = myArray.Count

For i = 0 To myArray.Count
    If myArray(i) = mySearch Then
        IndexInArray = i
        Exit Function
    End If
Next i
'mySearch var was not found in the array, return null
IndexInArray = Null
End Function


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
            
        Next i
        'write the row total
        rs.Edit
        rs.Fields("Rescreen_counts_BT_to_OWLD") = myTCcounter
        rs.Update
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
        
        Next i
        'write the row total
        rs.Edit
        rs.Fields("Rescreen_counts_AT_to_OWLD") = myTCcounter
        rs.Update
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
            
        Next i
        'write the row total
        rs.Edit
        rs.Fields("Rescreen_counts_FCT_to_OWLD") = myTCcounter
        rs.Update
        rs.MoveNext
    Wend 'While end loop
    
    rs.Close
    mySQLstr = Empty

Next myTableVarList

Set rs = Nothing

End Sub
