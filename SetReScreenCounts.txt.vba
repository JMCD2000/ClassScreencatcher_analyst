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


Sub SetAllReScreenCounts()

Dim mySQLstr As String 'This Gets the Table RecordSet
Dim rs As DAO.Recordset 'This is the table recordset
Dim myBTvar As String
Dim aCL_BT_index_pos As Long ' Index of the BT Event in the allColumnsList
Dim myATvar As String
Dim aCL_AT_index_pos As Long ' Index of the AT Event in the allColumnsList
Dim myFCTvar As String
Dim aCL_FCT_index_pos As Long ' Index of the FCT Event in the allColumnsList
Dim myOWLDvar As String
Dim aCL_OWLD_index_pos As Long ' Index of the OWLD Event in the allColumnsList
Dim myFinalvar As String
Dim aCL_Final_index_pos As Long ' Index of the Final Event in the allColumnsList
Dim i As Long ' Used as the allColumnsList(i) itterable
Dim myTCcounter As Long
Dim myTableVarList As Variant

AddingToMyDateLists

'DAO.DBEngine.SetOption dbMaxLocksPerFile, 15000

' trialsOnlyList(0).Add "2017/06/30" ' BT
myBTvar = trialsOnlyList(0)
aCL_BT_index_pos = IndexInArray(allColumnsList, myBTvar)
' trialsOnlyList(1).Add "2017/08/18" ' AT
myATvar = trialsOnlyList(1)
aCL_AT_index_pos = IndexInArray(allColumnsList, myATvar)
' trialsOnlyList(2).Add "2018/10/26" ' FCT
myFCTvar = trialsOnlyList(2)
aCL_FCT_index_pos = IndexInArray(allColumnsList, myFCTvar)
' trialsOnlyList(3).Add "2019/09/19" ' OWLD
myOWLDvar = trialsOnlyList(3)
aCL_OWLD_index_pos = IndexInArray(allColumnsList, myOWLDvar)
' trialsOnlyList(4).Add "2020/04/03" ' Final
myFinalvar = trialsOnlyList(4)
aCL_Final_index_pos = IndexInArray(allColumnsList, myFinalvar)

'Rescreen_counts_BT_to_OWLD
For Each myTableVarList In All_SparseMatrixList
    mySQLstr = "SELECT * FROM " & myTableVarList & ""
    
    Set rs = CurrentDb.OpenRecordset(mySQLstr, dbOpenDynaset)
    rs.MoveFirst
    
    'Rescreen_counts_BT_to_OWLD
    While (Not rs.EOF)
        myTCcounter = 0
        For i = aCL_BT_index_pos To aCL_OWLD_index_pos
            Debug.Print ("i = " & i)
            Debug.Print ("rs.field = " & rs.Fields("Trial_Card").Value)
            
            If rs.Fields(allColumnsList(i)) = "0" Then
                'Do Nothing, Pass
            ElseIf rs.Fields(allColumnsList(i)) = "~~~" Then
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
        rs.Edit
        rs.Fields("Rescreen_counts_BT_to_OWLD") = myTCcounter
        rs.Update
        rs.MoveNext
    Wend 'While end loop
    
    rs.MoveFirst
    
    'Rescreen_counts_AT_to_OWLD
    While (Not rs.EOF)
        myTCcounter = 0
        For i = aCL_AT_index_pos To aCL_OWLD_index_pos
            Debug.Print ("i = " & i)
            Debug.Print ("rs.field = " & rs.Fields("Trial_Card").Value)
            
            If rs.Fields(allColumnsList(i)) = "0" Then
                'Do Nothing, Pass
            ElseIf rs.Fields(allColumnsList(i)) = "~~~" Then
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
        rs.Edit
        rs.Fields("Rescreen_counts_AT_to_OWLD") = myTCcounter
        rs.Update
        rs.MoveNext
    Wend 'While end loop
    
    rs.MoveFirst
    
    'Rescreen_counts_FCT_to_OWLD
    While (Not rs.EOF)
        myTCcounter = 0
        For i = aCL_FCT_index_pos To aCL_OWLD_index_pos
            Debug.Print ("i = " & i)
            Debug.Print ("rs.field = " & rs.Fields("Trial_Card").Value)
            
            If rs.Fields(allColumnsList(i)) = "0" Then
                'Do Nothing, Pass
            ElseIf rs.Fields(allColumnsList(i)) = "~~~" Then
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

Sub SetEventReScreenCounts()

Dim mySQLstr As String 'This Gets the Table RecordSet
Dim rs As DAO.Recordset 'This is the table recordset
Dim myBTvar As String
Dim aCL_BT_index_pos As Long ' Index of the BT Event in the allColumnsList
Dim myATvar As String
Dim aCL_AT_index_pos As Long ' Index of the AT Event in the allColumnsList
Dim myFCTvar As String
Dim aCL_FCT_index_pos As Long ' Index of the FCT Event in the allColumnsList
Dim myOWLDvar As String
Dim aCL_OWLD_index_pos As Long ' Index of the OWLD Event in the allColumnsList
Dim myFinalvar As String
Dim aCL_Final_index_pos As Long ' Index of the Final Event in the allColumnsList
Dim i As Long ' Used as the allColumnsList(i) itterable
Dim myTCcounter As Long
Dim myTableVarList As Variant

AddingToMyDateLists

' trialsOnlyList(0).Add "2017/06/30" ' BT
myBTvar = trialsOnlyList(0)
aCL_BT_index_pos = IndexInArray(trialsOnlyList, myBTvar)
' trialsOnlyList(1).Add "2017/08/18" ' AT
myATvar = trialsOnlyList(1)
aCL_AT_index_pos = IndexInArray(trialsOnlyList, myATvar)
' trialsOnlyList(2).Add "2018/10/26" ' FCT
myFCTvar = trialsOnlyList(2)
aCL_FCT_index_pos = IndexInArray(trialsOnlyList, myFCTvar)
' trialsOnlyList(3).Add "2019/09/19" ' OWLD
myOWLDvar = trialsOnlyList(3)
aCL_OWLD_index_pos = IndexInArray(trialsOnlyList, myOWLDvar)
' trialsOnlyList(4).Add "2020/04/03" ' Final
myFinalvar = trialsOnlyList(4)
aCL_Final_index_pos = IndexInArray(trialsOnlyList, myFinalvar)

'Rescreen_counts_BT_to_OWLD
For Each myTableVarList In Events_SparseMatrixList
    mySQLstr = "SELECT * FROM " & myTableVarList & ""
    Set rs = CurrentDb.OpenRecordset(mySQLstr, dbOpenDynaset)
    rs.MoveFirst
    
    'Rescreen_counts_BT_to_OWLD
    While (Not rs.EOF)
        myTCcounter = 0
        For i = aCL_BT_index_pos To aCL_OWLD_index_pos
            If rs.Fields(allColumnsList(i)) = "0" Then
                'Do Nothing, Pass
            ElseIf rs.Fields(allColumnsList(i)) = "~~~" Then
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
        rs.Edit
        rs.Fields("Rescreen_counts_BT_to_OWLD") = myTCcounter
        rs.Update
        rs.MoveNext
    Wend 'While end loop
    
    rs.MoveFirst
    
    'Rescreen_counts_AT_to_OWLD
    While (Not rs.EOF)
        myTCcounter = 0
        For i = aCL_AT_index_pos To aCL_OWLD_index_pos
            If rs.Fields(allColumnsList(i)) = "0" Then
                'Do Nothing, Pass
            ElseIf rs.Fields(allColumnsList(i)) = "~~~" Then
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
        rs.Edit
        rs.Fields("Rescreen_counts_AT_to_OWLD") = myTCcounter
        rs.Update
        rs.MoveNext
    Wend 'While end loop
    
        rs.MoveFirst
    
    'Rescreen_counts_FCT_to_OWLD
    While (Not rs.EOF)
        myTCcounter = 0
        For i = aCL_FCT_index_pos To aCL_OWLD_index_pos
            If rs.Fields(allColumnsList(i)) = "0" Then
                'Do Nothing, Pass
            ElseIf rs.Fields(allColumnsList(i)) = "~~~" Then
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
