Option Compare Database
Option Explicit


Sub SetAllSummaryCounts()
'This function loops through the SparseMatrix Tables and _
Sums the Rescreen Changes by Date Column
Dim aCL_EventIndexList As Object
Dim EventFilterList As Object
Dim myReadSQLstr As String 'This queries the DB to get the Table RecordSet
Dim myWriteSQLstr As String 'This queries the DB to get the Table RecordSet
Dim rsReadData As DAO.Recordset 'This is the table recordset
Dim rsWriteData As DAO.Recordset 'This is the table recordset
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
Dim j As Long ' Used as the aCL_EventIndexList(j) itterable
Dim myColumnCounter As Long
Dim myTableVarList As Variant
Dim myEventFilterVar As Variant
Dim DateColumnVal As String
Dim DeltaRescreen As String
Dim myWhereClause As String
Dim myTableName As String
Dim myWriteRow As String

'Load the application lists and vars
AddingToMyDateLists

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

'List of allColumnsList Trial dates index
'Dim aCL_EventIndexList As Object ' This Declares the object handle, without any properties or methods
'Dim aCL_EventIndexList As New ArrayList ' This Declares the object handle with early binding
Set aCL_EventIndexList = CreateObject("System.Collections.ArrayList") ' This is late binding
'aCL_EventIndexList
' Add items
aCL_EventIndexList.Add aCL_BT_index_pos ' BT aCL_EventIndex(0)
aCL_EventIndexList.Add aCL_AT_index_pos ' AT aCL_EventIndex(1)
aCL_EventIndexList.Add aCL_FCT_index_pos ' FCT aCL_EventIndex(2)

'List of Event query paramiters to filter the returned recordset
'Dim EventFilterList As Object ' This Declares the object handle, without any properties or methods
'Dim EventFilterList As New ArrayList ' This Declares the object handle with early binding
Set EventFilterList = CreateObject("System.Collections.ArrayList") ' This is late binding
'EventFilterList
' Add items
EventFilterList.Add "((Event = 'BT') AND (Trial_ID LIKE 'B*'))" 'This is BT New Cards written at the BT Event. Trial_ID will have a leading "B" and may or may not have trailing Trial_ID's.
EventFilterList.Add "((Event <> 'BT') AND (Trial_ID LIKE '*B*'))" 'This is BT Roll and Splits
EventFilterList.Add "((Event = 'AT') AND (Trial_ID LIKE 'C*'))" 'This is AT New Cards written at the AT Event. Trial_ID will have a leading "C" and may or may not have trailing Trial_ID's.
EventFilterList.Add "((Event <> 'AT') AND (Trial_ID LIKE '*C*'))" 'This is AT Roll and Splits
EventFilterList.Add "((Event = 'FCT') AND (Trial_ID LIKE 'F*'))" ''This is FCT New Cards written at the FCT Event. Trial_ID will have a leading "F" and may or may not have trailing Trial_ID's.
EventFilterList.Add "((Event <> 'FCT') AND (Trial_ID LIKE '*F*'))" 'This is FCT Roll and Splits
EventFilterList.Add "((Event = 'BT') OR (Event = 'AT') OR (Event = 'FCT'))" 'This is All BT/AT/FCT New and Splits
EventFilterList.Add "(((Event = 'AT') OR (Event = 'FCT')) AND ((Trial_ID LIKE '*C*') OR (Trial_ID LIKE '*F*')))" 'This is INSURV New, Roll and Splits

'Rescreen_counts_Event_to_OWLD

Set rsWriteData = CurrentDb.OpenRecordset("All_Z_Summary", dbOpenDynaset)

'Cycle through the Sparse Matrix tables
For Each myTableVarList In All_SparseMatrixList
    '"All_Combined_Screenings_SparseMatrix" ' All_SparseMatrixList(0)
    '"All_Screenings_Only_SparseMatrix" ' All_SparseMatrixList(1)
    '"All_XX_Screen_Only_SparseMatrix" ' All_SparseMatrixList(2)
    myTableName = myTableVarList
    
    'Filter the table
    For Each myEventFilterVar In EventFilterList
        'EventFilterList.Add "(Final_Sts_A_T <> 'X/X') AND ((Events = 'BT') OR (Trial_ID LIKE 'B'))" 'This is BT New and Splits
        myWhereClause = myEventFilterVar
        j = 0 'This is the aCL_EventIndexList list index
        
        'Step accross the date columns
''' This could stop at "OWLD" or at "Final"
        'For i = aCL_EventIndexList(j) To aCL_OWLD_index_pos
        For i = aCL_EventIndexList(j) To aCL_Final_index_pos
            'allColumnsList.Add "2017/06/30" ' BT
            DateColumnVal = allColumnsList(i)
            
            myColumnCounter = 0
            
            'Build SQL string
            'myReadSQLstr = "SELECT * FROM All_Combined_Screenings_SparseMatrix WHERE Final_Sts_A_T <> 'X/X' AND [2017/06/30] = '1' AND (" & myWhereClause & ");"
            myReadSQLstr = "SELECT * FROM " & myTableVarList & " WHERE Final_Sts_A_T <> 'X/X' AND (([" & DateColumnVal & "] = '1') OR([" & DateColumnVal & "] <> '~~~')) AND (" & myWhereClause & ");"
            Debug.Print "myReadSQLstr: " & myReadSQLstr
            
            'Run SQL
            Set rsReadData = CurrentDb.OpenRecordset(myReadSQLstr, dbOpenDynaset)
            
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
            
            myWriteRow = "Date_Diff_from_Prior"
            
            'Write the column sum
            rsWriteData.FindFirst "[Table_Source] = '" & myWriteRow & "'"
            rsWriteData.Edit
            rsWriteData.Fields(DateColumnVal) = myColumnCounter
            rsWriteData.Update
            
            myWriteRow = Empty
            DateColumnVal = Empty
            'rsReadData = Empty
            
            j = j + 1
            
        Next i ' myDateVar
        
        i = Empty
    
    Next myEventFilterVar 'Filter the table
    
    myEventFilterVar = Empty
    
Next myTableVarList 'Cycle through the Sparse Matrix tables

myTableVarList = Empty
rsWriteData.Close
rsReadData.Close
myReadSQLstr = Empty

Set rsWriteData = Nothing
Set rsReadData = Nothing

End Sub


Sub GetDateDif()
'This function loops through the Dates and _
computes running days, and difference in date by Date Column
Dim aCL_EventIndexList As Object
Dim rsWriteData As DAO.Recordset 'This is the table recordset

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
Dim myWriteRow As String

'Load the application lists and vars
AddingToMyDateLists

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

'List of allColumnsList Trial dates index
'Dim aCL_EventIndexList As Object ' This Declares the object handle, without any properties or methods
'Dim aCL_EventIndexList As New ArrayList ' This Declares the object handle with early binding
Set aCL_EventIndexList = CreateObject("System.Collections.ArrayList") ' This is late binding
'aCL_EventIndexList
' Add items
aCL_EventIndexList.Add aCL_BT_index_pos ' BT aCL_EventIndex(0)
aCL_EventIndexList.Add aCL_AT_index_pos ' AT aCL_EventIndex(1)
aCL_EventIndexList.Add aCL_FCT_index_pos ' FCT aCL_EventIndex(2)

'Rescreen_counts_Event_to_OWLD

Set rsWriteData = CurrentDb.OpenRecordset("All_Z_Summary", dbOpenDynaset)

'Step accross all of the date columns to _
get the difference in days between last _
report and current report
dateIndex = aCL_BT_index_pos 'This is the first column
For Each d In allColumnsList
    DateColumnVal = d
    
    If d = myBTvar Then
        'd = 2017/06/30 or allColumnsList(0)
        DateColumnDiff = 0
        Debug.Print ("DateColumnDiff: " & DateColumnDiff & " CurCol:" & DateColumnVal)
    Else
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
        Debug.Print ("DateColumnDiff: " & DateColumnDiff & " CurCol:" & myCurDateCol & " PriorCol:" & myPriorDateCol)
    End If
    
    'Write date to row
    myWriteRow = "Date_Diff_from_Prior"
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
    
Next d

'Step accross all of the date columns to _
get the cumulative days from Event BT date _
report and current report date
dateIndex = aCL_BT_index_pos 'This is the first column
For Each d In allColumnsList
    DateColumnVal = d
    
    If d = myBTvar Then
        'd = 2017/06/30 or allColumnsList(0)
        DateColumnDiff = 0
        Debug.Print ("DateColumnDiff: " & DateColumnDiff & " CurCol:" & DateColumnVal)
    Else
        'get column string date values
        varCurDateCol = allColumnsList(dateIndex) 'YYYY/MM/DD 2017/07/04
        varPriorDateCol = allColumnsList(aCL_BT_index_pos) 'YYYY/MM/DD 2017/06/30
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
        Debug.Print ("DateColumnDiff: " & DateColumnDiff & " CurCol:" & myCurDateCol & " PriorCol:" & myPriorDateCol)
    End If
    
    'Write cumulative days to row
    myWriteRow = "DaysPost_BT"
    rsWriteData.FindFirst "[Table_Source] = '" & myWriteRow & "'"
    'Write the column sum
    rsWriteData.Edit
    rsWriteData.Fields(DateColumnVal) = DateColumnDiff
    rsWriteData.Update
    
    'Write cumulative days to row
    myWriteRow = "DaysPost_ALL"
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
    
Next d

'Step accross all of the date columns to _
get the cumulative days from Event AT date _
report and current report date
dateIndex = aCL_BT_index_pos 'This is the first column
For Each d In allColumnsList
    DateColumnVal = d
    
    If d = myATvar Then
        'd = 2017/08/18 or allColumnsList(7)
        DateColumnDiff = 0
        Debug.Print ("DateColumnDiff: " & DateColumnDiff & " CurCol:" & DateColumnVal)
    ElseIf d < myATvar Then
        'd < 2017/08/18 or allColumnsList(7)
        'Could make this a negitive number, but that has never been used
        DateColumnDiff = 0
        Debug.Print ("DateColumnDiff: " & DateColumnDiff & " CurCol:" & DateColumnVal & " < " & myATvar)
    Else
        'get column string date values
        varCurDateCol = allColumnsList(dateIndex) 'YYYY/MM/DD 2017/07/04
        varPriorDateCol = allColumnsList(aCL_AT_index_pos) 'YYYY/MM/DD 2017/06/30
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
        Debug.Print ("DateColumnDiff: " & DateColumnDiff & " CurCol:" & myCurDateCol & " PriorCol:" & myPriorDateCol)
    End If
    
    'Write cumulative days to row
    myWriteRow = "DaysPost_AT"
    rsWriteData.FindFirst "[Table_Source] = '" & myWriteRow & "'"
    'Write the column sum
    rsWriteData.Edit
    rsWriteData.Fields(DateColumnVal) = DateColumnDiff
    rsWriteData.Update
    
    'Write cumulative days to row
    myWriteRow = "DaysPost_INSURV"
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
    
Next d

'Step accross all of the date columns to _
get the cumulative days from Event AT date _
report and current report date
dateIndex = aCL_BT_index_pos 'This is the first column
For Each d In allColumnsList
    DateColumnVal = d
    
    If d = myFCTvar Then
        'd = 2018/10/26 or allColumnsList(30)
        DateColumnDiff = 0
        Debug.Print ("DateColumnDiff: " & DateColumnDiff & " CurCol:" & DateColumnVal)
    ElseIf d < myFCTvar Then
        'd < 2018/10/26 or allColumnsList(30)
        'Could make this a negitive number, but that has never been used
        DateColumnDiff = 0
        Debug.Print ("DateColumnDiff: " & DateColumnDiff & " CurCol:" & DateColumnVal & " < " & myFCTvar)
    Else
        'get column string date values
        varCurDateCol = allColumnsList(dateIndex) 'YYYY/MM/DD 2017/07/04
        varPriorDateCol = allColumnsList(aCL_FCT_index_pos) 'YYYY/MM/DD 2017/06/30
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
        Debug.Print ("DateColumnDiff: " & DateColumnDiff & " CurCol:" & myCurDateCol & " PriorCol:" & myPriorDateCol)
    End If
    
    'Write cumulative days to row
    myWriteRow = "DaysPost_FCT"
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
    
Next d

'Close the connection
rsWriteData.Close
'Empty the connection Object
Set rsWriteData = Nothing

End Sub
