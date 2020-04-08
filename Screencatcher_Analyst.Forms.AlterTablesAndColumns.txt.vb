Option Compare Database
Option Explicit

Public Sub AppendDateFieldsToDataTables()
'This adds, appends to the end of the table the date columns to the data tables.
'allColumnsList
'dataTablesList
Dim myInsertTableSQL As String
Dim myDateVarList As Variant  ' Used to cycle through each date column as an iterable allColumnsList()
Dim myTableVarList As Variant  ' Used to cycle through each date column as an iterable dataTablesList()

'Run the list builder
'AddingToMyDateLists

For Each myTableVarList In dataTablesList
    For Each myDateVarList In allColumnsList
        
        myInsertTableSQL = "ALTER TABLE " & myTableVarList & " ADD COLUMN [" & myDateVarList & "] TEXT(255); "
        Debug.Print "myInsertTableSQL Statement: " & myInsertTableSQL & "."
        
        CurrentDb.Execute myInsertTableSQL
        'Debug.Print "Completed column Date: " & myDateVarList & "."
        
    Next myDateVarList
    
    myDateVarList = Empty
    Debug.Print ("Finished Table: " & myTableVarList)
    
Next myTableVarList

myDateVarList = Empty
myTableVarList = Empty

'ClearMyDateLists 'Empty created list objects

' Debug.Print vbCrLf & "Completed the Non Trial Screenings Update Query." & vbCrLf
End Sub

Public Sub DropDateFieldsFromDataTables()
' This would be to remove the date columns _
right now it is a manual drill.

End Sub
