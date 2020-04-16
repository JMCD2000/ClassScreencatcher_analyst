Option Compare Database
Option Explicit

Public Sub AppendDateFieldsToDataTables()
'This adds, appends to the end the date columns to the data tables.
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
        'Debug.Print "myInsertTableSQL Statement: " & myInsertTableSQL & "."
        
        CurrentDb.Execute myInsertTableSQL
        'Debug.Print "Completed adding column Date: " & myDateVarList & "."
        
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
'This removes, drops the date columns in the data tables.
'allColumnsList
'dataTablesList
Dim myInsertTableSQL As String
Dim myDateVarList As Variant  ' Used to cycle through each date column as an iterable allColumnsList()
Dim myTableVarList As Variant  ' Used to cycle through each date column as an iterable dataTablesList()

'Run the list builder
'AddingToMyDateLists

For Each myTableVarList In dataTablesList
    For Each myDateVarList In allColumnsList
        
        myInsertTableSQL = "ALTER TABLE " & myTableVarList & " DROP Column [" & myDateVarList & "]; "
        'Debug.Print "myInsertTableSQL Statement: " & myInsertTableSQL & "."
        
        CurrentDb.Execute myInsertTableSQL
        'Debug.Print "Completed removing column Date: " & myDateVarList & "."
        
    Next myDateVarList
    
    myDateVarList = Empty
    Debug.Print ("Finished Table: " & myTableVarList)
    
Next myTableVarList

myDateVarList = Empty
myTableVarList = Empty

'ClearMyDateLists 'Empty created list objects

' Debug.Print vbCrLf & "Completed the Non Trial Screenings Update Query." & vbCrLf
End Sub

Public Sub InsertTrialCardNumbers()
'This adds, inserts the trial card values into the data tables _
Primary Key field.
'beanFinal = "2020/04/03_LPD26_Final" ' name reference to the _Final table used in the SQL statments
'dataTablesList
Dim myInsertTableSQL As String
Dim myTableVarList As Variant  ' Used to cycle through each date column as an iterable dataTablesList()

'Run the list builder _
Call Module SetListAndVars.AddingToMyDateLists()
AddingToMyDateLists

For Each myTableVarList In dataTablesList
    
    myInsertTableSQL = "INSERT INTO " & myTableVarList & " ([Trial_Card]) " _
    & "SELECT [Trial_Card] " _
    & "FROM [" & beanFinal & "]; "
    Debug.Print "myInsertTableSQL Statement: " & myInsertTableSQL & "."
        
    CurrentDb.Execute myInsertTableSQL
    'Debug.Print "Completed adding column Date: " & myDateVarList & "."
           
    Debug.Print ("Finished Table: " & myTableVarList)
    
Next myTableVarList

myTableVarList = Empty

'ClearMyDateLists 'Empty created list objects

' Debug.Print vbCrLf & "Completed the Non Trial Screenings Update Query." & vbCrLf

End Sub
