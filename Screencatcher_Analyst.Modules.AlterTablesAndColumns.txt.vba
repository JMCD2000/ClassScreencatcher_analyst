Option Compare Database
Option Explicit


Public Sub AppendDateFieldsToDataTables()
'This adds, appends to the end the date columns to the data tables.
'allColumnsList
'All_dataTablesList
'Events_dataTablesList
Dim myInsertTableSQL As String
Dim myDateVarList As Variant  ' Used to cycle through each date column as an iterable allColumnsList()
Dim myTableVarList As Variant  ' Used to cycle through each date column as an iterable dataTablesList()

    'Run the Add date columns to the tables beginging with ALL
    For Each myTableVarList In All_dataTablesList
        If myTableVarList = "All_TC_Screen_Agg" Then
            'This is commented out because it doesn't contain date columns
            'Pass 'Next myTableVarList
        Else
            For Each myDateVarList In allColumnsList
                myInsertTableSQL = "ALTER TABLE " & myTableVarList & " ADD COLUMN [" & myDateVarList & "] TEXT(255); "
                'Debug.Print "myInsertTableSQL Statement: " & myInsertTableSQL & "."
                CurrentDb.Execute myInsertTableSQL
                'Debug.Print "Completed adding column Date: " & myDateVarList & "."
            Next myDateVarList
            myDateVarList = Empty
            'Debug.Print ("Finished Table: " & myTableVarList)
        End If
    Next myTableVarList
    
    myDateVarList = Empty
    myTableVarList = Empty
    
    For Each myTableVarList In Events_dataTablesList
        If myTableVarList = "Events_TC_Screen_Agg" Then
            'This is commented out because it doesn't contain date columns
            'Pass 'Next myTableVarList
        Else
            For Each myDateVarList In trialsOnlyList
                myInsertTableSQL = "ALTER TABLE " & myTableVarList & " ADD COLUMN [" & myDateVarList & "] TEXT(255); "
                'Debug.Print "myInsertTableSQL Statement: " & myInsertTableSQL & "."
                CurrentDb.Execute myInsertTableSQL
                'Debug.Print "Completed adding column Date: " & myDateVarList & "."
            Next myDateVarList
            myDateVarList = Empty
            'Debug.Print ("Finished Table: " & myTableVarList)
        End If
    Next myTableVarList
    
    myDateVarList = Empty
    myTableVarList = Empty
    
    ' Debug.Print vbCrLf & "Completed the Non Trial Screenings Update Query." & vbCrLf
End Sub


Public Sub DropDateFieldsFromDataTables()
'This removes, drops the date columns in the data tables.
'allColumnsList
'All_dataTablesList
'Events_dataTablesList
Dim myInsertTableSQL As String
Dim myDateVarList As Variant  ' Used to cycle through each date column as an iterable allColumnsList()
Dim myTableVarList As Variant  ' Used to cycle through each date column as an iterable dataTablesList()

    'Run the Drop date columns to the tables beginging with ALL
    For Each myTableVarList In All_dataTablesList
        If myTableVarList = "All_TC_Screen_Agg" Then
            'This is commented out because it doesn't contain date columns
            'Pass 'Next myTableVarList
        Else
            For Each myDateVarList In allColumnsList
                myInsertTableSQL = "ALTER TABLE " & myTableVarList & " DROP Column [" & myDateVarList & "]; "
                'Debug.Print "myInsertTableSQL Statement: " & myInsertTableSQL & "."
                CurrentDb.Execute myInsertTableSQL
                'Debug.Print "Completed removing column Date: " & myDateVarList & "."
            Next myDateVarList
            myDateVarList = Empty
            'Debug.Print ("Finished Table: " & myTableVarList)
        End If
    Next myTableVarList
    
    myDateVarList = Empty
    myTableVarList = Empty
    
    'Run the Drop date columns to the tables beginging with Events
    For Each myTableVarList In Events_dataTablesList
        If myTableVarList = "Events_TC_Screen_Agg" Then
            'This is commented out because it doesn't contain date columns
            'Pass 'Next myTableVarList
        Else
            For Each myDateVarList In trialsOnlyList
                myInsertTableSQL = "ALTER TABLE " & myTableVarList & " DROP Column [" & myDateVarList & "]; "
                'Debug.Print "myInsertTableSQL Statement: " & myInsertTableSQL & "."
                CurrentDb.Execute myInsertTableSQL
                'Debug.Print "Completed removing column Date: " & myDateVarList & "."
            Next myDateVarList
            myDateVarList = Empty
            'Debug.Print ("Finished Table: " & myTableVarList)
        End If
    Next myTableVarList
    
    myDateVarList = Empty
    myTableVarList = Empty
    
' Debug.Print vbCrLf & "Completed the Non Trial Screenings Update Query." & vbCrLf
End Sub


Public Sub InsertTrialCardNumbers()
'This adds, inserts the trial card values into the data tables Primary Key field.
'beanFinal = "2020/04/03_LPD26_Final" ' name reference to the _Final table used in the SQL statments
'All_dataTablesList
'Events_dataTablesList
Dim myInsertTableSQL As String
Dim myTableVarList As Variant  ' Used to cycle through each date column as an iterable dataTablesList()

    'Run the Add date columns to the tables beginging with ALL
    For Each myTableVarList In All_dataTablesList
        myInsertTableSQL = "INSERT INTO " & myTableVarList & " ([Trial_Card]) " _
        & "SELECT [Trial_Card] " _
        & "FROM [" & beanFinal & "]; "
        Debug.Print "myInsertTableSQL Statement: " & myInsertTableSQL & "."
        CurrentDb.Execute myInsertTableSQL
        'Debug.Print "Completed adding column Date: " & myDateVarList & "."
        'Debug.Print ("Finished Table: " & myTableVarList)
    Next myTableVarList
    
    myTableVarList = Empty
    
    'Run the Add date columns to the tables beginging with Event
    For Each myTableVarList In Events_dataTablesList
        myInsertTableSQL = "INSERT INTO " & myTableVarList & " ([Trial_Card]) " _
        & "SELECT [Trial_Card] " _
        & "FROM [" & beanFinal & "]; "
        Debug.Print "myInsertTableSQL Statement: " & myInsertTableSQL & "."
        CurrentDb.Execute myInsertTableSQL
        'Debug.Print "Completed adding column Date: " & myDateVarList & "."
        'Debug.Print ("Finished Table: " & myTableVarList)
    Next myTableVarList
    
    myTableVarList = Empty
    
    ' Debug.Print vbCrLf & "Completed the Non Trial Screenings Update Query." & vbCrLf
End Sub


Public Sub ClearDataTables()
'This empties, deletes records from the All and Events data tables.
'All_dataTablesList
'Events_dataTablesList
Dim myRemoveRecordSQL As String
Dim myTableVarList As Variant  ' Used to cycle through each date column as an iterable dataTablesList()

    For Each myTableVarList In All_dataTablesList
        myRemoveRecordSQL = "DELETE * FROM [" & myTableVarList & "];"
        'Debug.Print "myRemoveRecordSQL Statement: " & myRemoveRecordSQL & "."
        CurrentDb.Execute myRemoveRecordSQL
        'Debug.Print "Completed removing column Date: " & myDateVarList & "."
        'Debug.Print ("Finished Table: " & myTableVarList)
    Next myTableVarList
    
    myTableVarList = Empty
    
    For Each myTableVarList In Events_dataTablesList
        myRemoveRecordSQL = "DELETE * FROM [" & myTableVarList & "];"
        'Debug.Print "myRemoveRecordSQL Statement: " & myRemoveRecordSQL & "."
        CurrentDb.Execute myRemoveRecordSQL
        'Debug.Print "Completed removing column Date: " & myDateVarList & "."
        'Debug.Print ("Finished Table: " & myTableVarList)
    Next myTableVarList
    
    myTableVarList = Empty
    
' Debug.Print vbCrLf & "Completed the Non Trial Screenings Update Query." & vbCrLf
End Sub


Public Sub ClearDateReportTables()
'This empties, deletes records from the weekly date data tables.
'allColumnsList
'dataTablesList
Dim myRemoveRecordSQL As String
Dim myTableVarList As Variant  ' Used to cycle through each date column as an iterable dataTablesList()

    For Each myTableVarList In allTablesList
        myRemoveRecordSQL = "DELETE * FROM [" & myTableVarList & "];"
        'Debug.Print "myRemoveRecordSQL Statement: " & myRemoveRecordSQL & "."
        CurrentDb.Execute myRemoveRecordSQL
        'Debug.Print "Completed removing column Date: " & myDateVarList & "."
        'Debug.Print ("Finished Table: " & myTableVarList)
    Next myTableVarList
    
    myTableVarList = Empty
    
    ' Debug.Print vbCrLf & "Completed the Non Trial Screenings Update Query." & vbCrLf
End Sub
