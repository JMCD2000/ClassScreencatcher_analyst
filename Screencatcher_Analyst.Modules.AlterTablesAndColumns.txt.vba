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


Public Sub CreateDateReportTables()
'This creates tables from the weekly reports date data in allTablesList.
'dataTablesList
Dim myTableVarList As Variant  ' Used to cycle through each date column as an iterable dataTablesList()

    For Each myTableVarList In allTablesList
        MakeNewReportsTables myTableVarList
        Debug.Print ("Finished Table: " & myTableVarList)
    Next myTableVarList
    
    myTableVarList = Empty
    
End Sub


Public Sub MakeNewReportsTables(ByVal myNewTable As String)

Dim db As DAO.Database
Dim tdf As DAO.TableDef
Dim fld_TC As DAO.Field ' Trial_Card CHAR
Dim fld_Star As DAO.Field ' Star CHAR
Dim fld_Pri As DAO.Field ' Priority CHAR
Dim fld_Safety As DAO.Field ' Safety CHAR
Dim fld_Screening As DAO.Field ' Screening CHAR
Dim fld_Act_1 As DAO.Field ' Act_1 CHAR
Dim fld_Act_2 As DAO.Field ' Act_2 CHAR
Dim fld_Status As DAO.Field ' Status CHAR
Dim fld_Action_Taken As DAO.Field ' Action_Taken CHAR
Dim fld_Date_Discovered As DAO.Field ' Date_Discovered DATE
Dim fld_Date_Closed As DAO.Field ' Date_Closed DATE
Dim fld_Trial_ID As DAO.Field ' Trial_ID CHAR
Dim fld_Event As DAO.Field ' Event CHAR
Dim fld_TC_Screening As DAO.Field ' TC_Screening CHAR
Dim fld_TC_Screening_AC1_AC2 As DAO.Field ' TC_Screening_AC1_AC2 CHAR
'Dim fld_Final_Sts_A_T As DAO.Field ' TC_Screening_AC1_AC2 CHAR

'Open connection to the current database
Set db = CurrentDb

'Create the new table object
Set tdf = db.CreateTableDef(myNewTable)

'Create the new field objects
Set fld_TC = tdf.CreateField("Trial_Card", dbText, 250)
Set fld_Star = tdf.CreateField("Star", dbText, 250)
Set fld_Pri = tdf.CreateField("Pri", dbText, 250)
Set fld_Safety = tdf.CreateField("Safety", dbText, 250)
Set fld_Screening = tdf.CreateField("Screening", dbText, 250)
Set fld_Act_1 = tdf.CreateField("Act_1", dbText, 250)
Set fld_Act_2 = tdf.CreateField("Act_2", dbText, 250)
Set fld_Status = tdf.CreateField("Status", dbText, 250)
Set fld_Action_Taken = tdf.CreateField("Action_Taken", dbText, 250)
Set fld_Date_Discovered = tdf.CreateField("Date_Discovered", dbDate, 250)
Set fld_Date_Closed = tdf.CreateField("Date_Closed", dbDate, 250)
Set fld_Trial_ID = tdf.CreateField("Trial_ID", dbText, 250)
Set fld_Event = tdf.CreateField("Event", dbText, 250)
Set fld_TC_Screening = tdf.CreateField("TC_Screening", dbText, 250)
Set fld_TC_Screening_AC1_AC2 = tdf.CreateField("TC_Screening_AC1_AC2", dbText, 250)
'Set fld_Final_Sts_A_T = tdf.CreateField("Final_Sts_A_T", dbText, 250)

tdf.Fields.Append fld_TC
tdf.Fields.Append fld_Star
tdf.Fields.Append fld_Pri
tdf.Fields.Append fld_Safety
tdf.Fields.Append fld_Screening
tdf.Fields.Append fld_Act_1
tdf.Fields.Append fld_Act_2
tdf.Fields.Append fld_Action_Taken
tdf.Fields.Append fld_Date_Discovered
tdf.Fields.Append fld_Date_Closed
tdf.Fields.Append fld_Trial_ID
tdf.Fields.Append fld_Event
tdf.Fields.Append fld_TC_Screening
tdf.Fields.Append fld_TC_Screening_AC1_AC2
'tdf.Fields.Append fld_Final_Sts_A_T

'Add the table to the current database
db.TableDefs.Append tdf

'Refresh current database
db.TableDefs.Refresh
Application.RefreshDatabaseWindow

'Empty the objects
Set fld_TC = Nothing
Set fld_Star = Nothing
Set fld_Pri = Nothing
Set fld_Safety = Nothing
Set fld_Screening = Nothing
Set fld_Act_1 = Nothing
Set fld_Act_2 = Nothing
Set fld_Status = Nothing
Set fld_Action_Taken = Nothing
Set fld_Date_Discovered = Nothing
Set fld_Date_Closed = Nothing
Set fld_Trial_ID = Nothing
Set fld_Event = Nothing
Set fld_TC_Screening = Nothing
Set fld_TC_Screening_AC1_AC2 = Nothing
'Set fld_Final_Sts_A_T = Nothing
Set tdf = Nothing
Set db = Nothing

MakeNewReportsTablesIndex myNewTable

End Sub


Public Sub MakeNewReportsTablesIndex(myNewTable As String)

Dim db As DAO.Database
Dim tdf As DAO.TableDef
Dim fld_Index As DAO.Index
Dim fld_TC As DAO.Field ' Trial_Card CHAR

Set db = CurrentDb
Set tdf = db.TableDefs(myNewTable)

'Create Index property
Set fld_Index = tdf.CreateIndex("PrimaryKey")
fld_Index.Primary = True
fld_Index.Required = True
fld_Index.Unique = True

'Create the new field objects
Set fld_TC = fld_Index.CreateField("Trial_Card", dbText, 250)

'Append the fields to the table
fld_Index.Fields.Append fld_TC
tdf.Indexes.Append fld_Index

tdf.Indexes.Refresh

'Set fld_Index = Empty
'Set fld_TC = Empty
'Set tdf = Empty
Set db = Nothing

End Sub
