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
        If myTableVarList = "All_Z_Summary" Then
            'This is commented out because it doesn't contain trial card rows
            'Pass 'Next myTableVarList
        Else
            myInsertTableSQL = "INSERT INTO " & myTableVarList & " ([Trial_Card]) " _
            & "SELECT [Trial_Card] " _
            & "FROM [" & beanFinal & "]; "
            Debug.Print "myInsertTableSQL Statement: " & myInsertTableSQL & "."
            CurrentDb.Execute myInsertTableSQL
            'Debug.Print "Completed adding column Date: " & myDateVarList & "."
            'Debug.Print ("Finished Table: " & myTableVarList)
        End If
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
        If myTableVarList = "All_Z_Summary" Then
            'This is commented out because it doesn't contain trial card rows
            'Pass 'Next myTableVarList
        Else
            myRemoveRecordSQL = "DELETE * FROM [" & myTableVarList & "];"
            'Debug.Print "myRemoveRecordSQL Statement: " & myRemoveRecordSQL & "."
            CurrentDb.Execute myRemoveRecordSQL
            'Debug.Print "Completed removing column Date: " & myDateVarList & "."
            'Debug.Print ("Finished Table: " & myTableVarList)
        End If
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
Dim tblEventFinalCheck As String 'name reference to the Final report
Dim fld_Final_Sts_A_T As DAO.Field ' TC_Screening_AC1_AC2 CHAR

'Get the right side of the passed in table name
tblEventFinalCheck = Right(myNewTable, 6) ' "2020/04/03_LPD26_Final"

'Open connection to the current database
Set db = CurrentDb

'Create the new table object
Set tdf = db.CreateTableDef(myNewTable)

'Create the new field objects
Set fld_TC = tdf.CreateField("Trial_Card", dbText, 250)
Set fld_Star = tdf.CreateField("Star", dbText, 250)
    fld_Star.AllowZeroLength = True
Set fld_Pri = tdf.CreateField("Pri", dbText, 250)
    fld_Pri.AllowZeroLength = True
Set fld_Safety = tdf.CreateField("Safety", dbText, 250)
    fld_Safety.AllowZeroLength = True
Set fld_Screening = tdf.CreateField("Screening", dbText, 250)
    fld_Screening.AllowZeroLength = True
Set fld_Act_1 = tdf.CreateField("Act_1", dbText, 250)
    fld_Act_1.AllowZeroLength = True
Set fld_Act_2 = tdf.CreateField("Act_2", dbText, 250)
    fld_Act_2.AllowZeroLength = True
Set fld_Status = tdf.CreateField("Status", dbText, 250)
    fld_Status.AllowZeroLength = True
Set fld_Action_Taken = tdf.CreateField("Action_Taken", dbText, 250)
    fld_Action_Taken.AllowZeroLength = True
Set fld_Date_Discovered = tdf.CreateField("Date_Discovered", dbDate, 250)
    'fld_Date_Discovered.AllowZeroLength = True
Set fld_Date_Closed = tdf.CreateField("Date_Closed", dbDate, 250)
    'fld_Date_Closed.AllowZeroLength = True
Set fld_Trial_ID = tdf.CreateField("Trial_ID", dbText, 250)
    fld_Trial_ID.AllowZeroLength = True
Set fld_Event = tdf.CreateField("Event", dbText, 250)
    fld_Event.AllowZeroLength = True
Set fld_TC_Screening = tdf.CreateField("TC_Screening", dbText, 250)
    fld_TC_Screening.AllowZeroLength = True
Set fld_TC_Screening_AC1_AC2 = tdf.CreateField("TC_Screening_AC1_AC2", dbText, 250)
    fld_TC_Screening_AC1_AC2.AllowZeroLength = True

'The table that has the final report gets this third column
If tblEventFinalCheck = "_Final" Then
    'The Table is the Final Event Table
    Set fld_Final_Sts_A_T = tdf.CreateField("Final_Sts_A_T", dbText, 250)
        fld_Final_Sts_A_T.AllowZeroLength = True
ElseIf tblEventFinalCheck <> "_Final" Then
    'The Table is not the Final Event Table
    'do nothing
Else
    'untraped error
End If

tdf.Fields.Append fld_TC
tdf.Fields.Append fld_Star
tdf.Fields.Append fld_Pri
tdf.Fields.Append fld_Safety
tdf.Fields.Append fld_Screening
tdf.Fields.Append fld_Act_1
tdf.Fields.Append fld_Act_2
tdf.Fields.Append fld_Status
tdf.Fields.Append fld_Action_Taken
tdf.Fields.Append fld_Date_Discovered
tdf.Fields.Append fld_Date_Closed
tdf.Fields.Append fld_Trial_ID
tdf.Fields.Append fld_Event
tdf.Fields.Append fld_TC_Screening
tdf.Fields.Append fld_TC_Screening_AC1_AC2
'The table that has the final report gets this third column
If tblEventFinalCheck = "_Final" Then
    'The Table is the Final Event Table
    tdf.Fields.Append fld_Final_Sts_A_T
ElseIf tblEventFinalCheck <> "_Final" Then
    'The Table is not the Final Event Table
    'do nothing
Else
    'untraped error
End If

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

'The table that has the final report gets this third column
If tblEventFinalCheck = "_Final" Then
    'The Table is the Final Event Table
    Set fld_Final_Sts_A_T = Nothing
ElseIf tblEventFinalCheck <> "_Final" Then
    'The Table is not the Final Event Table
    'do nothing
Else
    'untraped error
End If

Set tdf = Nothing
Set db = Nothing

MakeNewReportsTablesIndex myNewTable 'Set table Index and Primary Key
SetNewReportsTablesProps myNewTable 'Set Date Field Properties

End Sub


Public Sub MakeNewReportsTablesIndex(myNewTable As String)
'Set table Index and Primary Key
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


Public Sub SetNewReportsTablesProps(myNewTable As String)
'Set Date Field Properties
Dim db As DAO.Database
Dim tdf As DAO.TableDef
Dim prpDD As Property
Dim prpDC As Property
Dim prpName As String
Dim prpValue As String
Dim prpType As Long

Set db = CurrentDb
Set tdf = db.TableDefs(myNewTable)

'Create property
prpName = "Format"
prpValue = "m/d/yyyy"
prpType = dbText

' If the DAO object has the built-in Property objects then use the below assignment.
'tdf.Fields("Date_Discovered").Properties(prpName) = prpValue
tdf.Fields("Pri").Required = True
tdf.Fields("Screening").Required = True
tdf.Fields("Status").Required = True
tdf.Fields("Date_Discovered").Required = True
tdf.Fields("Event").Required = True

' If the DAO object doesn't have the built-in Property objects then use the below assignment.
Set prpDD = tdf.Fields("Date_Discovered").CreateProperty(prpName, prpType, prpValue)
Set prpDC = tdf.Fields("Date_Closed").CreateProperty(prpName, prpType, prpValue)

'Append the fields to the table
'Error 3367 if it is already been Appended
tdf.Fields("Date_Discovered").Properties.Append prpDD
tdf.Fields("Date_Closed").Properties.Append prpDC

Set db = Nothing

End Sub

Public Sub BackfitSetReportsTablesProps()
'Set Date Field Properties
Dim myTableVarList As Variant  ' Used to cycle through each date column as an iterable dataTablesList()
Dim db As DAO.Database
Dim tdf As DAO.TableDef
Dim prpDD As Property
Dim prpDC As Property
Dim prpName As String
Dim prpValue As String
Dim prpType As Long

AddingToMyDateLists

Set db = CurrentDb

For Each myTableVarList In allTablesList
    Set tdf = db.TableDefs(myTableVarList)

    'Create property
    prpName = "Format"
    prpValue = "m/d/yyyy"
    prpType = dbText
    
    ' If the DAO object has the built-in Property objects then use the below assignment.
    'tdf.Fields("Date_Discovered").Properties(prpName) = prpValue
    tdf.Fields("Pri").Required = True
    tdf.Fields("Screening").Required = True
    tdf.Fields("Status").Required = True
    tdf.Fields("Date_Discovered").Required = True
    tdf.Fields("Event").Required = True
    
    ' If the DAO object doesn't have the built-in Property objects then use the below assignment.
    Set prpDD = tdf.Fields("Date_Discovered").CreateProperty(prpName, prpType, prpValue)
    Set prpDC = tdf.Fields("Date_Closed").CreateProperty(prpName, prpType, prpValue)
    
    'Append the fields to the table
    'Error 3367 if it is already been Appended
    tdf.Fields("Date_Discovered").Properties.Append prpDD
    tdf.Fields("Date_Closed").Properties.Append prpDC

    Debug.Print ("Finished Table: " & myTableVarList)
    
Next myTableVarList
    
myTableVarList = Empty

Set db = Nothing

End Sub
