Option Compare Database
Option Explicit

Public nonTrialsList As Object ' https://excelmacromastery.com/vba-arraylist/
'allColumnsList_EventIndex
Public trialsOnlyList_BT As Long 'Index of the BT Event in the trialsOnlyList
Public trialsOnlyList_AT As Long 'Index of the AT Event in the trialsOnlyList
Public trialsOnlyList_DEL As Long 'Index of the DEL Event in the trialsOnlyList
Public trialsOnlyList_FCT As Long 'Index of the FCT Event in the trialsOnlyList
Public trialsOnlyList_OWLD As Long 'Index of the OWLD Event in the trialsOnlyList
Public trialsOnlyList_Final As Long 'Index of the Final Event in the trialsOnlyList
Public trialsOnlyList As Object 'Declare the Object handle for use
Public tablesTrialsOnlyList As Object 'Declare the Object handle for use
Public allColumnsList_BT As Long 'Index of the BT Event in the allColumnsList
Public allColumnsList_AT As Long 'Index of the AT Event in the allColumnsList
Public allColumnsList_DEL As Long 'Index of the DEL Event in the allColumnsList
Public allColumnsList_FCT As Long 'Index of the FCT Event in the allColumnsList
Public allColumnsList_OWLD As Long 'Index of the OWLD Event in the allColumnsList
Public allColumnsList_Final As Long 'Index of the Final Event in the allColumnsList
Public allColumnsList As Object 'Declare the Object handle for use
Public allTablesList As Object 'Declare the Object handle for use
Public PreAT_DatesVarList As Object 'Declare the Object handle for use
Public PreFCT_DatesVarList As Object 'Declare the Object handle for use
Public beanBT As String 'This is a varible holding the name reference to the _BT table used in the SQL statments
Public beanAT As String 'This is a varible holding the name reference to the _AT table used in the SQL statments
Public beanDEL As String 'This is a varible holding the name reference to the _DEL table used in the SQL statments
Public beanFCT As String 'This is a varible holding the name reference to the _FCT table used in the SQL statments
Public beanOWLD As String 'This is a varible holding the name reference to the _OWLD table used in the SQL statments
Public beanFinal As String 'This is a varible holding the name reference to the _Final table used in the SQL statments
Public columnBT As String ' name reference to the BT column used in the SQL statments
Public columnAT As String ' name reference to the AT column used in the SQL statments
Public columnDEL As String ' name reference to the DEL column used in the SQL statments
Public columnFCT As String ' name reference to the FCT column used in the SQL statments
Public columnOWLD As String ' name reference to the OWLD column used in the SQL statments
Public columnFinal As String ' name reference to the Final column used in the SQL statments
Public curHullNum As String ' This is the current ship hull number that is used in the table names
Public CurrentTable As String ' This is the current table that the module is modifing
Public All_dataTablesList As Object ' Declare the Object handle for use. These are the Aggreation data tables
Public All_SparseMatrixList As Object ' Declare the Object handle for use. These are the Aggreation data tables
Public Events_dataTablesList As Object ' Declare the Object handle for use. These are the Aggreation data tables
Public Events_SparseMatrixList As Object ' Declare the Object handle for use. These are the Aggreation data tables
Public All_or_Events As String ' This is a selector that switches column date lists in SQL table calls
Public SparseRefTable As String ' This is the table that the module is using to compare records columns values


Sub AddingToMyDateLists()
'https://excelmacromastery.com/vba-arraylist/

'This is overridden by useing the application CPU affinity
'DAO.DBEngine.SetOption dbMaxLocksPerFile, 1000000

'Hull number currently in use
curHullNum = "LPD27"  ' This is the current ship hull number that is used in the table names

' Tables used for reference to Events
beanBT = "2017/06/30_LPD27_BT" ' name reference to the _BT table used in the SQL statements
beanAT = "2017/08/18_LPD27_AT" '  name reference to the _AT table used in the SQL statments
beanDEL = "2017/09/15_LPD27_DEL" '  name reference to the _DEL table used in the SQL statments
beanFCT = "2018/10/26_LPD27_FCT" '  name reference to the _FCT table used in the SQL statments
beanOWLD = "2019/09/19_LPD27_OWLD" '  name reference to the _OWLD table used in the SQL statments
beanFinal = "2020/04/03_LPD27_Final" '  name reference to the _Final table used in the SQL statments

' Columns used to reference the Events
columnBT = "2017/06/30" ' name reference to the BT column used in the SQL statments
columnAT = "2017/08/18" ' name reference to the AT column used in the SQL statments
columnDEL = "2017/09/15" ' name reference to the DEL column used in the SQL statments
columnFCT = "2018/10/26" ' name reference to the FCT column used in the SQL statments
columnOWLD = "2019/09/19" ' name reference to the OWLD column used in the SQL statments
columnFinal = "2020/04/03" ' name reference to the Final column used in the SQL statments

'Columns used to reset values prior to the AT Event
'Dim PreAT_DatesVarList As Object ' This Declares the object handle, without any properties or methods
'Dim PreAT_DatesVarList As New ArrayList ' This Declares the object handle with early binding
Set PreAT_DatesVarList = CreateObject("System.Collections.ArrayList") ' This is late binding
'PreAT_DatesVarList
' Add items
PreAT_DatesVarList.Add "2017/06/30" ' BT, Called from SetLateAdds_TrialCards as PreAT_DatesVarList(0)
PreAT_DatesVarList.Add "2017/07/04" ' Called from SetLateAdds_TrialCards as PreAT_DatesVarList(1)
PreAT_DatesVarList.Add "2017/07/06" ' Called from SetLateAdds_TrialCards as PreAT_DatesVarList(2)
PreAT_DatesVarList.Add "2017/07/14" ' Called from SetLateAdds_TrialCards as PreAT_DatesVarList(3)
PreAT_DatesVarList.Add "2017/08/03" ' Called from SetLateAdds_TrialCards as PreAT_DatesVarList(4)
PreAT_DatesVarList.Add "2017/08/09" ' Called from SetLateAdds_TrialCards as PreAT_DatesVarList(5)
PreAT_DatesVarList.Add "2017/08/11" ' Called from SetLateAdds_TrialCards as PreAT_DatesVarList(6)

'Columns used to reset values prior to the FCT Event
'Dim PreFCT_DatesVarList As Object ' This Declares the object handle, without any properties or methods
'Dim PreFCT_DatesVarList As New ArrayList ' This Declares the object handle with early binding
Set PreFCT_DatesVarList = CreateObject("System.Collections.ArrayList") ' This is late binding
'PreFCT_DatesVarList
' Add items
PreFCT_DatesVarList.Add "2017/06/30" ' BT, Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(0)
PreFCT_DatesVarList.Add "2017/07/04" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(1)
PreFCT_DatesVarList.Add "2017/07/06" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(2)
PreFCT_DatesVarList.Add "2017/07/14" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(3)
PreFCT_DatesVarList.Add "2017/08/03" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(4)
PreFCT_DatesVarList.Add "2017/08/09" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(5)
PreFCT_DatesVarList.Add "2017/08/11" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(6)
PreFCT_DatesVarList.Add "2017/08/18" ' AT, Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(7)
PreFCT_DatesVarList.Add "2017/08/25" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(8)
PreFCT_DatesVarList.Add "2017/09/02" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(9)
PreFCT_DatesVarList.Add "2017/09/08" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(10)
PreFCT_DatesVarList.Add "2017/09/15" ' DEL, Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(11)
PreFCT_DatesVarList.Add "2017/09/22" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(12)
PreFCT_DatesVarList.Add "2017/09/26" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(13)
PreFCT_DatesVarList.Add "2017/09/29" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(14)
PreFCT_DatesVarList.Add "2017/10/06" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(15)
PreFCT_DatesVarList.Add "2017/10/24" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(16)
PreFCT_DatesVarList.Add "2017/11/03" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(17)
PreFCT_DatesVarList.Add "2017/11/10" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(18)
PreFCT_DatesVarList.Add "2017/11/17" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(19)
PreFCT_DatesVarList.Add "2018/01/25" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(20)
PreFCT_DatesVarList.Add "2018/04/12" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(21)
PreFCT_DatesVarList.Add "2018/08/25" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(22)
PreFCT_DatesVarList.Add "2018/09/07" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(23)
PreFCT_DatesVarList.Add "2018/09/15" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(24)
PreFCT_DatesVarList.Add "2018/09/21" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(25)
PreFCT_DatesVarList.Add "2018/09/28" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(26)
PreFCT_DatesVarList.Add "2018/10/05" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(27)
PreFCT_DatesVarList.Add "2018/10/11" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(28)
PreFCT_DatesVarList.Add "2018/10/20" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(29)

'Date values used in queries list

'Dim nonTrialsList As Object ' This Declares the object handle, without any properties or methods
'Dim nonTrialsList As New ArrayList ' This Declares the object handle with early binding
Set nonTrialsList = CreateObject("System.Collections.ArrayList") ' This is late binding
'nonTrialsList
' Add items
nonTrialsList.Add "2017/07/04"
nonTrialsList.Add "2017/07/06"
nonTrialsList.Add "2017/07/14"
nonTrialsList.Add "2017/08/03"
nonTrialsList.Add "2017/08/09"
nonTrialsList.Add "2017/08/11"
nonTrialsList.Add "2017/08/25"
nonTrialsList.Add "2017/09/02"
nonTrialsList.Add "2017/09/08"
nonTrialsList.Add "2017/09/22"
nonTrialsList.Add "2017/09/26"
nonTrialsList.Add "2017/09/29"
nonTrialsList.Add "2017/10/06"
nonTrialsList.Add "2017/10/24"
nonTrialsList.Add "2017/11/03"
nonTrialsList.Add "2017/11/10"
nonTrialsList.Add "2017/11/17"
nonTrialsList.Add "2018/01/25"
nonTrialsList.Add "2018/04/12"
nonTrialsList.Add "2018/08/25"
nonTrialsList.Add "2018/09/07"
nonTrialsList.Add "2018/09/15"
nonTrialsList.Add "2018/09/21"
nonTrialsList.Add "2018/09/28"
nonTrialsList.Add "2018/10/05"
nonTrialsList.Add "2018/10/11"
nonTrialsList.Add "2018/10/20"
nonTrialsList.Add "2018/11/04"
nonTrialsList.Add "2018/11/05"
nonTrialsList.Add "2018/11/16"
nonTrialsList.Add "2018/11/21"
nonTrialsList.Add "2018/11/30"
nonTrialsList.Add "2018/12/14"
nonTrialsList.Add "2018/12/21"
nonTrialsList.Add "2019/01/10"
nonTrialsList.Add "2019/01/25"
nonTrialsList.Add "2019/01/30"
nonTrialsList.Add "2019/02/08"
nonTrialsList.Add "2019/02/15"
nonTrialsList.Add "2019/02/22"
nonTrialsList.Add "2019/03/01"
nonTrialsList.Add "2019/03/15"
nonTrialsList.Add "2019/03/22"
nonTrialsList.Add "2019/04/01"
nonTrialsList.Add "2019/04/05"
nonTrialsList.Add "2019/04/12"
nonTrialsList.Add "2019/04/26"
nonTrialsList.Add "2019/04/30"
nonTrialsList.Add "2019/05/03"
nonTrialsList.Add "2019/05/17"
nonTrialsList.Add "2019/05/24"
nonTrialsList.Add "2019/05/30"
nonTrialsList.Add "2019/06/07"
nonTrialsList.Add "2019/06/13"
nonTrialsList.Add "2019/06/21"
nonTrialsList.Add "2019/06/24"
nonTrialsList.Add "2019/06/28"
nonTrialsList.Add "2019/07/12"
nonTrialsList.Add "2019/07/19"
nonTrialsList.Add "2019/07/26"
nonTrialsList.Add "2019/07/29"
nonTrialsList.Add "2019/08/09"
nonTrialsList.Add "2019/08/19"
nonTrialsList.Add "2019/08/20"
nonTrialsList.Add "2019/08/22"
nonTrialsList.Add "2019/08/26"
nonTrialsList.Add "2019/08/27"
nonTrialsList.Add "2019/09/05"
nonTrialsList.Add "2019/09/12"
nonTrialsList.Add "2019/09/18"
nonTrialsList.Add "2020/01/22"
nonTrialsList.Add "2020/02/07"

'trialsOnlyList_EventIndex
trialsOnlyList_BT = 0 ' Index of the BT Event in the trialsOnlyList
trialsOnlyList_AT = 1 ' Index of the AT Event in the trialsOnlyList
trialsOnlyList_DEL = 2 ' Index of the DEL Event in the trialsOnlyList
trialsOnlyList_FCT = 3 ' Index of the FCT Event in the trialsOnlyList
trialsOnlyList_OWLD = 4 ' Index of the OWLD Event in the trialsOnlyList
trialsOnlyList_Final = 5 ' Index of the Final Event in the trialsOnlyList

'Dim trialsOnlyList As Object
'Dim trialsOnlyList As New ArrayList
Set trialsOnlyList = CreateObject("System.Collections.ArrayList")
'trialsOnlyList
' Add items
trialsOnlyList.Add "2017/06/30" ' BT
trialsOnlyList.Add "2017/08/18" ' AT
trialsOnlyList.Add "2017/09/15" ' DEL
trialsOnlyList.Add "2018/10/26" ' FCT
trialsOnlyList.Add "2019/09/19" ' OWLD
trialsOnlyList.Add "2020/04/03" ' Final

'Dim tablesTrialsOnlyList As Object
'Dim tablesTrialsOnlyList As New ArrayList
Set tablesTrialsOnlyList = CreateObject("System.Collections.ArrayList")
'tablesTrialsOnlyList
' Add items
tablesTrialsOnlyList.Add "2017/06/30_LPD27_BT" ' BT
tablesTrialsOnlyList.Add "2017/08/18_LPD27_AT" ' AT
tablesTrialsOnlyList.Add "2017/09/15_LPD27_DEL" ' DEL
tablesTrialsOnlyList.Add "2018/10/26_LPD27_FCT" ' FCT
tablesTrialsOnlyList.Add "2019/09/19_LPD27_OWLD" ' OWLD
tablesTrialsOnlyList.Add "2020/04/03_LPD27_Final" ' Final

'allColumnsList_EventIndex
allColumnsList_BT = 0 ' Index of the BT Event in the allColumnsList
allColumnsList_AT = 7 ' Index of the AT Event in the allColumnsList
allColumnsList_DEL = 11 ' Index of the DEL Event in the allColumnsList
allColumnsList_FCT = 30 ' Index of the FCT Event in the allColumnsList
allColumnsList_OWLD = 74 ' Index of the OWLD Event in the allColumnsList
allColumnsList_Final = 77 ' Index of the Final Event in the allColumnsList

'Dim allColumnsList As Object
'Dim allColumnsList As New ArrayList
Set allColumnsList = CreateObject("System.Collections.ArrayList")
'allColumnsList
' Add items
allColumnsList.Add "2017/06/30" ' BT
allColumnsList.Add "2017/07/04"
allColumnsList.Add "2017/07/06"
allColumnsList.Add "2017/07/14"
allColumnsList.Add "2017/08/03"
allColumnsList.Add "2017/08/09"
allColumnsList.Add "2017/08/11"
allColumnsList.Add "2017/08/18" ' AT
allColumnsList.Add "2017/08/25"
allColumnsList.Add "2017/09/02"
allColumnsList.Add "2017/09/08"
allColumnsList.Add "2017/09/15" ' DEL
allColumnsList.Add "2017/09/22"
allColumnsList.Add "2017/09/26"
allColumnsList.Add "2017/09/29"
allColumnsList.Add "2017/10/06"
allColumnsList.Add "2017/10/24"
allColumnsList.Add "2017/11/03"
allColumnsList.Add "2017/11/10"
allColumnsList.Add "2017/11/17"
allColumnsList.Add "2018/01/25"
allColumnsList.Add "2018/04/12"
allColumnsList.Add "2018/08/25"
allColumnsList.Add "2018/09/07"
allColumnsList.Add "2018/09/15"
allColumnsList.Add "2018/09/21"
allColumnsList.Add "2018/09/28"
allColumnsList.Add "2018/10/05"
allColumnsList.Add "2018/10/11"
allColumnsList.Add "2018/10/20"
allColumnsList.Add "2018/10/26" ' FCT
allColumnsList.Add "2018/11/04"
allColumnsList.Add "2018/11/05"
allColumnsList.Add "2018/11/16"
allColumnsList.Add "2018/11/21"
allColumnsList.Add "2018/11/30"
allColumnsList.Add "2018/12/14"
allColumnsList.Add "2018/12/21"
allColumnsList.Add "2019/01/10"
allColumnsList.Add "2019/01/25"
allColumnsList.Add "2019/01/30"
allColumnsList.Add "2019/02/08"
allColumnsList.Add "2019/02/15"
allColumnsList.Add "2019/02/22"
allColumnsList.Add "2019/03/01"
allColumnsList.Add "2019/03/15"
allColumnsList.Add "2019/03/22"
allColumnsList.Add "2019/04/01"
allColumnsList.Add "2019/04/05"
allColumnsList.Add "2019/04/12"
allColumnsList.Add "2019/04/26"
allColumnsList.Add "2019/04/30"
allColumnsList.Add "2019/05/03"
allColumnsList.Add "2019/05/17"
allColumnsList.Add "2019/05/24"
allColumnsList.Add "2019/05/30"
allColumnsList.Add "2019/06/07"
allColumnsList.Add "2019/06/13"
allColumnsList.Add "2019/06/21"
allColumnsList.Add "2019/06/24"
allColumnsList.Add "2019/06/28"
allColumnsList.Add "2019/07/12"
allColumnsList.Add "2019/07/19"
allColumnsList.Add "2019/07/26"
allColumnsList.Add "2019/07/29"
allColumnsList.Add "2019/08/09"
allColumnsList.Add "2019/08/19"
allColumnsList.Add "2019/08/20"
allColumnsList.Add "2019/08/22"
allColumnsList.Add "2019/08/26"
allColumnsList.Add "2019/08/27"
allColumnsList.Add "2019/09/05"
allColumnsList.Add "2019/09/12"
allColumnsList.Add "2019/09/18"
allColumnsList.Add "2019/09/19" ' OWLD
allColumnsList.Add "2020/01/22"
allColumnsList.Add "2020/02/07"
allColumnsList.Add "2020/04/03" ' Final

'Dim allTablesList As Object
'Dim allTablesList As New ArrayList
Set allTablesList = CreateObject("System.Collections.ArrayList")
'allTablesList
' Add items
allTablesList.Add "2017/06/30_LPD27_BT"
allTablesList.Add "2017/07/04_LPD27"
allTablesList.Add "2017/07/06_LPD27"
allTablesList.Add "2017/07/14_LPD27"
allTablesList.Add "2017/08/03_LPD27"
allTablesList.Add "2017/08/09_LPD27"
allTablesList.Add "2017/08/11_LPD27"
allTablesList.Add "2017/08/18_LPD27_AT"
allTablesList.Add "2017/08/25_LPD27"
allTablesList.Add "2017/09/02_LPD27"
allTablesList.Add "2017/09/08_LPD27"
allTablesList.Add "2017/09/15_LPD27_DEL"
allTablesList.Add "2017/09/22_LPD27"
allTablesList.Add "2017/09/26_LPD27"
allTablesList.Add "2017/09/29_LPD27"
allTablesList.Add "2017/10/06_LPD27"
allTablesList.Add "2017/10/24_LPD27"
allTablesList.Add "2017/11/03_LPD27"
allTablesList.Add "2017/11/10_LPD27"
allTablesList.Add "2017/11/17_LPD27"
allTablesList.Add "2018/01/25_LPD27"
allTablesList.Add "2018/04/12_LPD27"
allTablesList.Add "2018/08/25_LPD27"
allTablesList.Add "2018/09/07_LPD27"
allTablesList.Add "2018/09/15_LPD27"
allTablesList.Add "2018/09/21_LPD27"
allTablesList.Add "2018/09/28_LPD27"
allTablesList.Add "2018/10/05_LPD27"
allTablesList.Add "2018/10/11_LPD27"
allTablesList.Add "2018/10/20_LPD27"
allTablesList.Add "2018/10/26_LPD27_FCT"
allTablesList.Add "2018/11/04_LPD27"
allTablesList.Add "2018/11/05_LPD27"
allTablesList.Add "2018/11/16_LPD27"
allTablesList.Add "2018/11/21_LPD27"
allTablesList.Add "2018/11/30_LPD27"
allTablesList.Add "2018/12/14_LPD27"
allTablesList.Add "2018/12/21_LPD27"
allTablesList.Add "2019/01/10_LPD27"
allTablesList.Add "2019/01/25_LPD27"
allTablesList.Add "2019/01/30_LPD27"
allTablesList.Add "2019/02/08_LPD27"
allTablesList.Add "2019/02/15_LPD27"
allTablesList.Add "2019/02/22_LPD27"
allTablesList.Add "2019/03/01_LPD27"
allTablesList.Add "2019/03/15_LPD27"
allTablesList.Add "2019/03/22_LPD27"
allTablesList.Add "2019/04/01_LPD27"
allTablesList.Add "2019/04/05_LPD27"
allTablesList.Add "2019/04/12_LPD27"
allTablesList.Add "2019/04/26_LPD27"
allTablesList.Add "2019/04/30_LPD27"
allTablesList.Add "2019/05/03_LPD27"
allTablesList.Add "2019/05/17_LPD27"
allTablesList.Add "2019/05/24_LPD27"
allTablesList.Add "2019/05/30_LPD27"
allTablesList.Add "2019/06/07_LPD27"
allTablesList.Add "2019/06/13_LPD27"
allTablesList.Add "2019/06/21_LPD27"
allTablesList.Add "2019/06/24_LPD27"
allTablesList.Add "2019/06/28_LPD27"
allTablesList.Add "2019/07/12_LPD27"
allTablesList.Add "2019/07/19_LPD27"
allTablesList.Add "2019/07/26_LPD27"
allTablesList.Add "2019/07/29_LPD27"
allTablesList.Add "2019/08/09_LPD27"
allTablesList.Add "2019/08/19_LPD27"
allTablesList.Add "2019/08/20_LPD27"
allTablesList.Add "2019/08/22_LPD27"
allTablesList.Add "2019/08/26_LPD27"
allTablesList.Add "2019/08/27_LPD27"
allTablesList.Add "2019/09/05_LPD27"
allTablesList.Add "2019/09/12_LPD27"
allTablesList.Add "2019/09/18_LPD27"
allTablesList.Add "2019/09/19_LPD27_OWLD"
allTablesList.Add "2020/01/22_LPD27"
allTablesList.Add "2020/02/07_LPD27"
allTablesList.Add "2020/04/03_LPD27_Final"

'Dim All_dataTablesList As Object
'Dim All_dataTablesList As New ArrayList
Set All_dataTablesList = CreateObject("System.Collections.ArrayList")
'All_dataTablesList
' Add items
All_dataTablesList.Add "All_Combined_Screenings" ' All_dataTablesList(0)
All_dataTablesList.Add "All_Combined_Screenings_SparseMatrix" ' All_dataTablesList(1)
All_dataTablesList.Add "All_Screenings_Only" ' All_dataTablesList(2)
All_dataTablesList.Add "All_Screenings_Only_SparseMatrix" ' All_dataTablesList(3)
All_dataTablesList.Add "All_TC_Screen_Agg" ' All_dataTablesList(4)
All_dataTablesList.Add "All_XX_Screen_Only" ' All_dataTablesList(5)
All_dataTablesList.Add "All_XX_Screen_Only_SparseMatrix" ' All_dataTablesList(6)
All_dataTablesList.Add "All_Z_Summary" ' All_dataTablesList(7)

Set All_SparseMatrixList = CreateObject("System.Collections.ArrayList")
All_SparseMatrixList.Add "All_Combined_Screenings_SparseMatrix" ' All_SparseMatrixList(0)
All_SparseMatrixList.Add "All_Screenings_Only_SparseMatrix" ' All_SparseMatrixList(1)
All_SparseMatrixList.Add "All_XX_Screen_Only_SparseMatrix" ' All_SparseMatrixList(2)

'Dim Events_dataTablesList As Object
'Dim Events_dataTablesList As New ArrayList
Set Events_dataTablesList = CreateObject("System.Collections.ArrayList")
'Events_dataTablesList
' Add items
Events_dataTablesList.Add "Events_Combined_Screenings" ' dataTablesList(0)
Events_dataTablesList.Add "Events_Combined_Screenings_SparseMatrix" ' dataTablesList(1)
Events_dataTablesList.Add "Events_Screenings_Only" ' dataTablesList(2)
Events_dataTablesList.Add "Events_Screenings_Only_SparseMatrix" ' dataTablesList(3)
Events_dataTablesList.Add "Events_TC_Screen_Agg" ' dataTablesList(4)
Events_dataTablesList.Add "Events_XX_Screen_Only" ' dataTablesList(5)
Events_dataTablesList.Add "Events_XX_Screen_Only_SparseMatrix" ' dataTablesList(6)

Set Events_SparseMatrixList = CreateObject("System.Collections.ArrayList")
Events_SparseMatrixList.Add "Events_Combined_Screenings_SparseMatrix" ' Events_SparseMatrixList(0)
Events_SparseMatrixList.Add "Events_Screenings_Only_SparseMatrix" ' Events_SparseMatrixList(1)
Events_SparseMatrixList.Add "Events_XX_Screen_Only_SparseMatrix" ' Events_SparseMatrixList(2)

' Insert to first position
'   allTablesList.Insert 0, "2020/01/22_LPD23"
' Print a specfic item
'   allTablesList(i)
' Sort
'allTablesList.Sort
'   Debug.Print vbCrLf & "Sorted Ascending"
    ' Add this sub from "Reading through the items" section
'   PrintToImmediateWindow allTablesList
' Reverse sort
'   allTablesList.Reverse
'   Debug.Print vbCrLf & "Sorted Descending"
'   PrintToImmediateWindow allTablesList
' Remove all item
'   allTablesList.Clear

End Sub

Public Sub ClearMyDateLists()
'This is to empty the lists and make sure no double entry happens

PreAT_DatesVarList.Clear
PreFCT_DatesVarList.Clear
nonTrialsList.Clear
trialsOnlyList.Clear
tablesTrialsOnlyList.Clear
allColumnsList.Clear
allTablesList.Clear
All_dataTablesList.Clear
All_SparseMatrixList.Clear
Events_dataTablesList.Clear
Events_SparseMatrixList.Clear

End Sub
