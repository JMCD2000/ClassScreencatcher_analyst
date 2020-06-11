Option Compare Database
Option Explicit

Public nonTrialsList As Object ' https://excelmacromastery.com/vba-arraylist/
Public trialsOnlyList As Object 'Declare the Object handle for use
Public tablesTrialsOnlyList As Object 'Declare the Object handle for use
Public allColumnsList_BT As Long 'Index of the BT Event in the allColumnsList
Public allColumnsList_AT As Long 'Index of the AT Event in the allColumnsList
Public allColumnsList_FCT As Long 'Index of the FCT Event in the allColumnsList
Public allColumnsList_OWLD As Long 'Index of the OWLD Event in the allColumnsList
Public allColumnsList_Final As Long 'Index of the Final Event in the allColumnsList
Public allColumnsList As Object 'Declare the Object handle for use
Public allTablesList As Object 'Declare the Object handle for use
Public PreAT_DatesVarList As Object 'Declare the Object handle for use
Public PreFCT_DatesVarList As Object 'Declare the Object handle for use
Public beanBT As String 'This is a varible holding the name reference to the _BT table used in the SQL statments
Public beanAT As String 'This is a varible holding the name reference to the _AT table used in the SQL statments
Public beanFCT As String 'This is a varible holding the name reference to the _FCT table used in the SQL statments
Public beanOWLD As String 'This is a varible holding the name reference to the _OWLD table used in the SQL statments
Public beanFinal As String 'This is a varible holding the name reference to the _Final table used in the SQL statments
Public columnBT As String ' name reference to the BT column used in the SQL statments
Public columnAT As String ' name reference to the AT column used in the SQL statments
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
curHullNum = "LPD26"  ' This is the current ship hull number that is used in the table names

' Tables used for reference to Events
beanBT = "2016/03/06_LPD26_BT" ' name reference to the _BT table used in the SQL statements
beanAT = "2016/04/15_LPD26_AT" '  name reference to the _AT table used in the SQL statments
beanFCT = "2017/07/22_LPD26_FCT" '  name reference to the _FCT table used in the SQL statments
beanOWLD = "2018/05/04_LPD26_OWLD" '  name reference to the _OWLD table used in the SQL statments
beanFinal = "2020/04/03_LPD26_Final" '  name reference to the _Final table used in the SQL statments

' Columns used to reference the Events
columnBT = "2016/03/06" ' name reference to the BT column used in the SQL statments
columnAT = "2016/04/15" ' name reference to the AT column used in the SQL statments
columnFCT = "2017/07/22" ' name reference to the FCT column used in the SQL statments
columnOWLD = "2018/05/04" ' name reference to the OWLD column used in the SQL statments
columnFinal = "2020/04/03" ' name reference to the Final column used in the SQL statments

'Columns used to reset values prior to the AT Event
'Dim PreAT_DatesVarList As Object ' This Declares the object handle, without any properties or methods
'Dim PreAT_DatesVarList As New ArrayList ' This Declares the object handle with early binding
Set PreAT_DatesVarList = CreateObject("System.Collections.ArrayList") ' This is late binding
'PreAT_DatesVarList
' Add items
PreAT_DatesVarList.Add "2016/03/06" ' BT, Called from SetLateAdds_TrialCards as PreAT_DatesVarList(0)
PreAT_DatesVarList.Add "2016/04/08" ' Called from SetLateAdds_TrialCards as PreAT_DatesVarList(1)
PreAT_DatesVarList.Add "2016/04/09" ' Called from SetLateAdds_TrialCards as PreAT_DatesVarList(2)
PreAT_DatesVarList.Add "2016/04/10" ' Called from SetLateAdds_TrialCards as PreAT_DatesVarList(3)

'Columns used to reset values prior to the FCT Event
'Dim PreFCT_DatesVarList As Object ' This Declares the object handle, without any properties or methods
'Dim PreFCT_DatesVarList As New ArrayList ' This Declares the object handle with early binding
Set PreFCT_DatesVarList = CreateObject("System.Collections.ArrayList") ' This is late binding
'PreFCT_DatesVarList
' Add items
PreFCT_DatesVarList.Add "2016/03/06" ' BT, Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(0)
PreFCT_DatesVarList.Add "2016/04/08" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(1)
PreFCT_DatesVarList.Add "2016/04/09" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(2)
PreFCT_DatesVarList.Add "2016/04/10" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(3)
PreFCT_DatesVarList.Add "2016/04/15" ' AT, Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(4)
PreFCT_DatesVarList.Add "2016/04/22" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(5)
PreFCT_DatesVarList.Add "2016/05/19" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(6)
PreFCT_DatesVarList.Add "2016/05/27" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(7)
PreFCT_DatesVarList.Add "2016/06/03" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(8)
PreFCT_DatesVarList.Add "2016/06/17" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(9)
PreFCT_DatesVarList.Add "2016/06/24" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(10)
PreFCT_DatesVarList.Add "2016/07/01" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(11)
PreFCT_DatesVarList.Add "2016/07/08" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(12)
PreFCT_DatesVarList.Add "2016/07/12" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(13)
PreFCT_DatesVarList.Add "2016/07/16" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(14)
PreFCT_DatesVarList.Add "2016/07/22" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(15)
PreFCT_DatesVarList.Add "2016/07/29" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(16)
PreFCT_DatesVarList.Add "2016/08/05" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(17)
PreFCT_DatesVarList.Add "2016/08/12" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(18)
PreFCT_DatesVarList.Add "2016/08/19" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(19)
PreFCT_DatesVarList.Add "2016/08/26" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(20)
PreFCT_DatesVarList.Add "2016/09/02" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(21)
PreFCT_DatesVarList.Add "2016/09/09" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(22)
PreFCT_DatesVarList.Add "2016/09/16" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(23)
PreFCT_DatesVarList.Add "2016/09/24" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(24)
PreFCT_DatesVarList.Add "2016/09/30" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(25)
PreFCT_DatesVarList.Add "2016/10/07" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(26)
PreFCT_DatesVarList.Add "2016/10/21" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(27)
PreFCT_DatesVarList.Add "2016/10/28" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(28)
PreFCT_DatesVarList.Add "2016/11/04" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(29)
PreFCT_DatesVarList.Add "2016/11/14" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(30)
PreFCT_DatesVarList.Add "2016/11/18" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(31)
PreFCT_DatesVarList.Add "2016/12/03" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(32)
PreFCT_DatesVarList.Add "2016/12/09" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(33)
PreFCT_DatesVarList.Add "2016/12/16" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(34)
PreFCT_DatesVarList.Add "2016/12/22" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(35)
PreFCT_DatesVarList.Add "2017/01/06" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(36)
PreFCT_DatesVarList.Add "2017/01/13" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(37)
PreFCT_DatesVarList.Add "2017/01/20" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(38)
PreFCT_DatesVarList.Add "2017/01/27" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(39)
PreFCT_DatesVarList.Add "2017/02/03" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(40)
PreFCT_DatesVarList.Add "2017/02/10" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(41)
PreFCT_DatesVarList.Add "2017/02/17" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(42)
PreFCT_DatesVarList.Add "2017/02/24" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(43)
PreFCT_DatesVarList.Add "2017/03/03" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(44)
PreFCT_DatesVarList.Add "2017/03/10" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(45)
PreFCT_DatesVarList.Add "2017/03/17" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(46)
PreFCT_DatesVarList.Add "2017/03/24" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(47)
PreFCT_DatesVarList.Add "2017/03/31" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(48)
PreFCT_DatesVarList.Add "2017/04/07" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(49)
PreFCT_DatesVarList.Add "2017/04/14" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(50)
PreFCT_DatesVarList.Add "2017/04/21" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(51)
PreFCT_DatesVarList.Add "2017/04/28" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(52)
PreFCT_DatesVarList.Add "2017/05/12" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(53)
PreFCT_DatesVarList.Add "2017/05/19" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(54)
PreFCT_DatesVarList.Add "2017/05/24" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(55)
PreFCT_DatesVarList.Add "2017/06/02" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(56)
PreFCT_DatesVarList.Add "2017/06/09" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(57)
PreFCT_DatesVarList.Add "2017/06/16" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(58)
PreFCT_DatesVarList.Add "2017/06/23" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(59)
PreFCT_DatesVarList.Add "2017/07/04" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(60)
PreFCT_DatesVarList.Add "2017/07/07" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(61)
PreFCT_DatesVarList.Add "2017/07/14" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(62)

'Date values used in queries list

'Dim nonTrialsList As Object ' This Declares the object handle, without any properties or methods
'Dim nonTrialsList As New ArrayList ' This Declares the object handle with early binding
Set nonTrialsList = CreateObject("System.Collections.ArrayList") ' This is late binding
'nonTrialsList
' Add items
nonTrialsList.Add "2016/04/08"
nonTrialsList.Add "2016/04/09"
nonTrialsList.Add "2016/04/10"
nonTrialsList.Add "2016/04/22"
nonTrialsList.Add "2016/05/19"
nonTrialsList.Add "2016/05/27"
nonTrialsList.Add "2016/06/03"
nonTrialsList.Add "2016/06/17"
nonTrialsList.Add "2016/06/24"
nonTrialsList.Add "2016/07/01"
nonTrialsList.Add "2016/07/08"
nonTrialsList.Add "2016/07/12"
nonTrialsList.Add "2016/07/16"
nonTrialsList.Add "2016/07/22"
nonTrialsList.Add "2016/07/29"
nonTrialsList.Add "2016/08/05"
nonTrialsList.Add "2016/08/12"
nonTrialsList.Add "2016/08/19"
nonTrialsList.Add "2016/08/26"
nonTrialsList.Add "2016/09/02"
nonTrialsList.Add "2016/09/09"
nonTrialsList.Add "2016/09/16"
nonTrialsList.Add "2016/09/24"
nonTrialsList.Add "2016/09/30"
nonTrialsList.Add "2016/10/07"
nonTrialsList.Add "2016/10/21"
nonTrialsList.Add "2016/10/28"
nonTrialsList.Add "2016/11/04"
nonTrialsList.Add "2016/11/14"
nonTrialsList.Add "2016/11/18"
nonTrialsList.Add "2016/12/03"
nonTrialsList.Add "2016/12/09"
nonTrialsList.Add "2016/12/16"
nonTrialsList.Add "2016/12/22"
nonTrialsList.Add "2017/01/06"
nonTrialsList.Add "2017/01/13"
nonTrialsList.Add "2017/01/20"
nonTrialsList.Add "2017/01/27"
nonTrialsList.Add "2017/02/03"
nonTrialsList.Add "2017/02/10"
nonTrialsList.Add "2017/02/17"
nonTrialsList.Add "2017/02/24"
nonTrialsList.Add "2017/03/03"
nonTrialsList.Add "2017/03/10"
nonTrialsList.Add "2017/03/17"
nonTrialsList.Add "2017/03/24"
nonTrialsList.Add "2017/03/31"
nonTrialsList.Add "2017/04/07"
nonTrialsList.Add "2017/04/14"
nonTrialsList.Add "2017/04/21"
nonTrialsList.Add "2017/04/28"
nonTrialsList.Add "2017/05/12"
nonTrialsList.Add "2017/05/19"
nonTrialsList.Add "2017/05/24"
nonTrialsList.Add "2017/06/02"
nonTrialsList.Add "2017/06/09"
nonTrialsList.Add "2017/06/16"
nonTrialsList.Add "2017/06/23"
nonTrialsList.Add "2017/07/04"
nonTrialsList.Add "2017/07/07"
nonTrialsList.Add "2017/07/14"
nonTrialsList.Add "2017/07/28"
nonTrialsList.Add "2017/08/03"
nonTrialsList.Add "2017/08/09"
nonTrialsList.Add "2017/08/11"
nonTrialsList.Add "2017/08/19"
nonTrialsList.Add "2017/08/25"
nonTrialsList.Add "2017/09/02"
nonTrialsList.Add "2017/09/08"
nonTrialsList.Add "2017/09/15"
nonTrialsList.Add "2017/09/26"
nonTrialsList.Add "2017/09/29"
nonTrialsList.Add "2017/10/06"
nonTrialsList.Add "2017/10/24"
nonTrialsList.Add "2017/11/14"
nonTrialsList.Add "2017/12/18"
nonTrialsList.Add "2018/01/25"
nonTrialsList.Add "2018/04/04"
nonTrialsList.Add "2018/04/12"
nonTrialsList.Add "2019/04/01"
nonTrialsList.Add "2019/09/05"
nonTrialsList.Add "2020/01/22"

'Dim trialsOnlyList As Object
'Dim trialsOnlyList As New ArrayList
Set trialsOnlyList = CreateObject("System.Collections.ArrayList")
'trialsOnlyList
' Add items
trialsOnlyList.Add "2016/03/06" ' BT
trialsOnlyList.Add "2016/04/15" ' AT
trialsOnlyList.Add "2017/07/22" ' FCT
trialsOnlyList.Add "2018/05/04" ' OWLD
trialsOnlyList.Add "2020/04/03" ' Final

'Dim tablesTrialsOnlyList As Object
'Dim tablesTrialsOnlyList As New ArrayList
Set tablesTrialsOnlyList = CreateObject("System.Collections.ArrayList")
'tablesTrialsOnlyList
' Add items
tablesTrialsOnlyList.Add "2016/03/06_LPD26_BT" ' BT
tablesTrialsOnlyList.Add "2016/04/15_LPD26_AT" ' AT
tablesTrialsOnlyList.Add "2017/07/22_LPD26_FCT" ' FCT
tablesTrialsOnlyList.Add "2018/05/04_LPD26_OWLD" ' OWLD
tablesTrialsOnlyList.Add "2020/04/03_LPD26_Final" ' Final

'allColumnsList_EventIndex
allColumnsList_BT = 0 ' Index of the BT Event in the allColumnsList
allColumnsList_AT = 4 ' Index of the AT Event in the allColumnsList
allColumnsList_FCT = 63 ' Index of the FCT Event in the allColumnsList
allColumnsList_OWLD = 82 ' Index of the OWLD Event in the allColumnsList
allColumnsList_Final = 86 ' Index of the Final Event in the allColumnsList

'Dim allColumnsList As Object
'Dim allColumnsList As New ArrayList
Set allColumnsList = CreateObject("System.Collections.ArrayList")
'allColumnsList
' Add items
allColumnsList.Add "2016/03/06" ' BT
allColumnsList.Add "2016/04/08"
allColumnsList.Add "2016/04/09"
allColumnsList.Add "2016/04/10"
allColumnsList.Add "2016/04/15" ' AT
allColumnsList.Add "2016/04/22"
allColumnsList.Add "2016/05/19"
allColumnsList.Add "2016/05/27"
allColumnsList.Add "2016/06/03"
allColumnsList.Add "2016/06/17"
allColumnsList.Add "2016/06/24"
allColumnsList.Add "2016/07/01"
allColumnsList.Add "2016/07/08"
allColumnsList.Add "2016/07/12"
allColumnsList.Add "2016/07/16"
allColumnsList.Add "2016/07/22"
allColumnsList.Add "2016/07/29"
allColumnsList.Add "2016/08/05"
allColumnsList.Add "2016/08/12"
allColumnsList.Add "2016/08/19"
allColumnsList.Add "2016/08/26"
allColumnsList.Add "2016/09/02"
allColumnsList.Add "2016/09/09"
allColumnsList.Add "2016/09/16"
allColumnsList.Add "2016/09/24"
allColumnsList.Add "2016/09/30"
allColumnsList.Add "2016/10/07"
allColumnsList.Add "2016/10/21"
allColumnsList.Add "2016/10/28"
allColumnsList.Add "2016/11/04"
allColumnsList.Add "2016/11/14"
allColumnsList.Add "2016/11/18"
allColumnsList.Add "2016/12/03"
allColumnsList.Add "2016/12/09"
allColumnsList.Add "2016/12/16"
allColumnsList.Add "2016/12/22"
allColumnsList.Add "2017/01/06"
allColumnsList.Add "2017/01/13"
allColumnsList.Add "2017/01/20"
allColumnsList.Add "2017/01/27"
allColumnsList.Add "2017/02/03"
allColumnsList.Add "2017/02/10"
allColumnsList.Add "2017/02/17"
allColumnsList.Add "2017/02/24"
allColumnsList.Add "2017/03/03"
allColumnsList.Add "2017/03/10"
allColumnsList.Add "2017/03/17"
allColumnsList.Add "2017/03/24"
allColumnsList.Add "2017/03/31"
allColumnsList.Add "2017/04/07"
allColumnsList.Add "2017/04/14"
allColumnsList.Add "2017/04/21"
allColumnsList.Add "2017/04/28"
allColumnsList.Add "2017/05/12"
allColumnsList.Add "2017/05/19"
allColumnsList.Add "2017/05/24"
allColumnsList.Add "2017/06/02"
allColumnsList.Add "2017/06/09"
allColumnsList.Add "2017/06/16"
allColumnsList.Add "2017/06/23"
allColumnsList.Add "2017/07/04"
allColumnsList.Add "2017/07/07"
allColumnsList.Add "2017/07/14"
allColumnsList.Add "2017/07/22" ' FCT
allColumnsList.Add "2017/07/28"
allColumnsList.Add "2017/08/03"
allColumnsList.Add "2017/08/09"
allColumnsList.Add "2017/08/11"
allColumnsList.Add "2017/08/19"
allColumnsList.Add "2017/08/25"
allColumnsList.Add "2017/09/02"
allColumnsList.Add "2017/09/08"
allColumnsList.Add "2017/09/15"
allColumnsList.Add "2017/09/26"
allColumnsList.Add "2017/09/29"
allColumnsList.Add "2017/10/06"
allColumnsList.Add "2017/10/24"
allColumnsList.Add "2017/11/14"
allColumnsList.Add "2017/12/18"
allColumnsList.Add "2018/01/25"
allColumnsList.Add "2018/04/04"
allColumnsList.Add "2018/04/12"
allColumnsList.Add "2018/05/04" ' OWLD
allColumnsList.Add "2019/04/01"
allColumnsList.Add "2019/09/05"
allColumnsList.Add "2020/01/22"
allColumnsList.Add "2020/04/03" ' Final

'Dim allTablesList As Object
'Dim allTablesList As New ArrayList
Set allTablesList = CreateObject("System.Collections.ArrayList")
'allTablesList
' Add items
allTablesList.Add "2016/03/06_LPD26_BT"
allTablesList.Add "2016/04/08_LPD26"
allTablesList.Add "2016/04/09_LPD26"
allTablesList.Add "2016/04/10_LPD26"
allTablesList.Add "2016/04/15_LPD26_AT"
allTablesList.Add "2016/04/22_LPD26"
allTablesList.Add "2016/05/19_LPD26"
allTablesList.Add "2016/05/27_LPD26"
allTablesList.Add "2016/06/03_LPD26"
allTablesList.Add "2016/06/17_LPD26"
allTablesList.Add "2016/06/24_LPD26"
allTablesList.Add "2016/07/01_LPD26"
allTablesList.Add "2016/07/08_LPD26"
allTablesList.Add "2016/07/12_LPD26"
allTablesList.Add "2016/07/16_LPD26"
allTablesList.Add "2016/07/22_LPD26"
allTablesList.Add "2016/07/29_LPD26"
allTablesList.Add "2016/08/05_LPD26"
allTablesList.Add "2016/08/12_LPD26"
allTablesList.Add "2016/08/19_LPD26"
allTablesList.Add "2016/08/26_LPD26"
allTablesList.Add "2016/09/02_LPD26"
allTablesList.Add "2016/09/09_LPD26"
allTablesList.Add "2016/09/16_LPD26"
allTablesList.Add "2016/09/24_LPD26"
allTablesList.Add "2016/09/30_LPD26"
allTablesList.Add "2016/10/07_LPD26"
allTablesList.Add "2016/10/21_LPD26"
allTablesList.Add "2016/10/28_LPD26"
allTablesList.Add "2016/11/04_LPD26"
allTablesList.Add "2016/11/14_LPD26"
allTablesList.Add "2016/11/18_LPD26"
allTablesList.Add "2016/12/03_LPD26"
allTablesList.Add "2016/12/09_LPD26"
allTablesList.Add "2016/12/16_LPD26"
allTablesList.Add "2016/12/22_LPD26"
allTablesList.Add "2017/01/06_LPD26"
allTablesList.Add "2017/01/13_LPD26"
allTablesList.Add "2017/01/20_LPD26"
allTablesList.Add "2017/01/27_LPD26"
allTablesList.Add "2017/02/03_LPD26"
allTablesList.Add "2017/02/10_LPD26"
allTablesList.Add "2017/02/17_LPD26"
allTablesList.Add "2017/02/24_LPD26"
allTablesList.Add "2017/03/03_LPD26"
allTablesList.Add "2017/03/10_LPD26"
allTablesList.Add "2017/03/17_LPD26"
allTablesList.Add "2017/03/24_LPD26"
allTablesList.Add "2017/03/31_LPD26"
allTablesList.Add "2017/04/07_LPD26"
allTablesList.Add "2017/04/14_LPD26"
allTablesList.Add "2017/04/21_LPD26"
allTablesList.Add "2017/04/28_LPD26"
allTablesList.Add "2017/05/12_LPD26"
allTablesList.Add "2017/05/19_LPD26"
allTablesList.Add "2017/05/24_LPD26"
allTablesList.Add "2017/06/02_LPD26"
allTablesList.Add "2017/06/09_LPD26"
allTablesList.Add "2017/06/16_LPD26"
allTablesList.Add "2017/06/23_LPD26"
allTablesList.Add "2017/07/04_LPD26"
allTablesList.Add "2017/07/07_LPD26"
allTablesList.Add "2017/07/14_LPD26"
allTablesList.Add "2017/07/22_LPD26_FCT"
allTablesList.Add "2017/07/28_LPD26"
allTablesList.Add "2017/08/03_LPD26"
allTablesList.Add "2017/08/09_LPD26"
allTablesList.Add "2017/08/11_LPD26"
allTablesList.Add "2017/08/19_LPD26"
allTablesList.Add "2017/08/25_LPD26"
allTablesList.Add "2017/09/02_LPD26"
allTablesList.Add "2017/09/08_LPD26"
allTablesList.Add "2017/09/15_LPD26"
allTablesList.Add "2017/09/26_LPD26"
allTablesList.Add "2017/09/29_LPD26"
allTablesList.Add "2017/10/06_LPD26"
allTablesList.Add "2017/10/24_LPD26"
allTablesList.Add "2017/11/14_LPD26"
allTablesList.Add "2017/12/18_LPD26"
allTablesList.Add "2018/01/25_LPD26"
allTablesList.Add "2018/04/04_LPD26"
allTablesList.Add "2018/04/12_LPD26"
allTablesList.Add "2018/05/04_LPD26_OWLD"
allTablesList.Add "2019/04/01_LPD26"
allTablesList.Add "2019/09/05_LPD26"
allTablesList.Add "2020/01/22_LPD26"
allTablesList.Add "2020/04/03_LPD26_Final"

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
