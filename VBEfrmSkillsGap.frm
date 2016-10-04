VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSkillsGap 
   Caption         =   "מיפוי מערכות"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10200
   OleObjectBlob   =   "VBEfrmSkillsGap.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSkillsGap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Description:  This form is the interface for searching
'               for gaps is skills for a group of systems.
'               User inserts a list of systems and a list
'               of employees, and the application shows
'               the gap of skills needed to cover the selected
'               systems.
'
'
' Authors:      Nir Gallner, nir@verisoft.co
'
'
' Date                 Comment
' --------------------------------------------------------------
' 9/25/2016            Initial version
'
Option Explicit
Dim MyDb As LeumiDB

Private Sub ReloadCboSystems(status As Boolean)
     
22:    Dim i As Integer
23:    Dim itemToAdd As String
    
    'connect to DB
26:    Call MyDb.ConnectToDB
    
    'in case combo box is full- empty it first.
29:    If cboSystems.ListCount > 0 Then
30:        Do While cboSystems.ListCount > 0
31:            cboSystems.RemoveItem 0
32:        Loop
        
34:    End If
    
         
    'load system infromation
38:    MyDb.pSQLQuery = "SELECT SystemName FROM tblSystems inner join tblSystemDescription on " & _
        " tblSystems.systemid = tblsystemdescription.systemid where status=" & status & ";"
    
41:    MyDb.ExecuteSelect
    
43:    If IsArrayAllocated(MyDb.pDataArray) Then
44:        For i = 0 To UBound(MyDb.pDataArray, 2)
45:            itemToAdd = MyDb.pDataArray(0, i)
46:            cboSystems.AddItem itemToAdd
47:        Next
48:    End If
    
   

52: End Sub

Private Sub ReloadCboEmployees()
     
56:    Dim i As Integer
57:    Dim itemToAdd As String
    
    'connect to DB
60:    Call MyDb.ConnectToDB
    
    'in case combo box is full- empty it first.
63:    If cboEmployee.ListCount > 0 Then
64:        Do While cboEmployee.ListCount > 0
65:            cboEmployee.RemoveItem 0
66:        Loop
        
68:    End If
    
         
    'load system infromation
72:    MyDb.pSQLQuery = "SELECT FirstName + ' ' + LastName as 'Name' FROM tblEmployees;"
    
74:    MyDb.ExecuteSelect
    
76:    If IsArrayAllocated(MyDb.pDataArray) Then
77:        For i = 0 To UBound(MyDb.pDataArray, 2)
78:            itemToAdd = MyDb.pDataArray(0, i)
79:            cboEmployee.AddItem itemToAdd
80:        Next
81:    End If
    
   

85: End Sub


Private Sub cmdAddEmployeeToList_Click()
89:     lstEmployees.AddItem cboEmployee.value
90: End Sub

Private Sub cmdAddSysemToList_Click()
93:     Me.lstSystems.AddItem cboSystems.value
94: End Sub


Private Sub cmdCalculate_Click()

    'Error handling - identical to all subs and functions
100:    If gcfHandleErrors Then On Error GoTo cmdCalculate_Click_Error
101:    PushCallStack "frmSkillsGap.cmdCalculate_Click"
    
103:    Dim i As Integer
104:    Dim sSystems() As String
105:    Dim sEmployees() As String
106:    Dim sResult() As String
    
    'get the list of systems
109:    ReDim sSystems(lstSystems.ListCount - 1)
110:    For i = 0 To UBound(sSystems)
111:        sSystems(i) = lstSystems.List(i)
112:    Next i
    
    'get the list of systems
115:    ReDim sEmployees(lstEmployees.ListCount - 1)
116:    For i = 0 To UBound(sEmployees)
117:        sEmployees(i) = lstEmployees.List(i)
118:    Next i
    
120:    sResult = BuildSkillGapMap(sSystems, sEmployees)
     
122:    Call DisplayArray(sResult, sSystems, sEmployees)
123:    Me.Hide

'Exit Point
cmdCalculate_Click_Exit:
127:    PopCallStack
128:    On Error GoTo 0
129:    Exit Sub
     
'Error Handling
cmdCalculate_Click_Error:
133:        GlobalErrHandler
134:        Resume cmdCalculate_Click_Exit
135: End Sub

Private Sub cmdClearForm_Click()
138:    lstEmployees.Clear
139:    lstSystems.Clear
    
141: End Sub

Private Sub UserForm_Initialize()

145:    Set MyDb = New LeumiDB
146:    Call ReloadCboSystems(True)
147:    Call ReloadCboEmployees
148: End Sub
