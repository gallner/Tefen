VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSystems 
   Caption         =   "מיפוי מערכות"
   ClientHeight    =   2460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5280
   OleObjectBlob   =   "VBEfrmSystems.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSystems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
' Description:  This form allows users to choose
'               a system from the list of systems.
'               The application uses this for the following uses:
'               1. Search for a system
'               2. Search for an archived system
'               3. Delete a system
'
'
' Authors:      Nir Gallner, nir@verisoft.co
'
'
' Date                 Comment
' --------------------------------------------------------------
' 6/01/2016            Version 1.0
' 9/25/2016            Version 2.0
'
Option Explicit
Dim MyDb As LeumiDB
Dim MyLog As Log
Dim bActiveSystems As Boolean
   


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   Private Sub brnViewSystem_Click()
'
'   Sub looks for a specific system in the DB and loads it's information
'   into the Excel Form.
'   Sub makes sure the system is in the Combo Box before starting.
'   If the system is already loaded- sub moves focus to the system.
'   Else- sub uses Sub LoadSystemToForm with the system name
'   from the Combo Box to load system information.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub brnViewSystem_Click()
       
    'Error Handling
40:    If gcfHandleErrors Then On Error GoTo brnViewSystem_Error
41:    PushCallStack "frmSystems.brnViewSystem"

43:    Call PerfromanceEnabled(True)
    
45:    Dim MySystem As System
46:    Set MySystem = New System
47:    Call MySystem.InitClass
    
49:    Dim strName As String
     
    'get the system name
52:    strName = cboSystems.value
     
    'create a new sheet
55:    Application.ScreenUpdating = False
56:    With Sheets("Template")
57:        .Visible = True
58:        .Copy After:=Sheets(Sheets.count)
59:        .Visible = False
60:    End With
    
62:    ActiveSheet.Name = strName
63:    Application.ScreenUpdating = True

    
    'search for system in DB
67:    Dim SystemId As Integer
68:    SystemId = MySystem.FindSystemId(Trim(strName))
    
    'found system in DB
71:    If SystemId > 0 Then
72:        MySystem.SysId = SystemId
73:        Call MySystem.FromDbToClass
74:        Call FromClassToForm(MySystem)
75:    Else
76:        MsgBox "אין אפשרות להציג את המערכת המבוקשת", vbInformation, "צפיה בנתוני מערכת"
77:    End If
   
79:    frmSystems.Hide
    
'Exit Point
brnViewSystem_Exit:
83:    PopCallStack
84:    Call PerfromanceEnabled(False)
85:    On Error GoTo 0
86:    Exit Sub

brnViewSystem_Error:
89:    GlobalErrHandler
90:    MsgBox "אירעה שגיאה, אנא פנה למנהל מערכת", vbCritical, "שגיאה"
91:    Resume brnViewSystem_Exit

93: End Sub
     '''OK
Private Sub ReloadCboSystems(status As Boolean)
     
    'Error Handling
98:    If gcfHandleErrors Then On Error GoTo ReloadCboSystems_Error
99:    PushCallStack "frmSystems.ReloadCboSystems"
    
101:    Dim i As Integer, LastRow As Integer
102:    Dim itemToAdd As String
    
    'connect to DB
105:    Call MyDb.ConnectToDB
    
    'in case combo box is full- empty it first.
108:    If cboSystems.ListCount > 0 Then
109:        Do While cboSystems.ListCount > 0
110:            cboSystems.RemoveItem 0
111:        Loop
        
113:    End If
    
         
    'load system infromation
117:    MyDb.pSQLQuery = "SELECT tblSystems.SystemName FROM tblSystems " & _
        " inner join tblSystemDescription on tblSystems.systemid = tblSystemDescription.systemid where status=" & status & ";"
    
120:    MyDb.ExecuteSelect
    
122:    If IsArrayAllocated(MyDb.pDataArray) Then
123:        For i = 0 To UBound(MyDb.pDataArray, 2)
124:            itemToAdd = MyDb.pDataArray(0, i)
125:            cboSystems.AddItem itemToAdd
126:        Next
127:    End If
    
   
'Exit Point
ReloadCboSystems_Exit:
132:    PopCallStack
133:    On Error GoTo 0
134:    Exit Sub

ReloadCboSystems_Error:
137:    GlobalErrHandler
138:    Resume ReloadCboSystems_Exit
    
140: End Sub


Private Sub UserForm_Activate()
144:    Call ReloadCboSystems(Me.isActiveSystem)
145: End Sub

'''''OK
Private Sub UserForm_Initialize()
   
150:    Set MyDb = New LeumiDB
151:    Set MyLog = New Log
    
153: End Sub

Property Let isActiveSystem(active As Boolean)
156:    bActiveSystems = active
157: End Property

Property Get isActiveSystem() As Boolean
160:    isActiveSystem = bActiveSystems
161: End Property


