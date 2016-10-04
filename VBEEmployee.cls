VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Description:  This class is a container for employee informaion
'               It is an infrastructure class.
'
' Authors:      Nir Gallner, nir@verisoft.co
'
'
' Date                 Comment
' --------------------------------------------------------------
' 9/25/2016            Initial version
'
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                               Class Memebers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private sFirstName As String
Private sLastName As String
Private iEmpId As Integer
Private iPosition As Integer
Private MyLog As Log

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Sub serves as sort of constructor for the class.
'           It initializes relevant parmeters:
'           1. Strings will be assigned value ""
'           2. Integers are set to -1
'           3. Boolean values are set to False
'           4. Arrays are set to 2 spots (redim (1))- if applicable.
'
' Arguments:    None
'
' Returns:      None
'
' Date          Developer         Action
' --------------------------------------------------------------
' 09/25/16      Nir Gallner       Initial version
'
Public Sub InitClass()
40:    sFirstName = ""
41:    sLastName = ""
42:    iEmpId = -1
43:    iPosition = -1

45:    Set MyLog = New Log
46: End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Sub serves as sort of destructor for the class.
'           It frees up memory of the Log object
'
' Arguments:    None
'
' Returns:      None
'
' Date          Developer         Action
' --------------------------------------------------------------
' 09/25/16      Nir Gallner       Initial version
'
Public Sub FinalizeClass()
63:    Set MyLog = Nothing
64: End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Public Sub PopulateClass(Name As String, id As Integer)
'
'   function receives Employee Name (FirstName + Last Name) and Employee Id
'   as parameter, and populates the object with the data.
'   Sub splits the Name into First Name and Last Name, and calls FindGetEmployeeId
'   to get the ID.
'   Sub assumes that at least the name is present. If the name is not passed as
'   parameter, Sub exits.
'   Sub will look for employee id. If present, sub will populate the class
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateClass(Name As String, Optional id As Integer, Optional role As Integer)

    'Error Handling
82:    If gcfHandleErrors Then On Error GoTo PopulateEmployeeClass_Error
83:    PushCallStack "Employee.PopulateEmployeeClass"
    
85:    Dim NameArray() As String, i As Integer

    'values came in blank
88:    If Name = "" Then
89:        ThrowError CustomError.STRING_IS_EMPTY, "Employee.PopulateEmployeeClass", _
                    "Cannot populate employee class with an empty string"
91:    End If

    'if class is populated, delete current data
94:    If Not Me.IsClassEmpty Then Call Me.InitClass

96:    NameArray = Split(Name, " ")

    'fill in First Name and Last Name
99:    Me.FirstName = Trim(NameArray(0))

101:    For i = 1 To UBound(NameArray)
102:         Me.LastName = Me.LastName & Trim(NameArray(i)) & " "
103:    Next

105:    Me.LastName = Trim(Me.LastName)

107:    If id > 0 Then Me.EmpId = id
108:    If role > 0 Then Me.Position = role
    
    'there is no id, so look in the DB
111:    If id <= 0 Then
112:        Dim Query As String, MyDb As LeumiDB
113:        Set MyDb = New LeumiDB: MyDb.InitClass
        
115:        Query = "select tblEmployees.EmployeeId from tblEmployees where firstname ='" & _
        Me.FirstName & "' and lastname = '" & Me.LastName & "';"
        
118:        MyDb.pSQLQuery = Query
119:        MyDb.ExecuteSelect
        
121:        If CInt(MyDb.pDataArray(0, 0)) > 0 Then Me.EmpId = CInt(MyDb.pDataArray(0, 0))
122:    End If

'Exit Point
PopulateEmployeeClass_Exit:
126:    PopCallStack
127:    Set MyDb = Nothing
128:    On Error GoTo 0
129:    Exit Sub

PopulateEmployeeClass_Error:
132:    GlobalErrHandler
133:    Resume PopulateEmployeeClass_Exit
    
135: End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   Public Function IsClassEmpty()
'
'   Function checks if the class is empty:
'   1. EmpId = -1
'   2. FirstName= ""
'   3. LastName = ""
'
'   Function returns a boolean number representing whether the class is empty or not.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsClassEmpty() As Boolean

    'Error Handling
151:    If gcfHandleErrors Then On Error GoTo IsClassEmpty_Error
152:    PushCallStack "Employee.IsClassEmpty"
    
154:    Dim result As Boolean

156:    If Me.EmpId = -1 And Me.FirstName = "" And Me.LastName = "" Then
157:        result = True
158:    Else
159:       result = False
160:    End If

162:    IsClassEmpty = result

'Exit Point
IsClassEmpty_Exit:
166:    PopCallStack
167:    On Error GoTo 0
168:    Exit Function

IsClassEmpty_Error:
171:    GlobalErrHandler
172:    Resume IsClassEmpty_Exit
    
174: End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                               Start of Properties Section
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get FirstName() As String
181:    FirstName = sFirstName
182: End Property


Public Property Let FirstName(FName As String)
186:    sFirstName = FName
187: End Property

Public Property Get LastName() As String
190:    LastName = sLastName
191: End Property


Public Property Let LastName(LName As String)
195:    sLastName = LName
196: End Property


Public Property Get EmpId() As Integer
200:    EmpId = iEmpId
201: End Property


Public Property Let EmpId(id As Integer)
205:    iEmpId = id
206: End Property

Public Property Get Log()
209:    Log = MyLog
210: End Property
 
 Public Property Let Position(id As Integer)
213:    iPosition = id
214: End Property
 
 Public Property Get Position() As Integer
217:    Position = iPosition
218: End Property
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                               End of Properties Section
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


