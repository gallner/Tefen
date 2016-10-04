VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Description:  Logging events in the SystemLog sheet
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
Private Const iFatal = 1
Private Const iError = 2
Private Const iWarning = 3
Private Const iInfo = 4
Private Const iDebug_Info = 5

Private Const LogSheetName = "SystemLog"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       Public Sub initClass()
'
'   Sub serves as sort of Constructor for class
'   Currently not doing any thing, but exist for best practices
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitClass()

32: End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       Public Sub FinalizeClass()
'
'   Sub serves as sort of destructor for class
'   Currently not doing any thing, but exist for best practices
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FinalizeClass()

42: End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'     Public Sub Log(Msg As String, Optional SubName As String, Optional LogLevel As Variant)
'
'   Sub receives the following parameters, and printes them out to the log sheet:
'   1. Msg - the message to be logged
'   2. SubName- this is an optional parameter. The sub that in her watch a log was called
'   3. LogLevel - FATAL, ERROR, WARNING, INFO, DEBUG_INFO
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Log(msg As String, Optional SubName As String, Optional LogLevel As Variant)

    'Error Handling
54:    If gcfHandleErrors Then On Error GoTo Log_Error
55:    PushCallStack "Log.Log"
    
57:    Dim LastRow As Integer, data()
     
    'get the range of the end of the interfaces list
60:    LastRow = Sheets(LogSheetName).Cells(Sheets(LogSheetName).Rows.count, "A").End(xlUp).row + 1
    
62:    ReDim data(1 To 1, 1 To 4)
63:    data(1, 1) = Now()
64:    data(1, 2) = msg
65:    If Not IsMissing(SubName) Then data(1, 3) = SubName
66:    If Not IsMissing(LogLevel) Then data(1, 4) = GetLevelName(CInt(LogLevel))
    
68:    Dim MyRange As String
69:    MyRange = "A" & LastRow & ":D" & LastRow
70:    Sheets(LogSheetName).Range(MyRange) = data
    
'Exit Point
Log_Exit:
74:    PopCallStack
75:    On Error GoTo 0
76:    Exit Sub

Log_Error:
79:    GlobalErrHandler
80:    Resume Log_Exit
        
82: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Private Function GetLevelName(Level As Integer) As String
'
'   Function receives Level is as integer and returns a string representing the
'   Level name according to the id. If the level id is not within the
'   list of known level id's, function returns a string "UNDEFINED"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetLevelName(Level As Integer) As String
    
    'Error Handling
95:    If gcfHandleErrors Then On Error GoTo GetLevelName_Error
96:    PushCallStack "Log.GetLevelName"
    
98:    Dim result As String
    
100:    Select Case Level
    Case FATAL
102:        result = "FATAL"
103:    Case ERROR
104:        result = "ERROR"
105:    Case WARNING
106:        result = "WARNING"
107:    Case INFO
108:        result = "INFO"
109:    Case DEBUG_INFO
110:        result = "DEBUG_INFO"
111:    Case Else
112:        result = "UNDEFINED"
113:    End Select
    
115:    GetLevelName = result
    
'Exit Point
GetLevelName_Exit:
119:    PopCallStack
120:    On Error GoTo 0
121:    Exit Function

GetLevelName_Error:
124:    GlobalErrHandler
125:    GetLevelName = "UNDEFINED"
126:    Resume GetLevelName_Exit
     
128: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Start of Properties Section
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get FATAL() As Integer
136:    FATAL = iFatal
137: End Property

Public Property Get ERROR() As Integer
140:    ERROR = iError
141: End Property

Public Property Get WARNING() As Integer
144:    WARNING = iWarning
145: End Property

Public Property Get INFO() As Integer
148:    INFO = iInfo
149: End Property

Public Property Get DEBUG_INFO() As Integer
152:    DEBUG_INFO = iDebug_Info
153: End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   End of Properties Section
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''


