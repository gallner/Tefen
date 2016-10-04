Attribute VB_Name = "mdlErrorHandler"

' Description:  An Error handling module. This entire module was
'               downloaded from http://www.fmsinc.com/tpapers/vbacode/Debug.asp#AdvancesErrorHandling
'               The module was updated in order to fit the applications' needs.
'
'
' Authors:      Luke Chung
'
' Editor:       Nir Gallner, nir@verisoft.co
'
'
' Date                 Comment
' --------------------------------------------------------------
' 9/25/2016            Initial version
'
Option Explicit

'sets the application in an error collection mode
Public Const gcfHandleErrors As Boolean = True

' Current pointer to the array element of the call stack
Private mintStackPointer As Integer

' Array of procedure names in the call stack
Private mastrCallStack() As String

' The number of elements to increase the array
Private Const mcintIncrementStackSize As Integer = 10


''''''''''''''''''''''''''''''''''''''''''''''''
'   List of Custom Error Codes
'''''''''''''''''''''''''''''''''''''''''''''''
Public Enum CustomError
    STRING_IS_EMPTY = 21001
    CONNECTION_TO_DB_FAIL = 21002
    FAIL_TO_EXECUTE_SELECT_STATEMENT = 21003
    FAIL_TO_EXECUTE_INSERT_STATEMENT = 21004
    FAIL_TO_EXECUTE_UPDATE_STATEMENT = 21005
    FAIL_TO_EXECUTE_DELETE_STATEMENT = 21006
    SYSTEM_ID_IS_MISSING = 21007
    NO_DATA_FOUND_IN_DB = 21008
    MANDATORY_ID_IS_MISSING = 21009
    INVALID_FORM_INPUT = 21010
    INVALID_INPUT = 21011
End Enum
    
    




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   Sub PushCallStack(strProcName As String)
'
'   Comments: Add the current procedure name to the Call Stack.
'   Should be called whenever a procedure is called
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PushCallStack(strProcName As String)
  
61:  On Error Resume Next

  ' Verify the stack array can handle the current array element
64:  If mintStackPointer > UBound(mastrCallStack) Then
    ' If array has not been defined, initialize the error handler
66:    If Err.Number = 9 Then
67:      ErrorHandlerInit
68:    Else
      ' Increase the size of the array to not go out of bounds
70:      ReDim Preserve mastrCallStack(UBound(mastrCallStack) + mcintIncrementStackSize)
71:    End If
72:  End If

74:  On Error GoTo 0

76:  mastrCallStack(mintStackPointer) = strProcName

  ' Increment pointer to next element
79:  mintStackPointer = mintStackPointer + 1
80: End Sub

Private Sub ErrorHandlerInit()
83:  mintStackPointer = 1
84:  ReDim mastrCallStack(1 To mcintIncrementStackSize)
85: End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   Sub PopCallStack()
'
'   Comments: Remove a procedure name from the call stack
'   Should be called whenever a procedure Ends
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PopCallStack()

96:  If mintStackPointer <= UBound(mastrCallStack) Then
97:    mastrCallStack(mintStackPointer - 1) = ""
98:  End If

  ' Reset pointer to previous element
101:  mintStackPointer = mintStackPointer - 1
102: End Sub

Sub SampleErrorWithLineNumbers()

106: MsgBox "Error Line: " & Erl & vbCrLf & vbCrLf & _
             "Error: (" & Err.Number & ") " & Err.Description, vbCritical
108: End Sub
  ' Comments: Main procedure to handle errors that occur.
Sub GlobalErrHandler()

112:    Dim MyLog As Log
113:    Set MyLog = New Log: Call MyLog.InitClass

115:  Dim strError As String
116:  Dim lngError As Long
117:  Dim intErl As Integer
118:  Dim strMsg As String

  ' Variables to preserve error information
121:  strError = Err.Description
122:  lngError = Err.Number
123:  intErl = Erl

 
  
  ' Prompt the user with information on the error:
128:  strMsg = "Procedure: " & CurrentProcName() & vbCrLf & _
           "Line : " & intErl & vbCrLf & _
           "Error : (" & lngError & ")" & strError & vbCrLf & _
           "Stack Trace:" & Join(mastrCallStack)
132:  MyLog.Log (strMsg)
  
134:  Err.Clear

  
 
138: End Sub

Private Function CurrentProcName() As String
141:  CurrentProcName = mastrCallStack(mintStackPointer - 1)
142: End Function


Public Sub ThrowError(ErrorNumber As Long, ErrorInFunction As String, ErrorDescription As String)
    
147:    Err.Raise ErrorNumber, ErrorInFunction, ErrorDescription
    
149: End Sub




