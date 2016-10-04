VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LeumiDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Description:  Handles all connection to the DB, and perform the
'               queries: select, insert, update, delete.
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
Private DataArray
Private Conn As ADODB.Connection
Private RS As ADODB.Recordset
Private sqlQuery As String
Private DbPath As String
Private ConnString As String
Private lngRecsAff As Long
Private MyLog As Log

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       Public Sub InitClass()
'
'   Sub serves as sort of costruction for class.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitClass()

34:    Set MyLog = New Log

36: End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       Public Sub finalizeClass()
'
'   Sub serves as sort of destruction for class.
'   It closes the connection to the DB, and sets objects to Nothing.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FinalizeClass()
46:   Call CloseConnection
47:   Set Conn = Nothing
48:   Set RS = Nothing
49:   Set MyLog = Nothing
50: End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       Public Function ConnectToDB()
'
'   Connects to the DB in order to query or insert information into the DB.
'   This function should be called numerous times since connection should be closed every time an insert
'   update, delete or query is done.
'   returns a boolean value- success or fails.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ConnectToDB() As Boolean
    
    'Error Handling
64:    If gcfHandleErrors Then On Error GoTo ConnectToDB_Error
65:    PushCallStack "LeumiDB.ConnectToDB"
    
    
    'connection established
69:    If Not Conn Is Nothing Then
70:        ConnectToDB = True
71:        GoTo ConnectToDB_Exit
72:    End If
    
    'get the path do DB
75:    Dim wksSheet1 As Worksheet
76:    Set wksSheet1 = Sheets("SystemLog")

78:    DbPath = wksSheet1.CustomProperties.Item(1).value
79:    ConnString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                                      "Data Source=" & DbPath & ";" & _
                                      "Jet OLEDB:Engine Type=5;" & _
                                      "Persist Security Info=False;"
                                      
    'new connection
85:    Set Conn = New ADODB.Connection
86:    Conn.Open ConnString
      
    'test database connection
    'if error has been raised during connection to the DB or creating the Conn and Recordset objects.
90:    ConnectToDB = True

'Exit Point
ConnectToDB_Exit:
94:    PopCallStack
95:    On Error GoTo 0
96:    Exit Function

ConnectToDB_Error:
99:    GlobalErrHandler
100:    ThrowError CustomError.CONNECTION_TO_DB_FAIL, "LeumiDB.ConnectToDB", "Error in connecting to DB. View Log for more information"
101:    ConnectToDB = False
102:    Resume ConnectToDB_Exit
    
104: End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       Public Sub CloseConnection()
'
'   Closes the current connection to the DB and sets the Conn variable to be nothing.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CloseConnection()
    
    'Error Handling
115:    If gcfHandleErrors Then On Error GoTo CloseConnection_Error
116:    PushCallStack "LeumiDB.CloseConnection"
    
    'only close the connection
119:    If Not Conn Is Nothing Then
120:        Conn.Close
121:        Set Conn = Nothing
122:    End If
    
    'if error has been raised
125:    If Err.Number Then
126:        Call MyLog.Log("Error in Closing connection to DB. Error: " & Err.Description, "LeumiDB.CloseConnection", MyLog.ERROR)
127:        Err.Clear
128:    End If
    
'Exit Point
CloseConnection_Exit:
132:    PopCallStack
133:    On Error GoTo 0
134:    Exit Sub

CloseConnection_Error:
137:    GlobalErrHandler
138:    Resume CloseConnection_Exit
    
140: End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       Public Sub ExecuteSelect()
'
'   Sub uses sqlQuery to run a select query in the DB.
'   Sub assumes sqlQuery is set correctly.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ExecuteSelect()
    
    'Error Handling
154:    If gcfHandleErrors Then On Error GoTo ExecuteSelect_Error
155:    PushCallStack "LeumiDB.ExecuteSelect"
    
     
158:    If RS Is Nothing Then Set RS = New ADODB.Recordset
159:    If Conn Is Nothing Then Call ConnectToDB
    
161:    Set DataArray = Nothing
162:    Dim tmpArr
    
164:    RS.Open sqlQuery, Conn, adOpenStatic, adLockBatchOptimistic, adCmdText
    
166:    If RS.RecordCount > 0 Then tmpArr = RS.GetRows
    
    
    
'Exit Point
ExecuteSelect_Exit:
172:    PopCallStack
173:    DataArray = tmpArr
174:    RS.Close: Set RS = Nothing
175:    On Error GoTo 0
176:    Exit Sub

ExecuteSelect_Error:
179:    GlobalErrHandler
180:    ThrowError CustomError.FAIL_TO_EXECUTE_SELECT_STATEMENT, "LeumiDB.ExecuteSelect", _
            "Error executing Select statement. View Log for further information"
182:    Resume ExecuteSelect_Exit
     
         
185: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       Public Sub ExecuteInsert()
'
'   Sub uses sqlQuery to run an insert query in the DB.
'   Sub assumes sqlQuery is set correctly.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ExecuteInsert()
    
    'Error Handling
198:    If gcfHandleErrors Then On Error GoTo ExecuteInsert_Error
199:    PushCallStack "LeumiDB.ExecuteInsert"
    
    'db object locked by another function or sub
     
203:    If Conn Is Nothing Then Call ConnectToDB
        
205:    Conn.Execute sqlQuery, lngRecsAff, adExecuteNoRecords
    
'Exit Point
ExecuteInsert_Exit:
209:    PopCallStack
210:    On Error GoTo 0
211:    Exit Sub

ExecuteInsert_Error:
214:    GlobalErrHandler
215:    ThrowError CustomError.FAIL_TO_EXECUTE_INSERT_STATEMENT, "LeumiDB.ExecuteInsert", _
            "Error in inserting data. View log for more information"
217:    Resume ExecuteInsert_Exit
        
219: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       Public Sub ExecuteUpdate()
'
'   Sub uses sqlQuery to run an Update query in the DB.
'   Sub assumes sqlQuery is set correctly.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ExecuteUpdate()
    
    'Error Handling
232:    If gcfHandleErrors Then On Error GoTo ExecuteUpdate_Error
233:    PushCallStack "LeumiDB.ExecuteUpdate"
    
    'db object locked by another function or sub
     
237:    If Conn Is Nothing Then Call ConnectToDB
238:    Conn.Execute sqlQuery, lngRecsAff, adExecuteNoRecords
    

'Exit Point
ExecuteUpdate_Exit:
243:    PopCallStack
244:    On Error GoTo 0
245:    Exit Sub

ExecuteUpdate_Error:
248:    GlobalErrHandler
249:    ThrowError CustomError.FAIL_TO_EXECUTE_UPDATE_STATEMENT, "LeumiDB.ExecuteUpdate", _
                "Error in updating data. View log for more information"
251:    Resume ExecuteUpdate_Exit
    
253: End Sub

  
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       Public Sub ExecuteDelete()
'
'   Sub uses sqlQuery to run a delete query in the DB.
'   Sub assumes sqlQuery is set correctly.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ExecuteDelete()
    
    'Error Handling
266:    If gcfHandleErrors Then On Error GoTo ExecuteDelete_Error
267:    PushCallStack "LeumiDB.ExecuteDelete"
    
269:     If Conn Is Nothing Then Call ConnectToDB
270:    Conn.Execute sqlQuery, lngRecsAff, adExecuteNoRecords
    
'Exit Point
ExecuteDelete_Exit:
274:    PopCallStack
275:    On Error GoTo 0
276:    Exit Sub

ExecuteDelete_Error:
279:    GlobalErrHandler
280:    ThrowError CustomError.FAIL_TO_EXECUTE_DELETE_STATEMENT, "LeumiDB.ExecuteDelete", _
                "Error in deleting data. View log for more infromation"
282:    Resume ExecuteDelete_Exit
    
        
285: End Sub
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Start of Properties Section
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get pDataArray()
292:    pDataArray = DataArray
293: End Property

Public Property Get pConn() As ADODB.Connection
296:    pConn = Conn
297: End Property

Public Property Get pRS() As ADODB.Recordset
'    pRS = RS
'TODO - fix this
302: End Property
Public Property Get pDbPath() As String
304:    pDbPath = DbPath
305: End Property

Public Property Get pSQLQuery() As String
308:    pSQLQuery = sqlQuery
309: End Property
Public Property Let pSQLQuery(qry As String)
311:    If qry <> "" Then sqlQuery = qry
312: End Property
Public Property Get pConnectionString() As String
314:    pConnectionString = ConnString
315: End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   End of Properties Section
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''



