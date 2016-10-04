Attribute VB_Name = "Debug_Lines"
'
' Description:  This module's functionality is to either add
'               or remove debug lines (10: XXXX) in front of
'               statementst. It was downloaded from an unknown
'               source on the web and tweaked to fit the project's
'               needs
'
'
' Authors:      Unknown
'
' Editor:       Nir Gallner, nir@verisoft.co
'
' Date                 Comment
' --------------------------------------------------------------
' 9/25/2016            Initial version
'
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'           Private Sub callAddLines()
'
'   Sub goes over all the files in the VBA project and adds lines in the right places.
'   Currently, every new file needs to be updated manually into the sub
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub callAddLines()

    'forms
29:    Call AddLineNumbers("frmFindEmpBySkill")
30:    Call AddLineNumbers("frmSkillsGap")
31:    Call AddLineNumbers("frmSystems")
    
    'modules
34:    Call AddLineNumbers("Debug_Lines")
35:    Call AddLineNumbers("mdlEnums")
36:    Call AddLineNumbers("mdlErrorHandler")
37:    Call AddLineNumbers("mdlFindEmpBySkill")
38:    Call AddLineNumbers("mdlPresentation")
39:    Call AddLineNumbers("mdlRibbonHandler")
40:    Call AddLineNumbers("mdlSkillsGap")
41:    Call AddLineNumbers("mdlUtils")
    
    'classes
44:    Call AddLineNumbers("Employee")
45:    Call AddLineNumbers("Interface")
46:    Call AddLineNumbers("LeumiDB")
47:    Call AddLineNumbers("Log")
48:    Call AddLineNumbers("Skill")
49:    Call AddLineNumbers("System")
    

52: End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'           Private Sub callRemoveLines()
'
'   Sub goes over all the files in the VBA project and removes lines from the right places.
'   Currently, every new file needs to be updated manually into the sub
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub callRemoveLines()

    'forms
64:    Call RemoveLineNumbers("frmFindEmpBySkill")
65:    Call RemoveLineNumbers("frmSkillsGap")
66:    Call RemoveLineNumbers("frmSystems")
    
    'modules
69:    Call RemoveLineNumbers("Debug_Lines")
70:    Call RemoveLineNumbers("mdlEnums")
71:    Call RemoveLineNumbers("mdlErrorHandler")
72:    Call RemoveLineNumbers("mdlFindEmpBySkill")
73:    Call RemoveLineNumbers("mdlPresentation")
74:    Call RemoveLineNumbers("mdlRibbonHandler")
75:    Call RemoveLineNumbers("mdlSkillsGap")
76:    Call RemoveLineNumbers("mdlUtils")
    
    'classes
79:    Call RemoveLineNumbers("Employee")
80:    Call RemoveLineNumbers("Interface")
81:    Call RemoveLineNumbers("LeumiDB")
82:    Call RemoveLineNumbers("Log")
83:    Call RemoveLineNumbers("Skill")
84:    Call RemoveLineNumbers("System")
    
86: End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'           Private Sub AddLineNumbers(vbCompName As String, Optional wbName As String)
'
'   This is the main sub of the module. It creceives a name of the file, and goes through
'   all the source lines, adding line numbers in the right places. Most of the logic for this
'   sub was downloaded from an unknown source in the web (some where in stackoverflow.com),
'   and the rest was tweaked
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddLineNumbers(vbCompName As String, Optional wbName As String)
    
100:    If wbName = "" Then wbName = ActiveWorkbook.Name
    'See MakeUF
102:    Dim i As Long, j As Long, lineN As Long
103:    Dim ProcName As String, ProcType As vbext_ProcKind
104:    Dim startOfProceedure As Long
105:    Dim lengthOfProceedure As Long
106:    Dim newLine As String
107:    Dim SelectFlag As Boolean
    
109:    SelectFlag = False
110:    With Workbooks(wbName).VBProject.VBComponents(vbCompName).CodeModule
111:        .CodePane.Window.Visible = False

113:        For i = 1 To .CountOfLines
114:            ProcType = GetProcType(wbName)
            
116:            ProcName = .ProcOfLine(i, ProcType)
117:            If ProcName <> vbNullString Then
118:                startOfProceedure = .ProcStartLine(ProcName, ProcType)
119:                lengthOfProceedure = .ProcCountLines(ProcName, ProcType)

121:                If startOfProceedure < i And i < startOfProceedure + lengthOfProceedure Then
122:                    newLine = RemoveOneLineNumber(.Lines(i, 1))
123:                    If Not HasLabel(newLine) And Not (.Lines(i - 1, 1) Like "* _") And Not IsComment(newLine) And Not IsFunction(newLine) And Not IsEmptyLine(newLine) Then
                
125:                        If SelectFlag = False Then .ReplaceLine i, CStr(i) & ":" & newLine
126:                        If InStr(1, newLine, "Select Case") Then
                            SelectFlag = True
128:                        Else
129:                            SelectFlag = False
130:                        End If
131:                    End If
132:                End If
133:            End If

135:        Next i
136:        .CodePane.Window.Visible = True
137:    End With
138: End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Private Sub RemoveLineNumbers(vbCompName As String, Optional wbName As String)
'
'   Sub receives a file name and goes through the file, removing the line numbers
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveLineNumbers(vbCompName As String, Optional wbName As String)
    
148:    If wbName = "" Then wbName = ActiveWorkbook.Name
    'See MakeUF
150:    Dim i As Long
151:    With Workbooks(wbName).VBProject.VBComponents(vbCompName).CodeModule
152:        For i = 1 To .CountOfLines
153:            .ReplaceLine i, RemoveOneLineNumber(.Lines(i, 1))
154:        Next i
155:    End With
156: End Sub

Private Function RemoveOneLineNumber(aString)
159:    RemoveOneLineNumber = aString
160:    If aString Like "#:*" Or aString Like "#*#:*" Then
161:        RemoveOneLineNumber = Mid(aString, 1 + InStr(1, aString, ":", vbTextCompare))
162:    End If
163: End Function

Private Function HasLabel(ByVal aString As String) As Boolean
166:    HasLabel = InStr(1, aString & ":", ":") < InStr(1, aString & " ", " ")
167: End Function

Private Function IsComment(ByVal aString As String) As Boolean
170:    IsComment = Left(Trim(aString), 1) = "'"
171: End Function
Private Function IsFunction(ByVal aString As String) As Boolean
173:    Dim StrArr As Variant, result As Boolean
    
175:    StrArr = Split(Trim(aString), " ")
    
     
178:    If UBound(StrArr) = -1 Then
179:        IsFunction = False
180:        Exit Function
181:    End If
    
183:    result = True
184:    Select Case StrArr(0)
        Case "Public"
186:        Case "Private"
187:        Case "Function"
188:        Case "Sub"
189:        Case "Property"
190:        Case Else
191:            result = False
192:    End Select
    
194:    IsFunction = result
195: End Function

Private Function IsEmptyLine(ByVal aString As String) As Boolean
198:    Dim StrArr As Variant
    
200:    StrArr = Split(Trim(aString), " ")
    
     
203:    If UBound(StrArr) = -1 Then
204:        IsEmptyLine = True
205:    Else
206:        IsEmptyLine = False
207:    End If
    
209: End Function

Private Function GetProcType(wbName As String, Optional vbCompName As Integer) As vbext_ProcKind
    
213:    On Error Resume Next
214:    Dim ProcName As String
215:    Dim i As Long 'TODO - figure out why vbCompName and i can be used empty

217:    With Workbooks(wbName).VBProject.VBComponents(vbCompName).CodeModule
        
219:        If Not IsError(.ProcOfLine(i, vbext_pk_Proc)) Then
220:             GetProcType = vbext_pk_Proc
221:        ElseIf Not IsError(.ProcOfLine(i, vbext_pk_Get)) Then
222:            GetProcType = vbext_pk_Get
223:        ElseIf Not IsError(.ProcOfLine(i, vbext_pk_Let)) Then
224:            GetProcType = vbext_pk_Let
225:        ElseIf Not IsError(.ProcOfLine(i, vbext_pk_Set)) Then
226:            GetProcType = vbext_pk_Set
227:        Else
228:            GetProcType = vbNullString
229:        End If
230:    End With
    
232: End Function
