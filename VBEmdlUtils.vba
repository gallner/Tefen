Attribute VB_Name = "mdlUtils"
Option Explicit




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               Function PrepareStrToDB(str As String)
'
'   Function receives a string and removes illegal values.
'   Values which will not pass well into the DB.
'   Currently the only char that is removed is the - '
'   char, which caused the SQL query to break if present
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function PrepareStrToDB(str As String)
    
    'Error Handling
17:    If gcfHandleErrors Then On Error GoTo PrepareStrToDB_Error
18:    PushCallStack "mdlUtils.PrepareStrToDB"
    
20:    Dim SpecialCharacters As String
21:    SpecialCharacters = "'" ' & Chr(10)  'modify as needed"
22:    Dim myString As String
23:    Dim newString As String
24:    Dim char As Variant
25:    newString = str
26:    For Each char In Split(SpecialCharacters, ",")
27:        newString = Replace(newString, char, " ")
28:    Next
29:    PrepareStrToDB = Trim(newString)

'Exit Point
PrepareStrToDB_Exit:
33:    PopCallStack
34:    On Error GoTo 0
35:    Exit Function

PrepareStrToDB_Error:
38:    GlobalErrHandler
39:    Resume PrepareStrToDB_Exit

41: End Function

Function ConvertNullToEmptyString(str)
44:    If IsNull(str) Then
45:        ConvertNullToEmptyString = ""
46:    Else
47:        ConvertNullToEmptyString = str
48:    End If

50: End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               Public Function YesNoPartial(str As String) As Integer
'
'   Function YesNoPartial: translate words from the form into integer.
'   Function receives as parameter a string: str, which is a string value to evaluate
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function YesNoPartial(str As String) As Integer

    'Error Handling
63:    If gcfHandleErrors Then On Error GoTo YesNoPartial_Error
64:    PushCallStack "mdlUtils.YesNoPartial"
    
66:    If Len(str) > 4 Then
67:        YesNoPartial = YesNoPartial(Left(str, 2))
68:        Exit Function
69:    End If

71:    Select Case str
        Case "קיים"
73:        YesNoPartial = 1
74:    Case "קי"
75:        YesNoPartial = 1
76:    Case "עובר"
77:        YesNoPartial = 1
78:    Case "0"
79:        YesNoPartial = 1
80:    Case "1"
81:        YesNoPartial = 1
82:    Case "כן"
83:        YesNoPartial = 1
84:    Case "לא"
85:        YesNoPartial = 0
86:    Case "חלקי"
87:        YesNoPartial = 2
88:    Case "חלקית"
89:        YesNoPartial = 2
90:    Case "חל"
91:        YesNoPartial = 2
92:    Case Else
93:        ThrowError CustomError.INVALID_INPUT, "mdlUtils.YesNoPartial", "cannot parse input as Yes/No/Partial input"
94:    End Select

'Exit Point
YesNoPartial_Exit:
98:    PopCallStack
99:    On Error GoTo 0
100:    Exit Function

YesNoPartial_Error:
103:    GlobalErrHandler
104:    Resume YesNoPartial_Exit

106: End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               Private Function CStrToBool(str As String) As Boolean
'
'   Function CStrToBool: translate words of boolean meaning into true or false.
'   Function receives as parameter a string: str, which is a string value to evaluate YesNoPartial
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CStrToBool(str As String) As Boolean

    'Error Handling
119:    If gcfHandleErrors Then On Error GoTo CStrToBool_Error
120:    PushCallStack "mdlUtils.CStrToBool"


123:    If Len(str) > 4 Then
124:        CStrToBool = CStrToBool(Left(str, 2))
125:        Exit Function
126:    End If

128:    Select Case str
        Case "קיים"
130:        CStrToBool = True
131:    Case "קי"
132:        CStrToBool = True
133:    Case "עובר"
134:        CStrToBool = True
135:    Case "0"
136:        CStrToBool = False
137:    Case "1"
138:        CStrToBool = True
139:    Case "כן"
140:        CStrToBool = True
141:    Case "לא"
142:        CStrToBool = False
143:    Case Else
144:        ThrowError CustomError.INVALID_INPUT, "mdlUtils.CStrToBool", "cannot parse function input into TRUE or FALSE"
145:    End Select

'Exit Point
CStrToBool_Exit:
149:    PopCallStack
150:    On Error GoTo 0
151:    Exit Function

CStrToBool_Error:
154:    GlobalErrHandler
155:    Resume CStrToBool_Exit

157: End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               Public Function IsArrayAllocated(Arr As Variant) As Boolean
'
' Returns TRUE if the array is allocated (either a static array or a dynamic array that has been
' sized with Redim) or FALSE if the array is not allocated (a dynamic that has not yet
' been sized with Redim, or a dynamic array that has been Erased). Static arrays are always
' allocated.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'
' This function is just the reverse of IsArrayEmpty.
' downloaded from : http://www.cpearson.com/excel/vbaarrays.htm
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsArrayAllocated(Arr As Variant) As Boolean

178: Dim N As Long
179: On Error Resume Next

' if Arr is not an array, return FALSE and get out.
182: If IsArray(Arr) = False Then
183:    IsArrayAllocated = False
184:    Exit Function
185: End If

' Attempt to get the UBound of the array. If the array has not been allocated,
' an error will occur. Test Err.Number to see if an error occurred.
189: N = UBound(Arr, 1)
190: If (Err.Number = 0) Then
    ''''''''''''''''''''''''''''''''''''''
    ' Under some circumstances, if an array
    ' is not allocated, Err.Number will be
    ' 0. To acccomodate this case, we test
    ' whether LBound <= Ubound. If this
    ' is True, the array is allocated. Otherwise,
    ' the array is not allocated.
    '''''''''''''''''''''''''''''''''''''''
199:    If LBound(Arr) <= UBound(Arr) Then
        ' no error. array has been allocated.
201:        IsArrayAllocated = True
202:    Else
203:        IsArrayAllocated = False
204:    End If
205: Else
    ' error. unallocated array
207:    IsArrayAllocated = False
208:    Err.Clear
209: End If

'Error Handling
IsArrayAllocated_Error:

214: End Function

 
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'             Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
'
'   Function receives a name of a sheet as a String and check whether the sheet exists.
'   Function returns a boolean - if sheet exits = true.
'
'   downloaded from: http://stackoverflow.com/questions/6688131/test-or-check-if-sheet-exists
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function SheetExists(Optional shtId As Variant, Optional wb As Workbook) As Boolean
226:    Dim sht As Worksheet

228:    If wb Is Nothing Then Set wb = ThisWorkbook

230:    On Error Resume Next

232:    Set sht = wb.Sheets(shtId)

234:    On Error GoTo 0
235:    Err.Clear

237:    SheetExists = Not sht Is Nothing
238: End Function
 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'             Private Sub cleanLists1()
'
'   Sub cleans all the general values
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cleanLists1()
    
    'Error Handling
249:    If gcfHandleErrors Then On Error GoTo cleanLists1_Error
250:    PushCallStack "mdlUtils.cleanLists1"
    
252:    Sheets("רשימות").Range("A3:A900").value = ""
253:    Sheets("רשימות").Range("C3:C900").value = ""
254:    Sheets("רשימות").Range("E3:E900").value = ""
255:    Sheets("רשימות").Range("G3:GA900").value = ""
256:    Sheets("רשימות").Range("I3:I900").value = ""
257:    Sheets("רשימות").Range("M3:N900").value = ""
    
'Exit Point
cleanLists1_Exit:
261:    PopCallStack
262:    On Error GoTo 0
263:    Exit Sub

cleanLists1_Error:
266:    GlobalErrHandler
267:    Resume cleanLists1_Exit
268: End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'           Public Sub PerfromanceEnabled(SetToActive As Boolean)
'
'   Sub turns on / off functionality in Excel in order to improve
'   performance. This sub is called prior to heavy sub using the
'   Excel GUI.
'   Sub receives a boolean:
'   True: Turn off functionality for best performance
'   False: Turn on functionality
'
'   For further information, see: https://blogs.office.com/2009/03/12/excel-vba-performance-coding-best-practices/
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PerfromanceEnabled(SetToActive As Boolean)

    'Error Handling
286:    If gcfHandleErrors Then On Error GoTo PerfromanceEnabled_Error
287:    PushCallStack "mdlUtils.PerfromanceEnabled"
    
289:    If SetToActive Then

' cannot use the next 2 lines because status bar is needed when loading system maps.
'301:        Application.ScreenUpdating = False
'302:        Application.DisplayStatusBar = False

295:        Application.Calculation = xlCalculationManual
296:        Application.EnableEvents = False
297:    Else
298:        Application.ScreenUpdating = True
299:        Application.DisplayStatusBar = True
300:        Application.Calculation = xlCalculationAutomatic
301:        Application.EnableEvents = True
302:    End If

'Exit Point
PerfromanceEnabled_Exit:
306:    PopCallStack
307:    On Error GoTo 0
308:    Exit Sub

PerfromanceEnabled_Error:
311:    GlobalErrHandler
312:    Resume PerfromanceEnabled_Exit
    
314: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   Sub BuildMap(MapType As String)
'
'   Sub builds a sheet with system general information.
'   It can be either system map of archive map,
'   according to MapType parameter received in the sub
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub BuildMap(MapType As String)

    'error handling
328:    If gcfHandleErrors Then On Error GoTo BuildMap_Error
329:    PushCallStack "mdlAdminFunctions.BuildMap"
    
331:    Dim MyDb As LeumiDB
332:    Set MyDb = New LeumiDB
    
334:    Dim Query As String
    
336:    Call MyDb.ConnectToDB
    
338:    If MapType = "System" Then
339:        MyDb.pSQLQuery = "select tblsystems.systemname from tblsystems inner join tblSystemDescription on " & _
                "tblsystems.systemid = tblSystemDescription.systemid where status=true;"
341:    Else
342:        MyDb.pSQLQuery = "select tblsystems.systemname from tblsystems inner join tblSystemDescription on " & _
                "tblsystems.systemid = tblSystemDescription.systemid where status=false;"
344:    End If
    
    
347:    MyDb.ExecuteSelect
    
    
350:    Dim i As Integer
     
352:    If IsArrayAllocated(MyDb.pDataArray) Then
            
            'should be True already. Just in case...
355:            Application.DisplayStatusBar = True
             
357:        For i = 0 To UBound(MyDb.pDataArray, 2)
358:            Call FromClassToMapSheet(MapType, CStr(MyDb.pDataArray(0, i)))
                
                'update the status bar
361:                Application.StatusBar = "Loading System " & i & " of " & UBound(MyDb.pDataArray, 2) & _
                "... " & Format(i / UBound(MyDb.pDataArray, 2), "0%") & " Completed"
363:        Next
    
365:    End If
    
    

        'close the starus bar
370:        Application.StatusBar = False
        
372:    Call MyDb.FinalizeClass: Set MyDb = Nothing
    
'Exit Point
BuildMap_Exit:
376:    PopCallStack
377:    On Error GoTo 0
378:    Exit Sub

BuildMap_Error:
381:    GlobalErrHandler
382:    Resume BuildMap_Exit
    
384: End Sub


''''''''''''''''''''''''''''''''''''''''
'           Private Sub BuildSystemMap()
'
'   Build system map using BuildMap sub
'
''''''''''''''''''''''''''''''''''''''''
Private Sub BuildSystemMap()
    
    'Error Handling
396:    If gcfHandleErrors Then On Error GoTo BuildSystemMap_Error
397:    PushCallStack "mdlAdminFunctions.BuildSystemMap"

399:    Call BuildMap("System")
    
    'Exit Point
BuildSystemMap_Exit:
403:    PopCallStack
404:    On Error GoTo 0
405:    Exit Sub

BuildSystemMap_Error:
408:    GlobalErrHandler
409:    Resume BuildSystemMap_Exit

411: End Sub


Public Sub BuildListsSheet()

    'Error Handling
417:    If gcfHandleErrors Then On Error GoTo BuildListsSheet_Error
418:    PushCallStack "mdlUtils.BuildListsSheet"
    
420:    PerfromanceEnabled (True)
    
    'Step 1: clean list sheet
423:    Sheets("רשימות").Activate
424:    ActiveSheet.Range("A3:N900").value = ""
    
    'Step 2: Connect to DB
427:    Dim MyDb As LeumiDB
428:    Dim LastRow As Integer, i As Integer
429:    Dim MyRange As String
430:    Dim data() As String
    
432:    Set MyDb = New LeumiDB: MyDb.InitClass
433:    MyDb.ConnectToDB
    
    'Interface Category
436:    MyDb.pSQLQuery = "select InterfaceCategoryDescription from tblInterfaceCategory"
437:    MyDb.ExecuteSelect
     
439:    LastRow = UBound(MyDb.pDataArray, 2)
440:    MyRange = "A3:A" & LastRow + 3
441:    ReDim data(0 To LastRow, 1 To 1)
442:    For i = 0 To LastRow
443:        data(i, 1) = MyDb.pDataArray(0, i)
444:    Next i
445:    Sheets("רשימות").Range(MyRange) = data
    
    'Interface Type
448:    MyDb.pSQLQuery = "select InterfaceTypeDescription from tblInterfaceType"
449:    MyDb.ExecuteSelect
     
451:    LastRow = UBound(MyDb.pDataArray, 2)
452:    MyRange = "C3:C" & LastRow + 3
453:    ReDim data(0 To LastRow, 1 To 1)
454:    For i = 0 To LastRow
455:        data(i, 1) = MyDb.pDataArray(0, i)
456:    Next i
457:    Sheets("רשימות").Range(MyRange) = data
    
    'Interface KnowledgeLevel
460:    MyDb.pSQLQuery = "select KnoledgeLevelDescription from tblInterfaceKnowledgeLevel"
461:    MyDb.ExecuteSelect
     
463:    LastRow = UBound(MyDb.pDataArray, 2)
464:    MyRange = "E3:E" & LastRow + 3
465:    ReDim data(0 To LastRow, 1 To 1)
466:    For i = 0 To LastRow
467:        data(i, 1) = MyDb.pDataArray(0, i)
468:    Next i
469:    Sheets("רשימות").Range(MyRange) = data
    
    'Skill Type
472:    MyDb.pSQLQuery = "select SkillTypeDescription from tblSkillsType"
473:    MyDb.ExecuteSelect
     
475:    LastRow = UBound(MyDb.pDataArray, 2)
476:    MyRange = "G3:G" & LastRow + 3
477:    ReDim data(0 To LastRow, 1 To 1)
478:    For i = 0 To LastRow
479:        data(i, 1) = MyDb.pDataArray(0, i)
480:    Next i
481:    Sheets("רשימות").Range(MyRange) = data
    
    'Skill Knowledge Level
484:    MyDb.pSQLQuery = "select SkillKnowledgeLevel from tblSkillsKnowledgeLevel"
485:    MyDb.ExecuteSelect
     
487:    LastRow = UBound(MyDb.pDataArray, 2)
488:    MyRange = "I3:I" & LastRow + 3
489:    ReDim data(0 To LastRow, 1 To 1)
490:    For i = 0 To LastRow
491:        data(i, 1) = MyDb.pDataArray(0, i)
492:    Next i
493:    Sheets("רשימות").Range(MyRange) = data
    
    'Skills
496:    MyDb.pSQLQuery = "SELECT tblSkills.SkillDescription, tblSkillsType.SkillTypeDescription" & _
                    " FROM tblSkillsType INNER JOIN tblSkills ON tblSkillsType.SkillTypeId = tblSkills.SkillType;"
498:    MyDb.ExecuteSelect
     
500:    LastRow = UBound(MyDb.pDataArray, 2)
501:    MyRange = "M3:N" & LastRow + 3
502:    ReDim data(0 To LastRow, 0 To 1)
503:    For i = 0 To LastRow
504:        data(i, 0) = MyDb.pDataArray(0, i)
505:        data(i, 1) = MyDb.pDataArray(1, i)
506:    Next i
507:    Sheets("רשימות").Range(MyRange) = data
    
    
'Exit Point
BuildListsSheet_Exit:
512:    PopCallStack
513:    PerfromanceEnabled (False)
514:    On Error GoTo 0
515:    Exit Sub

BuildListsSheet_Error:
518:    GlobalErrHandler
519:    Resume BuildListsSheet_Exit
    
521: End Sub

''''''''''''''''''''''''''''''''''''''''
'           Private Sub BuildArchiveMap()
'
'   Build archive map using BuildMap sub
'
''''''''''''''''''''''''''''''''''''''''
Private Sub BuildArchiveMap()

    'Error Handling
532:    If gcfHandleErrors Then On Error GoTo BuildArchiveMap_Error
533:    PushCallStack "mdlAdminFunctions.BuildArchiveMap"
 
535:    Call BuildMap("Archive")
    
    'Exit Point
BuildArchiveMap_Exit:
539:    PopCallStack
540:    On Error GoTo 0
541:    Exit Sub

BuildArchiveMap_Error:
544:    GlobalErrHandler
545:    Resume BuildArchiveMap_Exit

547: End Sub


