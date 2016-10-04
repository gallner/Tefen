Attribute VB_Name = "mdlPresentation"
Option Explicit



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                           Public Sub FromFormToClass()
'
'   Sub loads class with the infomation from the form:
'
'   1. Loads all the form data into temp array
'   2. Initializes / reInitializes the class to make sure no mix up with old data
'   3. loads all forms' data
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FromFormToClass() As System
    
    'Error Handling
17:    If gcfHandleErrors Then On Error GoTo FromFormToClass_Error
18:    PushCallStack "mdlPresentation.FromFormToClass"
    
    'create a new system object
21:    Dim MySystem As System
22:    Set MySystem = New System: MySystem.InitClass
    
24:    Dim data(1, 33) As String
25:    Dim i, MyRange
       
27:    For i = 0 To 33
28:        MyRange = "b" & (i + 1)
29:        data(0, i) = PrepareStrToDB(ActiveWorkbook.ActiveSheet.Range(MyRange).value)
        
31:        MyRange = "c" & (i + 1)
32:        data(1, i) = PrepareStrToDB(ActiveWorkbook.ActiveSheet.Range(MyRange).value)
33:    Next
    
     
36:    MySystem.SystemName = data(1, 1)                                            'SystemName
    'MySystem.SysId = MySystem.FindSystemId(CStr(MySystem.SystemName))          '"SystemId"  !!!! NO need for System Id at this point !!!!!
38:    MySystem.SubSystemName = data(1, 2)                                         '"SubSystemName"
39:    MySystem.SystemDescription = data(1, 3)                                     '"SystemDescription"
40:    MySystem.KnowledgeConsumer = data(1, 4)                                     '"KnowledgeConsumers"
    
    
    ''' employees
    '
45:    Dim TempStr() As String
46:    TempStr = Split(data(1, 5), ",")
47:    MySystem.RoshAnafBachir = MySystem.UpdateSystemEmployees(TempStr, EmployeePosition.RoshAnafBachir) '"RoshAnafBachir"
    
49:    TempStr = Split(data(1, 6), ",")
50:    MySystem.RoshAnaf = MySystem.UpdateSystemEmployees(TempStr, EmployeePosition.RoshAnaf)     '"RoshAnaf"
    
52:    TempStr = Split(data(1, 7), ",")
53:    MySystem.RoshMador = MySystem.UpdateSystemEmployees(TempStr, EmployeePosition.RoshMador)      '"RoshMador"
    
55:    TempStr = Split(data(1, 8), ",")
56:    MySystem.Workers = MySystem.UpdateSystemEmployees(TempStr, EmployeePosition.KnowledgeExpert) 'System Experts
   '
   '''''
    
60:    If data(1, 9) = "עסקית" Or data(1, 9) = "" Then   '"IsInfrastructure"
61:        MySystem.IsInfrastructure = 0
62:    ElseIf data(1, 9) = "תשתיתית" Then
63:        MySystem.IsInfrastructure = 1
64:    ElseIf data(1, 9) = "עסקית ותשתיתית" Then  'both
65:        MySystem.IsInfrastructure = 2
66:    Else 'not a valid data
67:        ThrowError CustomError.INVALID_FORM_INPUT, "mdlPresentation.FromFormToClass", "Cell C10 contains illegal input. Did you copy paste to the map form?"
68:    End If
    
70:    MySystem.BizDev = data(1, 10) & ", " & data(1, 11)               '"BizEnv"
    
72:    If data(1, 12) = "open" Or data(1, 12) = "MF" Or data(1, 12) = "MF+OPEN" Then '"DevEnv"
73:        MySystem.DevEnv = data(1, 12)
74:    Else
75:        ThrowError CustomError.INVALID_FORM_INPUT, "mdlPresentation.FromFormToClass", "Cell C10 contains illegal input. Did you copy paste to the map form?"
76:    End If
    
78:    MySystem.TechEnv = data(1, 13)                                  '"TechEnv"
79:    MySystem.IsCore = CStrToBool(data(1, 14))                       '"IsCore"
    
   '''dev langs B16
82:    TempStr = Split(data(1, 15), ",")
83:    Call MySystem.UpdateSystemDevLangs(TempStr)
   '
   '''
    
87:    MySystem.IsWebBased = CStrToBool(data(1, 16))                   '"IsWebBased"
        
    '''DB types - open
90:    TempStr = Split(data(1, 17), ",")
91:    Call MySystem.UpdateDbTypesOfSystem(TempStr, True)
    '
    ''''
    
    '''DB types- MF
96:    TempStr = Split(data(1, 18), ",")
97:    Call MySystem.UpdateDbTypesOfSystem(TempStr, False)
    '
    '''''
      
101:    MySystem.NumOfInterfaces = CInt(data(1, 19))                  '"NumOfInterfaces"
102:    MySystem.NumOfCriticalInterfaces = CInt(data(1, 20))          '"NumOfCriticalInterfaces"
103:    MySystem.SystemRisks = data(1, 21)                            '"SystemRisks"
104:    MySystem.PreservationTopics = data(1, 22)                     '"PreservationTopics"
105:    MySystem.PreservationSuggestions = data(1, 23)                '"PreservationSuggestions"
106:    MySystem.IsChangeManagement = YesNoPartial(data(1, 25))      '"IsChangeManagement"
107:    MySystem.IsPlannedToClose = YesNoPartial(data(1, 24))        '"IsPlannedToClose"
108:    MySystem.IsManagementBrief = YesNoPartial(data(1, 26))       '"IsManagementBrief"
109:    MySystem.IsApplicableDocuments = YesNoPartial(data(1, 27))   '"IsApplicableDocuments"
110:    MySystem.IsArchitectureDoc = YesNoPartial(data(1, 28))       '"IsArchitectureDoc"
111:    MySystem.IsShobDoc = YesNoPartial(data(1, 29))               '"IsShobDoc"
112:    MySystem.IsTesting = YesNoPartial(data(1, 30))               '"IsTesting"
113:    MySystem.IsSupportAndOp = YesNoPartial(data(1, 31))          '"IsSupportAndOp"
114:    MySystem.QaPassFail = YesNoPartial(data(1, 32))              '"QaPassFail"
    
116:    MySystem.Comments = data(1, 33)                            'Comments
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'interfaces
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
121:    Dim tempArr() As Variant
122:    Dim LastRow As Integer
     
    'get the range of the end of the interfaces list
125:    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.count, "H").End(xlUp).row
    
    'should be at least one interface. Otherwise: empty!
128:    If LastRow > 2 Then
129:        MyRange = "K" & (LastRow)
        
131:        tempArr = ActiveWorkbook.ActiveSheet.Range("h3", ActiveWorkbook.ActiveSheet.Range(MyRange)).value
132:        Call MySystem.BuildInterfacesArray(tempArr)
        
134:    End If
    
    
    ' Skills
    
    'get the range of the end of the interfaces list
140:    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.count, "S").End(xlUp).row
    
    'should be at least one skill. Otherwise: empty!
143:    If LastRow > 2 Then
144:        MyRange = "U" & (LastRow)
        
146:        tempArr = ActiveSheet.Range("S3", Range(MyRange)).value
        
148:        Call MySystem.BuildSkillsArray(tempArr)
149:    End If
    
    
152:    Set FromFormToClass = MySystem
        
    'clean up
155:    Erase data
156:    Erase tempArr
    
'Exit Point
FromFormToClass_Exit:
160:    PopCallStack
161:    On Error GoTo 0
162:    Exit Function
     
'Error Handling
FromFormToClass_Error:
166:        GlobalErrHandler
167:        MySystem.InitClass
168:        Resume FromFormToClass_Exit
        
    
171: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                Sub FromClassToForm(MySystem As System)
'
'   Sub receives a system as parameter.
'   Sub loads all system parameters from the class and displays them correctly
'   on the current active sheet.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromClassToForm(MySystem As System)
    
    'error handling
185:    If gcfHandleErrors Then On Error GoTo FromClassToForm_Error
186:    PushCallStack "mdlPresentation.FromClassToForm"
    
    'set up the log
189:    Dim MyLog As Log
190:    Set MyLog = New Log: Call MyLog.InitClass
    
    'General Information
193:    Dim data() As String
194:    data = MySystem.getGeneralInformationAsStringArray()
    
196:    If IsArrayAllocated(data) Then
197:        Dim MyRange As String
198:        MyRange = "C2:C34"
199:        ActiveWorkbook.ActiveSheet.Range(MyRange).Select
200:        ActiveWorkbook.ActiveSheet.Range(MyRange) = data
201:    End If
    
    
    'Interfaces
205:    Dim Interfaces() As String
206:    Interfaces = MySystem.GetInterfaceAsStringArray()
    
    'make sure there is at least one iterface
209:    If IsArrayAllocated(Interfaces) Then
          
211:        MyRange = "h3:k" & (3 + UBound(Interfaces))
212:        ActiveWorkbook.ActiveSheet.Range(MyRange).Select
213:        ActiveWorkbook.ActiveSheet.Range(MyRange) = Interfaces
    
215:    End If
    
    
    'Skills
        
220:    Dim Skills() As String
221:    Skills = MySystem.GetSkillsAsStringArray()
    
    'make sure there is at least one skill to show
224:    If IsArrayAllocated(Skills) Then
       
226:        MyRange = "s3:u" & (3 + UBound(Skills))
227:        ActiveWorkbook.ActiveSheet.Range(MyRange).Select
228:        ActiveWorkbook.ActiveSheet.Range(MyRange) = Skills
229:    End If
    
231:    ActiveWorkbook.ActiveSheet.Range("A1").Select
    
    'clean up
234:    Erase data
235:    Erase Interfaces
236:    Erase Skills
    
    
    'Exit Point
FromClassToForm_Exit:
241:    PopCallStack
242:    On Error GoTo 0
243:    Exit Sub
    
    
  'Error Handling
FromClassToForm_Error:
248:        GlobalErrHandler
249:        clearSheet
250:        Resume FromClassToForm_Exit
     
252: End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               Private Sub InsertActiveSheetSystemIntoDb()
'
'   sub enters the active sheet system form into the DB
'   1. Create new system object
'   2. Get the system id from the system name written in the form
'   3. Insert data into DB
'
'   TODO: This is awful!! correct this sub.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InsertActiveSheetSystemIntoDb()

    'Error Handling
268:    If gcfHandleErrors Then On Error GoTo InsertActiveSheetSystemIntoDb_Error
269:    PushCallStack "mdlPresentation.InsertActiveSheetSystemIntoDb"
    
    
272:    Dim MySystem As System
273:    Set MySystem = New System
    
275:    Set MySystem = FromFormToClass()
276:    MySystem.SysId = MySystem.FindSystemId(MySystem.SystemName)
277:    MySystem.FromClassToDb
    
279:    MySystem.FinalizeClass: Set MySystem = Nothing
    
    'Exit Point
InsertActiveSheetSystemIntoDb_Exit:
283:    PopCallStack
284:    On Error GoTo 0
285:    Exit Sub

InsertActiveSheetSystemIntoDb_Error:
288:    GlobalErrHandler
289:    Resume InsertActiveSheetSystemIntoDb_Exit

291: End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               Sub LoadSystemToForm(SystemName As String)
'
'   Sub receives system name as parameter, retrieve data from the DB
'   and then presents it in the form.
'   TODO! write it better
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub LoadSystemToForm(SystemName As String)
    
    'Error Handling
304:    If gcfHandleErrors Then On Error GoTo LoadSystemToForm_Error
305:    PushCallStack "mdlPresentation.LoadSystemToForm"
 
    
308:    Dim MySystem As System
309:    Set MySystem = New System
310:    MySystem.InitClass
    
312:    MySystem.SysId = MySystem.FindSystemId(SystemName)
313:    Sheets(SystemName).Activate
314:    MySystem.FromDbToClass
315:    Call FromClassToForm(MySystem)
    
    'clean up
318:    Call MySystem.FinalizeClass
319:    Set MySystem = Nothing
    
    'Exit Point
LoadSystemToForm_Exit:
323:    PopCallStack
324:    On Error GoTo 0
325:    Exit Sub

LoadSystemToForm_Error:
328:    GlobalErrHandler
329:    Resume LoadSystemToForm_Exit

331: End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       Private Sub clearSheet()
'
'   Sub Cleans the system form from all data
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub clearSheet()

    'Error Handling
343:    If gcfHandleErrors Then On Error GoTo clearSheet_Error
344:    PushCallStack "mdlPresentation.clearSheet"
    
346:    ActiveSheet.Range("c2:c34").value = ""
347:    ActiveSheet.Range("H2:K90").value = ""
348:    ActiveSheet.Range("S3:U90").value = ""
    
    
'Exit Point
clearSheet_Exit:
353:    PopCallStack
354:    On Error GoTo 0
355:    Exit Sub

clearSheet_Error:
358:    GlobalErrHandler
359:    Resume clearSheet_Exit
    
361: End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Sub FromClassToMapSheet(MapType As String, SystemName As String)
'
'   Sub converts system presentation from Form to system's map.
'   Sub receives as parameter a MapType (Atcive System or Archived System)
'   and a system to print, and prints the system's details into the correct
'   sheet:
'   MapType = System: System's map
'   MapType = Archive: Archive systems
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FromClassToMapSheet(MapType As String, SystemName As String)
    
    'Error Handling
377:    If gcfHandleErrors Then On Error GoTo FromClassToMapSheet_Error
378:    PushCallStack "mdlPresentation.FromClassToMapSheet"
    
380:    Dim MySystem As System
381:    Set MySystem = New System: MySystem.InitClass
382:    MySystem.FindSystemId (SystemName)
    
384:    Call MySystem.FromDbToClass
    
386:    Dim MyRange As String, MyData() As String
387:    Dim LastRow As Integer
    
389:    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.count, "A").End(xlUp).row + 1
    
391:    MyRange = "A" & LastRow & ":AG" & LastRow
392:    MyData = MySystem.getGeneralInformationAsStringArray()
    
394:    Dim data() As String
395:    ReDim data(0 To 0, 0 To UBound(MyData, 1)) As String
396:    Dim i As Integer
    
398:    For i = 0 To UBound(MyData, 1)
399:        data(0, i) = MyData(i, 1)
400:    Next
    
402:    Dim SheetName As String
403:    If MapType = "System" Then
404:        SheetName = "מפת המערכת"
405:    Else
406:        SheetName = "ארכיון"
407:    End If
408:    Sheets(SheetName).Range(MyRange) = data
    
    'clean up
411:    MySystem.FinalizeClass: Set MySystem = Nothing
412:    Erase MyData
413:    Erase data
    
    'Exit Point
FromClassToMapSheet_Exit:
417:    PopCallStack
418:    On Error GoTo 0
419:    Exit Sub

FromClassToMapSheet_Error:
422:    GlobalErrHandler
423:    Resume FromClassToMapSheet_Exit
    
425: End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Public Sub DisplayArray(sArrayToDisplay() As String)
'
'   sub receives a 2 dimention array, create a new worksheet and
'   displays the array in the newly created worksheet
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DisplayArray(sArrayToDisplay() As String, sListOfSystems() As String, sEmployees() As String)

    'Error Handling
437:    If gcfHandleErrors Then On Error GoTo DisplayArray_Error
438:    PushCallStack "mdlPresentation.DisplayArray"

    
   
442:    Application.ScreenUpdating = False
443:    With Sheets("פערי כישורים")
444:        .Visible = True
445:        .Copy After:=Sheets(Sheets.count)
446:        .Visible = False
447:    End With
    
449:    ActiveSheet.Name = "Skill Gap"
450:    Application.ScreenUpdating = True
    
    
453:    ActiveSheet.Range("B1") = Time()
454:    ActiveSheet.Range("B2") = Join(sListOfSystems, ",")
455:    ActiveSheet.Range("B3") = Join(sEmployees, ",")
    
    'put the info on the sheet
458:    ActiveSheet.Range(Cells(8, 3), Cells(UBound(sArrayToDisplay, 1) + 8, UBound(sArrayToDisplay, 2) + 3)) = sArrayToDisplay
    
'Exit Point
DisplayArray_Exit:
462:    PopCallStack
463:    On Error GoTo 0
464:    Exit Sub

DisplayArray_Error:
467:    GlobalErrHandler
468:    Resume DisplayArray_Exit
469: End Sub
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Public Sub DisplayArray(sArrayToDisplay() As String)
'
'   sub receives a 2 dimention array, create a new worksheet and
'   displays the array in the newly created worksheet
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DisplayArray2(sArrayToDisplay() As String)

    'Error Handling
480:    If gcfHandleErrors Then On Error GoTo DisplayArray_Error
481:    PushCallStack "mdlPresentation.DisplayArray"

    
484:   Application.ScreenUpdating = False
    
486:    Application.Sheets.Add
    
488:   Application.ScreenUpdating = True
    
    
    
    
    'put the info on the sheet
494:    ActiveSheet.Range(Cells(1, 1), Cells(UBound(sArrayToDisplay, 1) + 1, UBound(sArrayToDisplay, 2) + 1)) = sArrayToDisplay
    
'Exit Point
DisplayArray_Exit:
498:    PopCallStack
499:    On Error GoTo 0
500:    Exit Sub

DisplayArray_Error:
503:    GlobalErrHandler
504:    Resume DisplayArray_Exit
505: End Sub


Public Sub BuildSkillsMap()
    
    'Error Handling
511:    If gcfHandleErrors Then On Error GoTo BuildSkillsMap_Error
512:    PushCallStack "mdlPresentation.BuildSkillsMap"
    
    'init
515:    Dim i As Integer, j As Integer
516:    Dim itemToAdd As String
    
    'connect to DB
519:    Dim MyDb As LeumiDB
520:    Set MyDb = New LeumiDB
521:    Call MyDb.ConnectToDB
522:    Dim finalArr() As String
523:    Dim dListOfSkills As Dictionary
524:    Dim dListOfSystems As Dictionary
    
526:    Set dListOfSkills = New Dictionary
527:    Set dListOfSystems = New Dictionary
     
    'get all the skills in the system and their position
530:    MyDb.pSQLQuery = "SELECT SkillDescription FROM tblSkills ;"
531:    MyDb.ExecuteSelect
    
533:    If IsArrayAllocated(MyDb.pDataArray) Then
534:        ReDim finalArr(UBound(MyDb.pDataArray, 2) + 1, 1)
535:        For i = 0 To UBound(finalArr, 1) - 1
536:            If Not dListOfSkills.Exists(MyDb.pDataArray(0, i)) Then
537:                finalArr(i + 1, 0) = MyDb.pDataArray(0, i)
538:                dListOfSkills.Add MyDb.pDataArray(0, i), i + 1
539:            End If
540:        Next
541:    End If
    
    'get the list of the systems and their position
544:    MyDb.pSQLQuery = "SELECT tblSystems.SystemName FROM tblSystems inner join tblSystemDescription on " & _
                     " tblSystems.systemid = tblSystemDescription.systemid where tblSystems.status=true;"
546:    MyDb.ExecuteSelect

548:    If IsArrayAllocated(MyDb.pDataArray) Then
549:        i = UBound(finalArr, 1)
550:        ReDim Preserve finalArr(i, UBound(MyDb.pDataArray, 2) + 1)
551:        For i = 0 To UBound(finalArr, 2) - 1
552:            If Not dListOfSystems.Exists(MyDb.pDataArray(0, i)) Then
553:                finalArr(0, i + 1) = MyDb.pDataArray(0, i)
554:                dListOfSystems.Add MyDb.pDataArray(0, i), i + 1
555:            End If
556:        Next
557:    End If
    
    'start calculating
        
    'get the table of systems and skills
562:    MyDb.pSQLQuery = "SELECT tblSystems.SystemName, tblSkills.SkillDescription " & _
                    " FROM tblSystems INNER JOIN (tblSkills INNER JOIN tblSystemSkills ON tblSkills.SkillId " & _
                    "= tblSystemSkills.SkillId) ON tblSystems.SystemId = tblSystemSkills.SystemId " & _
                    " WHERE tblSystems.Status=True;"

567:    MyDb.ExecuteSelect

569:    If IsArrayAllocated(MyDb.pDataArray) Then
570:        Dim tempArr()
571:        tempArr = MyDb.pDataArray
        
        'search for a match
574:        For i = 0 To UBound(tempArr, 2)
575:                If dListOfSystems.Exists(tempArr(0, i)) And dListOfSkills.Exists(tempArr(1, i)) Then
576:                    finalArr(CInt(dListOfSkills(tempArr(1, i))), CInt(dListOfSystems(tempArr(0, i)))) = "+"
577:                End If
578:        Next i
579:    End If
    
    'print results
582:    Call DisplayArray2(finalArr)
    
'exit point
BuildSkillsMap_Exit:
586:    PopCallStack
587:    On Error GoTo 0
588:    Exit Sub

BuildSkillsMap_Error:
591:    GlobalErrHandler
592:    Resume BuildSkillsMap_Exit
593: End Sub

Public Function isValidSystemForm() As Boolean
    
597:    Dim result As Boolean
598:    result = True 'start optimistic
    
    'not the right form
601:    If ActiveSheet.Range("a1").value <> "מיפוי מערכת" Then isValidSystemForm = False
    
    'no system name
604:    If Len(ActiveSheet.Cell("c2").value) = 0 Then isValidSystemForm = False
    
606:    isValidSystemForm = result
607: End Function

