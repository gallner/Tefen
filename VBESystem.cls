VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "System"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Description:  This class is a container for a single system
'               description of the Leumi project.
'               This is the main class in the system, and contains
'               most of the application logic
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                               Class Memebers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const MainFrame = 0
Private Const OpenSystem = 1

Private cListToUpdate As Dictionary
Private iSysId As Integer
Private sSystemName As String
Private sSubSystemName As String
Private sSystemDescription As String
Private sKnowledgeConsumer As String
Private iIsInfrastructure As Integer
Private sBizDev As String
Private sDevEnv As String
Private sTechEnv As String
Private bIsCore As Boolean
Private bIsWebBased As Boolean
Private iNumOfInterfaces As Integer
Private iNumOfCriticalInterfaces As Integer
Private sSystemRisks As String
Private sPreservationTopics As String
Private sPreservationSuggestions As String
Private iIsChangeManagement As Integer
Private iIsPlannedToClose As Integer
Private iIsManagementBrief As Integer
Private iIsApplicableDocuments As Integer
Private iIsArchitectureDoc As Integer
Private iIsShobDoc As Integer
Private iIsTesting As Integer
Private iIsSupportAndOp As Integer
Private iQaPassFail As Integer
Private sComments As String

Private eRoshAnaf() As employee
Private eRoshAnafBachir() As employee
Private eRoshMador() As employee
Private eWorkers() As employee
Private sProgrammingLang() As String
Private sOpenDB() As String
Private sMFDB() As String
Private IInterfaces() As Interface
Private SSkills() As Skill

Private MyDb As LeumiDB
Private MyLog As Log


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                               Public Sub InitClass
'   Sub serves as sort of constructor for the class.
'   It initializes relevant parmeters:
'   1. Strings will be assigned value ""
'   2. Integers are set to -1
'   3. Boolean values are set to False
'   4. Arrays are set to 1 spots (redim (0)).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitClass()
    
    'Error Handling
78:    If gcfHandleErrors Then On Error GoTo InitClass_Error
79:    PushCallStack "System.InitClass"

    
82:    Set cListToUpdate = New Dictionary
    
84:    iSysId = -1
85:    sSubSystemName = ""
86:    sSystemDescription = ""
87:    sKnowledgeConsumer = ""
      
89:    iIsInfrastructure = -1
90:    sBizDev = ""
91:    sDevEnv = ""
92:    sTechEnv = ""
93:    bIsCore = False
94:    bIsWebBased = False
95:    iNumOfInterfaces = -1
96:    iNumOfCriticalInterfaces = -1
97:    sSystemRisks = ""
98:    sPreservationTopics = ""
99:    sPreservationSuggestions = ""
100:    iIsChangeManagement = -1
101:    iIsPlannedToClose = -1
102:    iIsManagementBrief = -1
103:    iIsApplicableDocuments = -1
104:    iIsArchitectureDoc = -1
105:    iIsShobDoc = -1
106:    iIsTesting = -1
107:    iIsSupportAndOp = -1
108:    iQaPassFail = -1
    
110:    ReDim eRoshAnaf(0)
111:    ReDim eRoshAnafBachir(0)
112:    ReDim eRoshMador(0)
113:    ReDim eWorkers(0)
114:    ReDim sProgrammingLang(0)
115:    ReDim sOpenDB(0)
116:    ReDim sMFDB(0)
117:    ReDim IInterfaces(0)
118:    ReDim SSkills(0)
    
120:    Set MyDb = New LeumiDB
121:    Set MyLog = New Log
    
'Exit Point
InitClass_Exit:
125:    PopCallStack
126:    On Error GoTo 0
127:    Exit Sub

InitClass_Error:
130:    GlobalErrHandler
131:    Resume InitClass_Exit
    
133: End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                      Public Sub FinalizeClass()
'   Sub serves as sort of destructor for the class.
'   set all objects to nothing and erase all dynamic arrays
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FinalizeClass()
    
    'Error Handling
144:    If gcfHandleErrors Then On Error GoTo FinalizeClass_Error
145:        PushCallStack "System.FinalizeClass"

    
148:    Set cListToUpdate = Nothing
149:    Erase eRoshAnaf
150:    Erase eRoshAnafBachir
151:    Erase eRoshMador
152:    Erase eWorkers
153:    Erase sProgrammingLang
154:    Erase sOpenDB
155:    Erase sMFDB
156:    Erase IInterfaces
157:    Erase SSkills
    
159:    MyDb.FinalizeClass: Set MyDb = Nothing
160:    MyLog.FinalizeClass: Set MyLog = Nothing

'Exit Point
FinalizeClass_Exit:
164:    PopCallStack
165:    On Error GoTo 0
166:    Exit Sub

FinalizeClass_Error:
169:    GlobalErrHandler
170:    Resume FinalizeClass_Exit
    
172: End Sub




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Public Sub AddInterfaceToClass(Name As String, Category As String, InterfaceType As String, KnowledgeLevel As String)
'
'   Sub adds a new interface to the interfaces array.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddInterfaceToClass(Name As String, Category As String, InterfaceType As String, KnowledgeLevel As String)
    
    'Error Handling
186:    If gcfHandleErrors Then On Error GoTo AddInterfaceToClass_Error
187:        PushCallStack "System.AddInterfaceToClass"

    
190:    If Not IInterfaces(UBound(IInterfaces)) Is Nothing Then ReDim Preserve IInterfaces(UBound(IInterfaces) + 1)
    
192:    Set IInterfaces(UBound(IInterfaces)) = New Interface
193:    Call IInterfaces(UBound(IInterfaces)).PopulateClass(Name, Category, InterfaceType, KnowledgeLevel)
    
'Exit Point
AddInterfaceToClass_Exit:
197:    PopCallStack
198:    On Error GoTo 0
199:    Exit Sub

AddInterfaceToClass_Error:
202:    GlobalErrHandler
203:    Resume AddInterfaceToClass_Exit
204: End Sub




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Public Sub AddSkillToClass(Name As String, sSkillType As String, sSkillType As Integer, Level As String)
'
'   Sub adds a new skill to the skills array.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddSkillToClass(Name As String, sSkillType As String, iSkillType As Integer, Level As String)
    
    'Error Handling
218:    If gcfHandleErrors Then On Error GoTo AddSkillToClass_Error
219:        PushCallStack "AddSkillToClass"
    
221:    If Not SSkills(UBound(SSkills)) Is Nothing Then ReDim Preserve SSkills(UBound(SSkills) + 1)
    
223:    Set SSkills(UBound(SSkills)) = New Skill
224:    Call SSkills(UBound(SSkills)).PopulateClass(Name, sSkillType, iSkillType, Level)
    
    'Exit Point
AddSkillToClass_Exit:
228:    PopCallStack
229:    On Error GoTo 0
230:    Exit Sub

AddSkillToClass_Error:
233:    GlobalErrHandler
234:    Resume AddSkillToClass_Exit

236: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                           Sub FromDbToClass
'
'   Sub loads all system parameters from tblSystem and tblSystemDescription
'   into the class. Sub assumes the correct SystemId is stored in the
'   SysId property.
'   Sub uses private functions go fill data from other tables:
'   1. Rosh Anaf
'   2. Rosh Anaf Bachir
'   3. Rosh Mador
'   4. eWorkers
'
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromDbToClass()
    
    'Error Handling
256:    If gcfHandleErrors Then On Error GoTo FromDbToClass_Error
257:        PushCallStack "System.FromDbToClass"
    
259:    Call MyDb.ConnectToDB
     
    'no system id to load
262:    If Me.SysId = -1 Then _
        ThrowError CustomError.SYSTEM_ID_IS_MISSING, "System.FromDbToClass", "System id is not present"
    
    'load system infromation
266:    MyDb.pSQLQuery = "select * from tblSystemDescription where SystemId=" & Me.SysId & ";"
    
268:    MyDb.ExecuteSelect
    
            
271:    If Not IsArrayAllocated(MyDb.pDataArray) Then _
            ThrowError CustomError.NO_DATA_FOUND_IN_DB, "System.FromDbToClass", "System was not found in DB"
    
    'Copy Results
275:    Dim myDataArr()
276:    ReDim myDataArr(UBound(MyDb.pDataArray))
277:    Dim i As Integer
278:    For i = 0 To UBound(MyDb.pDataArray)
279:        myDataArr(i) = ConvertNullToEmptyString(MyDb.pDataArray(i, 0))
280:    Next
    
    
    
284:    Me.SystemName = Me.FindSystemName(CInt(myDataArr(0)))
285:    SubSystemName = ConvertNullToEmptyString(myDataArr(1))
286:    SystemDescription = ConvertNullToEmptyString(myDataArr(2))
287:    KnowledgeConsumer = ConvertNullToEmptyString(myDataArr(3))
    
289:    Dim TempEmployeeIds() As Integer
    
'    eRoshAnafBachir.EmpId = myDataArr(5) 'TODO - at the end, delete the colomns and change numbers
292:    TempEmployeeIds = GetSystemEmployeesFromDb(EmployeePosition.RoshAnafBachir)
    
294:    If IsArrayAllocated(TempEmployeeIds) Then
295:        For i = 0 To UBound(TempEmployeeIds)
296:            If i > UBound(eRoshAnafBachir) Then ReDim Preserve eRoshAnafBachir(i)
297:            Set eRoshAnafBachir(i) = New employee
298:            eRoshAnafBachir(i).EmpId = TempEmployeeIds(i)
299:            Call GetEmployeeInfoFromDb(eRoshAnafBachir(i))
300:        Next i
301:    End If
    
    'eRoshAnaf.EmpId = myDataArr(4)
304:    TempEmployeeIds = GetSystemEmployeesFromDb(EmployeePosition.RoshAnaf)
    
306:    If IsArrayAllocated(TempEmployeeIds) Then
307:        For i = 0 To UBound(TempEmployeeIds)
308:            If i > UBound(eRoshAnaf) Then ReDim Preserve eRoshAnaf(i)
309:            Set eRoshAnaf(i) = New employee
310:            eRoshAnaf(i).EmpId = TempEmployeeIds(i)
311:            Call GetEmployeeInfoFromDb(eRoshAnaf(i))
312:        Next i
313:    End If
      
    'eRoshMador.EmpId = myDataArr(6)
316:    TempEmployeeIds = GetSystemEmployeesFromDb(EmployeePosition.RoshMador)
    
318:    If IsArrayAllocated(TempEmployeeIds) Then
319:        For i = 0 To UBound(TempEmployeeIds)
320:            If i > UBound(eRoshMador) Then ReDim Preserve eRoshMador(i)
321:            Set eRoshMador(i) = New employee
322:            eRoshMador(i).EmpId = TempEmployeeIds(i)
323:            Call GetEmployeeInfoFromDb(eRoshMador(i))
324:        Next i
325:    End If
    
    'System experts.
    'Call GetSystemExpertsFromDb
329:    TempEmployeeIds = GetSystemEmployeesFromDb(EmployeePosition.KnowledgeExpert)
330:    If IsArrayAllocated(TempEmployeeIds) Then
331:        For i = 0 To UBound(TempEmployeeIds)
332:            If i > UBound(eWorkers) Then ReDim Preserve eWorkers(i)
333:            Set eWorkers(i) = New employee
334:            eWorkers(i).EmpId = TempEmployeeIds(i)
335:            Call GetEmployeeInfoFromDb(eWorkers(i))
336:        Next i
337:    End If
    
339:    IsInfrastructure = myDataArr(7)
340:    BizDev = myDataArr(8)
341:    DevEnv = myDataArr(9)
342:    TechEnv = myDataArr(10)
343:    IsCore = myDataArr(11)
    
    'load dev langs into class member: sProgrammingLang
346:    Call GetDevLangsInfoFromDB
        
348:    IsWebBased = myDataArr(12)
    
    'list of open databases
351:    Call GetSystemDBInforamtionFromDb(OpenSystem)
    
    'list of mainframe databases
354:    Call GetSystemDBInforamtionFromDb(MainFrame)
    
356:    NumOfInterfaces = myDataArr(13)
357:    NumOfCriticalInterfaces = myDataArr(14)
358:    SystemRisks = myDataArr(15)
359:    PreservationTopics = myDataArr(16)
360:    PreservationSuggestions = myDataArr(17)
361:    IsChangeManagement = myDataArr(18)
362:    IsPlannedToClose = myDataArr(19)
363:    IsManagementBrief = myDataArr(20)
364:    IsApplicableDocuments = myDataArr(21)
365:    IsArchitectureDoc = myDataArr(22)
366:    IsShobDoc = myDataArr(23)
367:    IsTesting = myDataArr(24)
368:    IsSupportAndOp = myDataArr(25)
369:    QaPassFail = myDataArr(26)
370:    Comments = myDataArr(27)
    
    'Interfaces
373:        Call GetSystemInterfaceFromDB
    
    'Skills
376:    Call GetSystemSkillsFromDB
     
 
'Exit Point
FromDbToClass_Exit:
381:    PopCallStack
382:    On Error GoTo 0
383:    Exit Sub

FromDbToClass_Error:
386:    GlobalErrHandler
387:    Call Me.FinalizeClass
388:    MsgBox "אירעה שגיאה בטעינת נתונים ממסד הנתונים" & Err.Description, vbCritical, "שגיאה"
389:    Resume FromDbToClass_Exit
 
 
392: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                           Sub GetEmployeeInfoFromDb(ByRef emp As Employee)
'
'   Sub Receives an employee data type and
'   loads The following info into the data type from tblEmployees:
'   1. Employee First Name
'   2. Employee Last Name
'
'   Sub assumes that the emp already contains employee id.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetEmployeeInfoFromDb(ByRef Emp As employee)
         
    'Error Handling
409:    If gcfHandleErrors Then On Error GoTo GetEmployeeInfoFromDb_Error
410:    PushCallStack "System.GetEmployeeInfoFromDb"
    
    'connect to DB
413:    Call MyDb.ConnectToDB
     
415:    If Emp.EmpId = -1 Then _
        ThrowError CustomError.MANDATORY_ID_IS_MISSING, "System.GetEmployeeInfoFromDb", "Employee Id is not present"
    
    'load system infromation
419:    MyDb.pSQLQuery = "select * from tblEmployees where EmployeeId=" & Emp.EmpId & ";"
420:    MyDb.ExecuteSelect
    
    'update information
423:    If IsArrayAllocated(MyDb.pDataArray) Then
424:        Emp.FirstName = MyDb.pDataArray(1, 0)
425:        Emp.LastName = MyDb.pDataArray(2, 0)
426:    End If
    
    

'Exit Point
GetEmployeeInfoFromDb_Exit:
432:    PopCallStack
433:    On Error GoTo 0
434:    Exit Sub

GetEmployeeInfoFromDb_Error:
437:    GlobalErrHandler
438:    Resume GetEmployeeInfoFromDb_Exit
439: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                           Private Sub GetDevLangsInfoFromDB()
'
'   Sub loads development languages from qryDevLangs, using system id from
'   tblSystemDevLang. It loads the dev langs into class array: sProgrammingLang
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetDevLangsInfoFromDB()
    
    'Error Handling
452:    If gcfHandleErrors Then On Error GoTo GetDevLangsInfoFromDB_Error
453:    PushCallStack "System.GetDevLangsInfoFromDB"
    
    'connect to DB
456:    Call MyDb.ConnectToDB
         
    'load system infromation
459:    MyDb.pSQLQuery = "SELECT tblSkills.SkillDescription FROM tblSkills INNER JOIN tblSystemDevLang ON tblSkills.skillId = tblSystemDevLang.DevLangId WHERE tblSystemDevLang.SystemId=" & Me.SysId & ";"
460:    MyDb.ExecuteSelect
    
    'update information
463:    Dim i As Integer
      
465:    If IsArrayAllocated(MyDb.pDataArray) Then
466:        ReDim sProgrammingLang(UBound(MyDb.pDataArray, 2))
467:        For i = 0 To UBound(MyDb.pDataArray, 2)
468:            sProgrammingLang(i) = MyDb.pDataArray(0, i)
469:        Next
470:    End If
    
    
    
'Exit Point
GetDevLangsInfoFromDB_Exit:
476:    PopCallStack
477:    On Error GoTo 0
478:    Exit Sub

GetDevLangsInfoFromDB_Error:
481:    GlobalErrHandler
482:    Resume GetDevLangsInfoFromDB_Exit
    
484: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                           Private Sub GetSystemInterfaceFromDB()
'
'   Sub loads system interfaces from qrySystemInterfaces, using system id from
'   Class. It loads the interfaces into class array: IInterfaces
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetSystemInterfaceFromDB()
    
    'Error Handling
497:    If gcfHandleErrors Then On Error GoTo GetSystemInterfaceFromDB_Error
498:    PushCallStack "System.GetSystemInterfaceFromDB"
    
    'connect to DB
501:    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "System.GetSystemInterfaceFromDB", "DB.ConnectToDB returned FALSE"
         
    'load system infromation
505:    MyDb.pSQLQuery = "SELECT InterfaceDescription,InterfaceCategoryDescription, InterfaceTypeDescription,KnoledgeLevelDescription FROM qrySystemInterfaces WHERE SystemId=" & Me.SysId & ";"
506:    MyDb.ExecuteSelect
    
    'update information
509:    Dim i As Integer
      
511:    If IsArrayAllocated(MyDb.pDataArray) Then
512:        ReDim IInterfaces(UBound(MyDb.pDataArray, 2))
513:        For i = 0 To UBound(IInterfaces)
514:            Set IInterfaces(i) = New Interface
515:            IInterfaces(i).InterfaceName = MyDb.pDataArray(0, i)
516:            IInterfaces(i).InterfaceCategory = ConvertNullToEmptyString(MyDb.pDataArray(1, i))
517:            IInterfaces(i).InterfaceType = ConvertNullToEmptyString(MyDb.pDataArray(2, i))
518:            IInterfaces(i).KnowledgeLevel = ConvertNullToEmptyString(MyDb.pDataArray(3, i))
            
            
521:        Next
522:    End If
    
    
    
'Exit Point
GetSystemInterfaceFromDB_Exit:
528:    PopCallStack
529:    On Error GoTo 0
530:    Exit Sub

GetSystemInterfaceFromDB_Error:
533:    GlobalErrHandler
534:    Resume GetSystemInterfaceFromDB_Exit
    
536: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                           Private Sub GetSystemSkillsFromDB()
'
'   Sub loads system Skills from qrySystemSkills, using system id from
'   Class. It loads the Skills into class array: SSkills
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetSystemSkillsFromDB()
    
    'Error Handling
549:    If gcfHandleErrors Then On Error GoTo GetSystemSkillsFromDB_Error
550:    PushCallStack "System.GetSystemSkillsFromDB"
    
    'connect to DB
553:    If Not MyDb.ConnectToDB Then _
         ThrowError CustomError.CONNECTION_TO_DB_FAIL, "System.GetSystemSkillsFromDB", "DB.ConnectToDB returned FALSE"
         
    'load system infromation
557:    MyDb.pSQLQuery = "SELECT SkillDescription,SkillTypeDescription, SkillKnowledgeLevel FROM qrySystemSkills WHERE SystemId=" & Me.SysId & ";"
558:    MyDb.ExecuteSelect
    
    'update information
561:    Dim i As Integer
      
563:    If IsArrayAllocated(MyDb.pDataArray) Then
564:        ReDim SSkills(UBound(MyDb.pDataArray, 2))
565:        For i = 0 To UBound(SSkills)
566:            Set SSkills(i) = New Skill
567:            SSkills(i).SkillName = ConvertNullToEmptyString(MyDb.pDataArray(0, i))
568:            SSkills(i).SkillType = ConvertNullToEmptyString(MyDb.pDataArray(1, i))
569:            SSkills(i).KnowledgeLevel = ConvertNullToEmptyString(MyDb.pDataArray(2, i))
            
            
572:        Next
573:    End If
    
    
    
'Exit Point
GetSystemSkillsFromDB_Exit:
579:    PopCallStack
580:    On Error GoTo 0
581:    Exit Sub

GetSystemSkillsFromDB_Error:
584:    GlobalErrHandler
585:    Resume GetSystemSkillsFromDB_Exit
586: End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Function GetSystemEmployeesFromDb(Optional RoleInSystem As Integer) As Integer()
'
'
'   Function recieves a Role type of employees in the system (e.g RoshMador, RoshAnaf, System
'   expert etc.), looks the relevant employees by role in tblSydtemEmployees and
'   builds an array of integers, each integer represents an employee with relevan role working
'   on the system.
'   Function return an array of integers - Employee Id's
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetSystemEmployeesFromDb(Optional RoleInSystem As Integer) As Integer()
    
    'Error Handling
604:    If gcfHandleErrors Then On Error GoTo GetSystemEmployeesFromDb_Error
605:    PushCallStack "System.GetSystemEmployeesFromDb"
    
607:    Dim ResultArray() As Integer
    
    'connect to DB
610:    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "System.GetSystemEmployeesFromDb", "DB.ConnectToDB returned FALSE"
         
    'load system infromation
614:    If RoleInSystem > 0 Then
615:        MyDb.pSQLQuery = "SELECT EmployeeId FROM tblSystemEmployees WHERE SystemId=" & _
        Me.SysId & " and RoleInSystem=" & RoleInSystem & ";"
617:    Else
618:        MyDb.pSQLQuery = "SELECT EmployeeId FROM tblSystemEmployees WHERE SystemId=" & Me.SysId & ";"
619:    End If
    
621:    MyDb.ExecuteSelect
    
    'get the employee id's
624:    Dim i As Integer
625:    If Not IsArrayAllocated(MyDb.pDataArray) Then GoTo GetSystemEmployeesFromDb_Exit

627:    ReDim ResultArray(UBound(MyDb.pDataArray, 2))
628:    For i = 0 To UBound(MyDb.pDataArray, 2)
629:        ResultArray(i) = MyDb.pDataArray(0, i)
630:    Next
    

    'returned value
634:    GetSystemEmployeesFromDb = ResultArray
    
   'Exit Point
GetSystemEmployeesFromDb_Exit:
638:    PopCallStack
639:    On Error GoTo 0
640:    Exit Function

GetSystemEmployeesFromDb_Error:
643:    GlobalErrHandler
644:    Erase ResultArray
645:    GetSystemEmployeesFromDb = ResultArray
646:    Resume GetSystemEmployeesFromDb_Exit
    
648: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                 Private Sub GetSystemDBInforamtionFromDb(DBType As Boolean)
'
'   Sub loads system Database information from tblSystemDbTypes,
'   using system skills information from tblSkills.
'   Sub receives DBType (const member of class - MF or OPen) and uses it
'   to load the appropriate DB types and put it in the correct member: sMFDB / sOpenDB
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetSystemDBInforamtionFromDb(DBType As Boolean)
    
    'Error Handling
663:    If gcfHandleErrors Then On Error GoTo GetSystemDBInforamtionFromDb_Error
664:    PushCallStack "System.GetSystemDBInforamtionFromDb"
    
    'connect to DB
667:    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "System.GetSystemDBInforamtionFromDb", "DB.ConnectToDB returned FALSE"
     
    'load system infromation
671:    MyDb.pSQLQuery = "SELECT tblSkills.SkillDescription FROM tblSystemDbTypes INNER JOIN tblSkills ON tblSystemDbTypes.DBTypeId = tblSkills.SkillId WHERE tblSystemDbTypes.SystemID=" & Me.SysId & " AND tblSystemDbTypes.openORMF=" & DBType & ";"
    
673:    MyDb.ExecuteSelect
    
    'get the employee id's
676:    Dim i As Integer
677:    Dim tmpArr() As String
    
679:    If IsArrayAllocated(MyDb.pDataArray) Then
680:        ReDim tmpArr(UBound(MyDb.pDataArray, 2))
681:        For i = 0 To UBound(MyDb.pDataArray, 2)
682:             tmpArr(i) = MyDb.pDataArray(0, i)
683:        Next
684:    Else
685:    End If
    
687:    If IsArrayAllocated(tmpArr) Then
        'put results in the right place - MF or open
689:        If DBType = MainFrame Then
690:            sMFDB = tmpArr
691:        Else
692:            sOpenDB = tmpArr
693:        End If
694:    End If
        
    
        
'Exit Point
GetSystemDBInforamtionFromDb_Exit:
700:    PopCallStack
701:    On Error GoTo 0
702:    Exit Sub

GetSystemDBInforamtionFromDb_Error:
705:    GlobalErrHandler
706:    Resume GetSystemDBInforamtionFromDb_Exit
    
708: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                 Public Function FindSystemId(SystemName As String) As Intege
'
'   function receives system name as parameter, and queries tblSystems to receive systemId.
'   If system name was not found, -1 is returned.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FindSystemId(SystemName As String) As Integer
    
    'Error Handling
721:    If gcfHandleErrors Then On Error GoTo FindSystemId_Error
722:    PushCallStack "System.FindSystemId"
     
    'connect to DB
725:    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "System.FindSystemId", "DB.ConnectToDB returned FALSE"
    
    'build sql query
729:    MyDb.pSQLQuery = "select SystemId from tblSystems where SystemName='" & SystemName & "';"
730:    MyDb.ExecuteSelect
    
732:    If IsArrayAllocated(MyDb.pDataArray) Then
733:        Me.SysId = CInt(MyDb.pDataArray(0, 0))
734:    End If
    
    
737:    FindSystemId = Me.SysId
       
'Exit Point
FindSystemId_Exit:
741:    PopCallStack
742:    On Error GoTo 0
743:    Exit Function

FindSystemId_Error:
746:    GlobalErrHandler
747:    FindSystemId = -1
748:    Resume FindSystemId_Exit
    
750: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                 Public Function FindGetSystemId(SystemName As String) As Intege
'
'   function receives system name as parameter, and queries tblSystems to receive systemId.
'   If system name was not found, a new system name is inserted into the DB
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FindGetSystemId(SystemName As String) As Integer
    
    'Error Handling
763:    If gcfHandleErrors Then On Error GoTo FindGetSystemId_Error
764:    PushCallStack "System.FindGetSystemId"
    
    
    'connect to DB
768:    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "System.FindGetSystemId", "DB.ConnectToDB returned FALSE"
    
    
    'build sql query
773:    MyDb.pSQLQuery = "select SystemId from tblSystems where SystemName='" & SystemName & "';"
774:    MyDb.ExecuteSelect
    
776:    If Not IsArrayAllocated(MyDb.pDataArray) Then
777:        Me.SystemName = SystemName
778:        Me.SysId = CreateNewSystemRecord(Me.SystemName)
779:    Else
780:        Me.SysId = CInt(MyDb.pDataArray(0, 0))
781:    End If
    
    
784:    FindGetSystemId = Me.SysId
       
'Exit Point
FindGetSystemId_Exit:
788:    PopCallStack
789:    On Error GoTo 0
790:    Exit Function

FindGetSystemId_Error:
793:    GlobalErrHandler
794:    FindGetSystemId = -1
795:    Resume FindGetSystemId_Exit
    
797: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                 Public Function FindSystemName(SystemId As Integer) As String
'
'   function receives system Id as parameter, and queries tblSystems to receive system Name.
'   If system name was not found, "" is returned.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FindSystemName(SystemId As Integer) As String
    
    'Error Handling
810:    If gcfHandleErrors Then On Error GoTo FindSystemName_Error
811:    PushCallStack "System.FindSystemName"
    
813:    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "System.FindSystemName", "DB.ConnectToDB returned FALSE"
    
    'build sql query
817:    MyDb.pSQLQuery = "select SystemName from tblSystems where SystemID=" & SystemId & ";"
818:    MyDb.ExecuteSelect
    
820:    Me.SystemName = CStr(MyDb.pDataArray(0, 0))
821:    FindSystemName = Me.SystemName
   
    
    
'Exit Point
FindSystemName_Exit:
827:    PopCallStack
828:    On Error GoTo 0
829:    Exit Function

FindSystemName_Error:
832:    GlobalErrHandler
833:    FindSystemName = ""
834:    Resume FindSystemName_Exit
 
836: End Function




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'     Public Sub UpdateSystemEmployees(Data() As String, ByRef Employees() As Employee, Optional role As Integer)
'
'   Sub receives a list of employees whoa are working on the system.
'   Sub builds out of every name a new Employee class and populate the
'   data in the class. The employees are located in an Employee array class member,
'   called eWorkers.
'   Finally, CListToUpdate is being called to flag a new change that needs to go into the
'   DB
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function UpdateSystemEmployees(data() As String, Optional role As Integer) As employee()
    
    'Error Handling
855:    If gcfHandleErrors Then On Error GoTo UpdateSystemEmployees_Error
856:    PushCallStack "System.UpdateSystemEmployees"
    
    'only update if there are employees working on the system
859:    If Not IsArrayAllocated(data) Then GoTo UpdateSystemEmployees_Exit
    
861:    Dim employees() As employee
862:    Dim i As Integer
        
864:    ReDim employees(UBound(data))
        
866:    For i = 0 To UBound(data)
867:        Set employees(i) = New employee
868:        Call employees(i).PopulateClass(Trim(data(i)), , role)
869:        If employees(i).EmpId = -1 Then employees(i).EmpId = FindGetEmployeeId(employees(i))
870:    Next
    
872:    Select Case role
    Case EmployeePosition.RoshAnafBachir
874:        cListToUpdate.Add "eRoshAnafBachir", eRoshAnafBachir
875:    Case EmployeePosition.RoshAnaf
876:        cListToUpdate.Add "eRoshAnaf", eRoshAnaf
877:    Case EmployeePosition.RoshMador
878:        cListToUpdate.Add "eRoshMador", eRoshMador
879:    Case EmployeePosition.KnowledgeExpert
880:        cListToUpdate.Add "eWorkers", eWorkers
881:    Case Else
882:        cListToUpdate.Add "eRoshAnafBachir", eRoshAnafBachir
883:        cListToUpdate.Add "eRoshAnaf", eRoshAnaf
884:        cListToUpdate.Add "eRoshMador", eRoshMador
885:        cListToUpdate.Add "eWorkers", eWorkers
886:    End Select
    
888:    UpdateSystemEmployees = employees

'Exit Point
UpdateSystemEmployees_Exit:
892:    PopCallStack
893:    On Error GoTo 0
894:    Exit Function

UpdateSystemEmployees_Error:
897:    GlobalErrHandler
898:    Erase employees
899:    UpdateSystemEmployees = employees
900:    Resume UpdateSystemEmployees_Exit
    
902: End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Public Sub UpdateSystemDevLangs(data() As String)
'
'   Sub receives a list of dev langs - an array of strings.
'   Sub has an array for dev langs- sProgrammingLang()
'   Sub puts the dev langs in the arrray
'   Finally, CListToUpdate is being called to flag a new change that needs to go into the
'   DB
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateSystemDevLangs(data() As String)
     
    'Error Handling
918:    If gcfHandleErrors Then On Error GoTo UpdateSystemDevLangs_Error
919:    PushCallStack "System.UpdateSystemDevLangs"
    
921:    If UBound(data) = -1 Then GoTo UpdateSystemDevLangs_Exit
922:    Dim i As Integer
    
924:    ReDim sProgrammingLang(UBound(data))
    
926:    For i = 0 To UBound(data)
927:        sProgrammingLang(i) = Trim(data(i))
928:    Next
    
    'signal an update to the DB for the programming lang
931:    cListToUpdate.Add "sProgrammingLang", sProgrammingLang
    
    'Exit Point
UpdateSystemDevLangs_Exit:
935:    PopCallStack
936:    On Error GoTo 0
937:    Exit Sub

UpdateSystemDevLangs_Error:
940:    GlobalErrHandler
941:    Resume UpdateSystemDevLangs_Exit
    
943: End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Public Sub UpdateDbTypesOfSystem(data() As String, IsOpen As Boolean) UpdateSystemDbTypes
'
'   Sub receives a list of Db Types - an array of strings.
'   Sub has 2 arrays for DB Types:
'   1. sOpenDb - for open DB Types
'   2. sMFDB - For MF DB types
'   Sub puts the Db Type in the arrray, according to it's type (represented by IsOpen boolean)
'   Finally, CListToUpdate is being called to flag a new change that needs to go into the
'   DB
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateDbTypesOfSystem(data() As String, IsOpen As Boolean)
     
    'Error Handling
961:    If gcfHandleErrors Then On Error GoTo UpdateDbTypesOfSystem_Error
962:    PushCallStack "System.UpdateDbTypesOfSystem"
    
964:    If UBound(data) = -1 Then GoTo UpdateDbTypesOfSystem_Exit
965:    Dim result() As String
966:    Dim i As Integer
    
968:    ReDim result(UBound(data))
969:    For i = 0 To UBound(data)
970:        result(i) = Trim(data(i))
971:    Next
     
    'for open systems DB
974:    If IsOpen = True Then
975:        sOpenDB = result
        'signal an update to the DB for the open DB types
977:        cListToUpdate.Add "sOpenDB", sOpenDB
978:    Else
979:        sMFDB = result
        'signal an update to the DB for the MF DB types
981:        cListToUpdate.Add "sMFDB", sMFDB
982:    End If
    
'Exit Point
UpdateDbTypesOfSystem_Exit:
986:    PopCallStack
987:    On Error GoTo 0
988:    Exit Sub

UpdateDbTypesOfSystem_Error:
991:    GlobalErrHandler
992:    Resume UpdateDbTypesOfSystem_Exit
    
994: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Private Function CreateNewSystemRecord(sysName As String) As Integer
'
'   Function receives System Name and inserts a new system record into the DB.
'   Function returns the ID of the system entered.
'   Important! Only call this function when you are sure the
'   the system does not exist in the DB already.
'   Function will not look for the system by name.
'   The function insert a new record to both tblSystems and tblSystemDescription
'   Function returns new system ID
'   On Error, function returns -1
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CreateNewSystemRecord(sysName As String) As Integer
         
    'Error Handling
1013:    If gcfHandleErrors Then On Error GoTo CreateNewSystemRecord_Error
1014:    PushCallStack "System.CreateNewSystemRecord"
    
1016:    Dim NewSystemNum
    
    'Step 1: find out what is the maximum value in the DB
1019:    MyDb.pSQLQuery = "select max(systemId) from tblSystems"
1020:    MyDb.ExecuteSelect
    
1022:    NewSystemNum = CInt(MyDb.pDataArray(0, 0)) + 1
      
    'new record into tblSystemDescription
1025:    MyDb.pSQLQuery = "insert into tblSystemDescription (SystemId) values (" & NewSystemNum & ");"
1026:    MyDb.ExecuteInsert
    
    'new system into tblSystems
1029:    MyDb.pSQLQuery = "insert into tblSystems values (" & NewSystemNum & ",'" & sysName & "'," & "true)"
1030:    MyDb.ExecuteInsert
        
1032:    CreateNewSystemRecord = NewSystemNum

'Exit Point
CreateNewSystemRecord_Exit:
1036:    PopCallStack
1037:    On Error GoTo 0
1038:    Exit Function

CreateNewSystemRecord_Error:
1041:    GlobalErrHandler
1042:    CreateNewSystemRecord = -1
1043:    Resume CreateNewSystemRecord_Exit
    
1045: End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Private Function FindGetEmployeeId(ByRef Emp As Employee) As Employee
'
'   function receives Employee Object as a parameter (ByREf) and looks for employee id
'   in tblemployees. If the Employee was not found by First Name + Last Name, function
'   will create a new employee with that name and assigns a new ID for it.
'   Function returns the id of the employee.
'   In case of an error, function returns -1
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FindGetEmployeeId(ByRef Emp As employee) As Integer

    'Error Handling
1061:    If gcfHandleErrors Then On Error GoTo FindGetEmployeeId_Error
1062:    PushCallStack "System.FindGetEmployeeId"
    
1064:    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "System.FindGetEmployeeId", "DB.ConnectToDB returned FALSE"
     
1067:    MyDb.pSQLQuery = "select employeeid from tblemployees where FirstName ='" & Emp.FirstName & "' and LastName = '" & Emp.LastName & "'"
1068:    MyDb.ExecuteSelect
    
1070:    If IsEmpty(MyDb.pDataArray) Then
        'Todo - find out how to add the position of the employee
1072:        Emp.EmpId = InsertNewEmployee(Emp.FirstName, Emp.LastName, 8)
1073:    Else
1074:        Emp.EmpId = CInt(MyDb.pDataArray(0, 0))
1075:    End If
     
    
    'returned value
1079:    FindGetEmployeeId = Emp.EmpId
 
'Exit Point
FindGetEmployeeId_Exit:
1083:    PopCallStack
1084:    On Error GoTo 0
1085:    Exit Function

FindGetEmployeeId_Error:
1088:    GlobalErrHandler
1089:    FindGetEmployeeId = -1
1090:    Resume FindGetEmployeeId_Exit
    
1092: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Private Function InsertNewEmployee(FirstName As String, LastName As String) As Integer
'
'   Function receives First Name and Last Name and inserts a new employee into the DB.
'   Function returns the ID of the employee entered.
'   Important! Only call this function when you are sure the
'   employee does not exist in the DB already.
'   Function will not look for the Employee.
'   Function returns employee ID.
'   In case of Error  function returns -1
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function InsertNewEmployee(FirstName As String, LastName As String, Position As Integer) As Integer
    
    'Error Handling
1110:    If gcfHandleErrors Then On Error GoTo InsertNewEmployee_Error
1111:    PushCallStack "System.InsertNewEmployee"
    
1113:    Dim NewEmpNum
    
    'Step 1: find out what is the maximum value in the DB
1116:    MyDb.pSQLQuery = "select max(EmployeeId) from tblEmployees"
1117:    MyDb.ExecuteSelect
    
1119:    NewEmpNum = CInt(MyDb.pDataArray(0, 0)) + 1
    
1121:    MyDb.pSQLQuery = "insert into tblemployees values (" & NewEmpNum & ", '" & FirstName & "', '" & LastName & "'," & Position & ")"
1122:    MyDb.ExecuteInsert
    
1124:    InsertNewEmployee = NewEmpNum
 
'Exit Point
InsertNewEmployee_Exit:
1128:    PopCallStack
1129:    On Error GoTo 0
1130:    Exit Function

InsertNewEmployee_Error:
1133:    GlobalErrHandler
1134:    InsertNewEmployee = -1
1135:    Resume InsertNewEmployee_Exit
 
1137: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Private Function FindGetSkillId(strName As String,SkillType as Integer) As Integer
'
'   function looks for a skill in the skills table - tblSkills.
'   if the string is "", function will return -1
'   If function did not find the value in the DB, it will call a sub
'   to insert the new Skill into tblSkills with it's skill type
'   In case of Error, function returns -1
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FindGetSkillId(strName As String, SkillType As Integer) As Integer
    
    'Error Handling
1153:    If gcfHandleErrors Then On Error GoTo FindGetSkillId_Error
1154:    PushCallStack "System.FindGetSkillId"
    
1156:    Dim data, SkillId As Integer
        
1158:    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "System.FindGetSkillId", "DB.ConnectToDB returned FALSE"
    
1161:    If Len(strName) = 0 Then _
        ThrowError CustomError.STRING_IS_EMPTY, "System.FindGetSkillId", "Skill Name cannot be empty at this point"

1164:    MyDb.pSQLQuery = "select SkillId from tblSkills where SkillDescription ='" & Trim(strName) & "'"
1165:    MyDb.ExecuteSelect
    
    'skill was not found
1168:    If IsEmpty(MyDb.pDataArray) Then
1169:        SkillId = -1
1170:    Else
1171:        SkillId = CInt(MyDb.pDataArray(0, 0))
1172:    End If
    
    'insert a new skill into the DB
1175:    If SkillId = -1 Then SkillId = InsertIntoSkills(Trim(strName), SkillType)
    
    'returned value
1178:    FindGetSkillId = SkillId

'Exit Point
FindGetSkillId_Exit:
1182:    PopCallStack
1183:    On Error GoTo 0
1184:    Exit Function

FindGetSkillId_Error:
1187:    GlobalErrHandler
1188:    FindGetSkillId = -1
1189:    Resume FindGetSkillId_Exit
 
1191: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Private Function InsertIntoSkills(strName As String, iType As Integer) As Integer
'
'   Function inserts into skills table a new skill.
'   Function receives Skill Name as String and skill type id (from tblSkillsType)
'   Function assumes skill name does not exist and skill type is valid and verified.
'   function returns the id of the new skill.
'   In case of an Error, function returns -1
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function InsertIntoSkills(strName As String, iType As Integer) As Integer
    
    'Error Handling
1207:    If gcfHandleErrors Then On Error GoTo InsertIntoSkills_Error
1208:    PushCallStack "System.InsertIntoSkills"
    
1210:    Dim NewSkillId
    
1212:    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "System.InsertIntoSkills", "DB.ConnectToDB returned FALSE"
           
1215:    If Len(strName) = 0 Then _
        ThrowError CustomError.STRING_IS_EMPTY, "System.InsertIntoSkills", "skill name cannot be empty at this point"
    
    'Step 1: find out what is the maximum value in the DB
1219:    MyDb.pSQLQuery = "select max(SkillId) from tblSkills"
1220:    MyDb.ExecuteSelect
    
1222:    NewSkillId = CInt(MyDb.pDataArray(0, 0)) + 1
    
    
1225:    MyDb.pSQLQuery = "insert into tblSkills values (" & NewSkillId & ", '" & strName & "'," & iType & ")"
1226:    Call MyDb.ExecuteInsert
       
1228:    InsertIntoSkills = CInt(NewSkillId)
    
'Exit Point
InsertIntoSkills_Exit:
1232:    PopCallStack
    
1234:    On Error GoTo 0
1235:    Exit Function

InsertIntoSkills_Error:
1238:    GlobalErrHandler
1239:    InsertIntoSkills = -1
1240:    Resume InsertIntoSkills_Exit
    
1242: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                               Public Sub FromClassToDb()
'
'   Sub updates the relevant tables with all the information needed, which is specified
'   in the cListToUpdate dictionary. The dictionary itself is updated for every propertyLet
'   or other places in the code, such as eWorkers.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromClassToDb()
    
    'Error Handling
1256:    If gcfHandleErrors Then On Error GoTo FromClassToDb_Error
1257:    PushCallStack "System.FromClassToDb"
    
    'make sure the system id is OK
1260:    If SysId = -1 Then FindGetSystemId (Me.SystemName)
    
    'nothing to update
1263:    If cListToUpdate.count = 0 Then GoTo FromClassToDb_Exit
    
    'check if RoshAnafBachir needs to be updated
1266:    If cListToUpdate.Exists("eRoshAnafBachir") Then
1267:        Call UpdateSystemEmployeeTable(eRoshAnafBachir, EmployeePosition.RoshAnafBachir)
1268:        cListToUpdate.Remove "eRoshAnafBachir"
1269:    End If
    
    'check if RoshAnaf needs to be updated
1272:    If cListToUpdate.Exists("eRoshAnaf") Then
1273:        Call UpdateSystemEmployeeTable(eRoshAnaf, EmployeePosition.RoshAnaf)
1274:        cListToUpdate.Remove "eRoshAnaf"
1275:    End If
    
    'check if RoshMador needs to be updated
1278:    If cListToUpdate.Exists("eRoshMador") Then
1279:        Call UpdateSystemEmployeeTable(eRoshMador, EmployeePosition.RoshMador)
1280:        cListToUpdate.Remove "eRoshMador"
1281:    End If

    
    'check if employees needs to be updated
1285:    If cListToUpdate.Exists("eWorkers") Then
1286:        Call UpdateSystemEmployeeTable(eWorkers, EmployeePosition.KnowledgeExpert)
1287:        cListToUpdate.Remove "eWorkers"
1288:    End If
     
    'check if programming languages needs to be update
1291:    If cListToUpdate.Exists("sProgrammingLang") Then
1292:        Call UpdateSystemDevLang(sProgrammingLang)
1293:        cListToUpdate.Remove "sProgrammingLang"
1294:    End If
    
    'check if openDB needs to be update
1297:    If cListToUpdate.Exists("sOpenDB") Then
1298:        Call UpdateSystemDbTypes(sOpenDB, True)
1299:        cListToUpdate.Remove "sOpenDB"
1300:    End If
    
    'check if MFDB needs to be update
1303:    If cListToUpdate.Exists("sMFDB") Then
1304:        Call UpdateSystemDbTypes(sMFDB, False)
1305:        cListToUpdate.Remove "sMFDB"
1306:    End If
    
    
    'check if Interfaces needs to be update
1310:    If cListToUpdate.Exists("IInterfaces") Then
1311:        Call UpdateSystemInterfaces(IInterfaces)
1312:        cListToUpdate.Remove "IInterfaces"
1313:    End If
    
    'check if Skills needs to be update
1316:    If cListToUpdate.Exists("SSkills") Then
1317:        Call UpdateSystemSkills(SSkills)
1318:        cListToUpdate.Remove "SSkills"
1319:    End If
    
       
1322:    Call UpdateSystemTable(cListToUpdate)
    
    
    'at the end, remove all items
1326:    cListToUpdate.RemoveAll
    

'Exit Point
FromClassToDb_Exit:
1331:    PopCallStack
1332:    On Error GoTo 0
1333:    Exit Sub

FromClassToDb_Error:
1336:    GlobalErrHandler
1337:    Call Me.PermanentDelete
1338:    MsgBox "שגיאה בהעברת נתוני מערכת למסד הנתונים: " & Err.Description, vbCritical, "שגיאה"
1339:    Resume FromClassToDb_Exit

1341: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                               Private Sub UpdateSystemTable(ListToUpdate As Dictionary)
'
'   Sub receives a dictionary with changes to be made
'   and inserts all changes  to tblDescritpion
'   Sub assumes that SysId is well defined.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateSystemTable(ListToUpdate As Dictionary)
    
    'Error Handling
1355:    If gcfHandleErrors Then On Error GoTo UpdateSystemTable_Error
1356:    PushCallStack "System.UpdateSystemTable"
    
    'nothing to update
1359:    If ListToUpdate.count = 0 Then GoTo UpdateSystemTable_Exit
   
1361:    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "System.UpdateSystemTable", "DB.ConnectToDB returned FALSE"
    
1364:    Dim i As Integer
1365:    For i = 0 To ListToUpdate.count - 1
1366:        If IsNumeric(ListToUpdate.Items(i)) Then
1367:            MyDb.pSQLQuery = "update tblSystemDescription set " & ListToUpdate.Keys(i) & " = " & ListToUpdate.Items(i) & " where SystemId=" & Me.SysId & ";"
1368:        Else
            'dim ReadyForInert
1370:            MyDb.pSQLQuery = "update tblSystemDescription set " & ListToUpdate.Keys(i) & " = '" & CStr(ListToUpdate.Items(i)) & "' where SystemId=" & Me.SysId & ";"
1371:        End If
1372:        MyDb.ExecuteUpdate
1373:    Next
    
'Exit Point
UpdateSystemTable_Exit:
1377:    PopCallStack
1378:    On Error GoTo 0
1379:    Exit Sub

UpdateSystemTable_Error:
1382:    GlobalErrHandler
1383:    Resume UpdateSystemTable_Exit
            
1385: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Private Sub UpdateSystemEmployeeTable(EmployeesToUpdate() As Employee)
'
'   Sub receives an array of workers to make sure table is updated
'   and inserts all changes  to tblDescritpion
'   Sub assumes that SysId is well defined.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateSystemEmployeeTable(EmployeesToUpdate() As employee, RoleInSystem As Integer)
    
    'Error Handling
1399:    If gcfHandleErrors Then On Error GoTo UpdateSystemEmployeeTable_Error
1400:    PushCallStack "System.UpdateSystemEmployeeTable"
    
    'nothing to update
1403:    If Not IsArrayAllocated(EmployeesToUpdate) = -1 Then GoTo UpdateSystemEmployeeTable_Exit
    
    'connect to DB
1406:    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "System.UpdateSystemEmployeeTable", "DB.ConnectToDB returned FALSE"
    
1409:    Dim i As Integer
1410:    For i = 0 To UBound(EmployeesToUpdate)
        
        'check if employee is defined in tbSystemEmployee
1413:        MyDb.pSQLQuery = "Select EmployeeId from  tblSystemEmployees where SystemId=" & _
            Me.SysId & " and EmployeeId =" & EmployeesToUpdate(i).EmpId & " and RoleInSystem=" & RoleInSystem & ";"
1415:        MyDb.ExecuteSelect
        
        
        'need to insert the record
1419:        If Not IsArrayAllocated(MyDb.pDataArray) Then
1420:            MyDb.pSQLQuery = "insert into tblSystemEmployees values ( " & _
                Me.SysId & "," & EmployeesToUpdate(i).EmpId & "," & RoleInSystem & ");"
1422:            MyDb.ExecuteInsert
1423:        End If
1424:    Next

'Exit Point
UpdateSystemEmployeeTable_Exit:
1428:    PopCallStack
1429:    On Error GoTo 0
1430:    Exit Sub

UpdateSystemEmployeeTable_Error:
1433:    GlobalErrHandler
1434:    Resume UpdateSystemEmployeeTable_Exit
            
1436: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Private Sub UpdateSystemDevLang(data() As String)
'
'   Sub receives an array of programming languages to make sure table is updated
'   and inserts all changes  to tblSystemDevLang
'   Sub assumes that SysId is well defined.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateSystemDevLang(data() As String)

    'Error Handling
1450:    If gcfHandleErrors Then On Error GoTo UpdateSystemDevLang_Error
1451:    PushCallStack "System.UpdateSystemDevLang"
    
1453:    Dim DevLangId As Integer
    
    'nothing to update
1456:    If UBound(data) = -1 Then GoTo UpdateSystemDevLang_Exit
    
    'connect to DB
1459:    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "System.UpdateSystemDevLang", "DB.ConnectToDB returned FALSE"
    
1462:    Dim i As Integer
1463:    For i = 0 To UBound(data)
        
        'check if dev lang is defined in tbSystemEmployee
1466:        DevLangId = FindGetSkillId(data(i), 8)
1467:        MyDb.pSQLQuery = "Select SystemId from  tblSystemDevLang where SystemId=" & Me.SysId & " and DevLangId =" & DevLangId & ";"
1468:        MyDb.ExecuteSelect
        
        
        'need to insert the record
1472:        If IsEmpty(MyDb.pDataArray) Then
1473:            MyDb.pSQLQuery = "insert into tblSystemDevLang values ( " & Me.SysId & "," & DevLangId & ");"
1474:            MyDb.ExecuteInsert
1475:        End If
1476:    Next

'Exit Point
UpdateSystemDevLang_Exit:
1480:    PopCallStack
1481:    On Error GoTo 0
1482:    Exit Sub

UpdateSystemDevLang_Error:
1485:    GlobalErrHandler
1486:    Resume UpdateSystemDevLang_Exit
        
1488: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Public Sub UpdateSystemDbTypes(data() As String, isForOpen As Boolean)
'
'   Sub receives an array of Db types to make sure table is updated
'   and inserts all changes  to tblSystemDBTypes
'   Sub assumes that SysId is well defined.
'   Sub takes care of both open DB types and MF DB types
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateSystemDbTypes(data() As String, isForOpen As Boolean)

    'Error Handling
1503:    If gcfHandleErrors Then On Error GoTo UpdateSystemDbTypes_Error
1504:    PushCallStack "System.UpdateSystemDbTypes"
    
1506:    Dim DBTypeId As Integer
    
    'nothing to update
1509:    If UBound(data) = -1 Then GoTo UpdateSystemDbTypes_Exit
    
    'connect to DB
1512:    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "System.UpdateSystemDevLang", "DB.ConnectToDB returned FALSE"
    
1515:    Dim i As Integer
1516:    For i = 0 To UBound(data)
        
1518:        If UCase(data(i)) <> "N/A" Then
            'check if DBType is defined in tbSystemEmployee
1520:            DBTypeId = FindGetSkillId(data(i), 5)
1521:            MyDb.pSQLQuery = "Select SystemId from  tblSystemDbTypes where SystemId=" & Me.SysId & " and DbTypeId =" & DBTypeId & ";"
1522:            MyDb.ExecuteSelect
            
            
            'need to insert the record
1526:            If IsEmpty(MyDb.pDataArray) Then
1527:                MyDb.pSQLQuery = "insert into tblSystemDbTypes values ( " & Me.SysId & "," & DBTypeId & "," & isForOpen & ");"
1528:                MyDb.ExecuteInsert
1529:            End If
1530:        End If
1531:    Next


'Exit Point
UpdateSystemDbTypes_Exit:
1536:    PopCallStack
1537:    On Error GoTo 0
1538:    Exit Sub

UpdateSystemDbTypes_Error:
1541:    GlobalErrHandler
1542:    Resume UpdateSystemDbTypes_Exit
            
1544: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Public Function GetDbTypesOfSystem(isForOpen As Boolean) As String
'
'   Function Builds a string representing the DB types of the system. Function receives
'   a boolean representing am I relating to Open DB's or MF DB's
'   In case of Error, function returns ""
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetDbTypesOfSystem(isForOpen As Boolean) As String

    'Error Handling
1557:    If gcfHandleErrors Then On Error GoTo GetDbTypesOfSystem_Error
1558:    PushCallStack "System.GetDbTypesOfSystem"
    
1560:    Dim strTmpResult As String, Source() As String, i As Integer
1561:    strTmpResult = ""
    
1563:    If isForOpen = True Then
1564:        Source = sOpenDB
1565:    Else
1566:        Source = sMFDB
1567:    End If
    
1569:    If IsEmpty(Source) Then
1570:        GetDbTypesOfSystem = ""
1571:        GoTo GetDbTypesOfSystem_Exit
1572:    End If
    
1574:    For i = 0 To UBound(Source)
1575:        strTmpResult = strTmpResult & Source(i) & ", "
1576:    Next
   
1578:    GetDbTypesOfSystem = Left(strTmpResult, Len(strTmpResult) - 2)

'Exit Point
GetDbTypesOfSystem_Exit:
1582:    PopCallStack
1583:    On Error GoTo 0
1584:    Exit Function

GetDbTypesOfSystem_Error:
1587:    GlobalErrHandler
1588:    GetDbTypesOfSystem = ""
1589:    Resume GetDbTypesOfSystem_Exit
1590: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'           Function getGeneralInformationAsStringArray() As String()
'
'   Function builds a string of the general information from the
'   data it stores of the system.
'   Note! this function assumes that the class is initialized and that
'   there is valid data in the class members.
'   TODO: create a validation to check if the class is truely initialized.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function getGeneralInformationAsStringArray() As String()
    
    'Error Handling
1605:    If gcfHandleErrors Then On Error GoTo getGeneralInformationAsStringArray_Error
1606:    PushCallStack "System.getGeneralInformationAsStringArray"
     
1608:    Dim i As Integer, data(0 To 33, 1 To 1) As String
     
1610:    data(0, 1) = SystemName                                                              'SystemName
1611:    data(1, 1) = SubSystemName                                                           '"SubSystemName"
1612:    data(2, 1) = SystemDescription                                                       '"SystemDescription"
1613:    data(3, 1) = KnowledgeConsumer                                                       'KnowledgeConsumers
1614:    data(4, 1) = GetEmployeesListAsString(eRoshAnafBachir)                               'RoshAnafBachir
1615:    data(5, 1) = GetEmployeesListAsString(eRoshAnaf)                                     '"RoshAnaf"
1616:    data(6, 1) = GetEmployeesListAsString(eRoshMador)                                    '"RoshMador"
1617:    data(7, 1) = GetEmployeesListAsString(eWorkers)                                                         'system experts on A9
    
    
    'business or infratructure. business = 0, infrastructure = 1                                'IsInfrastructure
1621:    Dim strTmpResult
1622:    If IsInfrastructure = 0 Then
1623:        strTmpResult = "עסקית"
1624:    ElseIf IsInfrastructure = 1 Then
1625:        strTmpResult = "תשתיתית"
1626:    Else
1627:        strTmpResult = "עסקית ותשתיתית"
1628:    End If
1629:    data(8, 1) = strTmpResult
    
    
1632:    data(9, 1) = BizDev                                                                  'bizEnv
    
    'B12 / data(10) - No need. Is included in data(9) - bizdev
    
1636:    data(11, 1) = DevEnv                                                                 'DevEnv
1637:    data(12, 1) = TechEnv                                                                'TechEnv
    
    'B15- isCore                                                                                'isCore
1640:    strTmpResult = ""
1641:    If IsCore = False Then
1642:        strTmpResult = "לא"
1643:    Else
1644:        strTmpResult = "כן"
1645:    End If
1646:    data(13, 1) = strTmpResult
    
    'B16- programming languages
1649:    data(14, 1) = GetProgrammingLangsAsString()                                                   'Programming languages
    
    'B17- isWebBased
1652:    strTmpResult = ""
1653:    If IsWebBased = False Then
1654:        strTmpResult = "לא"
1655:    Else
1656:        strTmpResult = "כן"
1657:    End If
1658:    data(15, 1) = strTmpResult
    
    
1661:    data(16, 1) = GetDbTypesOfSystem(True)                                               'B18 Open DB List
1662:    data(17, 1) = GetDbTypesOfSystem(False)                                              'B19 MainFrame DB
1663:    data(18, 1) = NumOfInterfaces                                                        'B20 - NumOfInterfaces
1664:    data(19, 1) = NumOfCriticalInterfaces                                                'B21- NumOfCriticalInterfaces
1665:    data(20, 1) = SystemRisks                                                            'B22- SystemRisks
1666:    data(21, 1) = PreservationTopics                                                     'B23- PreservationTopics
1667:    data(22, 1) = PreservationSuggestions                                                'B24 - PreservationSuggestions

    'B25 - IsPlannedToClose
1670:    strTmpResult = ""
1671:    If IsPlannedToClose = 0 Then
1672:        strTmpResult = "לא"
1673:    ElseIf IsPlannedToClose = 1 Then
1674:        strTmpResult = "כן"
1675:    Else
1676:        strTmpResult = "חלקי"
1677:    End If
1678:    data(23, 1) = strTmpResult
    
    'B26 - IsChangeManagement
1681:    strTmpResult = ""
1682:    If IsChangeManagement = 0 Then
1683:        strTmpResult = "לא"
1684:    ElseIf IsChangeManagement = 1 Then
1685:        strTmpResult = "כן"
1686:    Else
1687:        strTmpResult = "חלקי"
1688:    End If
1689:    data(24, 1) = strTmpResult
    
    'B27- IsManagementBrief
1692:    strTmpResult = ""
1693:    If IsManagementBrief = 0 Then
1694:        strTmpResult = "לא"
1695:    ElseIf IsManagementBrief = 1 Then
1696:        strTmpResult = "כן"
1697:    Else
1698:        strTmpResult = "חלקי"
1699:    End If
1700:    data(25, 1) = strTmpResult

    'B28 - IsApplicableDocuments
1703:    strTmpResult = ""
1704:    If IsApplicableDocuments = 0 Then
1705:        strTmpResult = "לא"
1706:    ElseIf IsApplicableDocuments = 1 Then
1707:        strTmpResult = "כן"
1708:    Else
1709:        strTmpResult = "חלקי"
1710:    End If
1711:    data(26, 1) = strTmpResult
    
    'B29-IsArchitectureDoc
1714:    strTmpResult = ""
1715:    If IsArchitectureDoc = 0 Then
1716:        strTmpResult = "לא"
1717:    ElseIf IsArchitectureDoc = 1 Then
1718:        strTmpResult = "כן"
1719:    Else
1720:        strTmpResult = "חלקי"
1721:    End If
1722:    data(27, 1) = strTmpResult
    
    'B30- IsShobDoc
1725:    strTmpResult = ""
1726:    If IsShobDoc = 0 Then
1727:        strTmpResult = "לא"
1728:    ElseIf IsShobDoc = 1 Then
1729:        strTmpResult = "כן"
1730:    Else
1731:        strTmpResult = "חלקי"
1732:    End If
1733:    data(28, 1) = strTmpResult
    
    'B31- IsTesting
1736:    strTmpResult = ""
1737:    If IsTesting = 0 Then
1738:        strTmpResult = "לא"
1739:    ElseIf IsTesting = 1 Then
1740:        strTmpResult = "כן"
1741:    Else
1742:        strTmpResult = "חלקי"
1743:    End If
1744:    data(29, 1) = strTmpResult
    
    'B32-IsSupportAndOp
1747:    strTmpResult = ""
1748:    If IsSupportAndOp = 0 Then
1749:        strTmpResult = "לא"
1750:    ElseIf IsSupportAndOp = 1 Then
1751:        strTmpResult = "כן"
1752:    Else
1753:        strTmpResult = "חלקי"
1754:    End If
1755:    data(30, 1) = strTmpResult
    
    'B33-QaPassFail
1758:    strTmpResult = ""
1759:    If QaPassFail = 0 Then
1760:        strTmpResult = "לא עובר"
1761:    ElseIf QaPassFail = 1 Then
1762:        strTmpResult = "עובר"
1763:    Else
1764:        strTmpResult = "עובר חלקית"
1765:    End If
1766:    data(31, 1) = strTmpResult
    
1768:    data(32, 1) = Comments
    
1770:    getGeneralInformationAsStringArray = data

   
'Exit Point
getGeneralInformationAsStringArray_Exit:
1775:    PopCallStack
1776:    On Error GoTo 0
1777:    Exit Function

getGeneralInformationAsStringArray_Error:
1780:    GlobalErrHandler
1781:    MsgBox "שגיאה בהצגת נתוני מערכת: " & Err.Description, vbCritical, "שגיאה"
1782:    Erase data
1783:    getGeneralInformationAsStringArray = data
1784:    Resume getGeneralInformationAsStringArray_Exit

1786: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Public Function GetInterfaceAsStringArray() As String()
'
'   Function Builds a string array representing the interfaces of the system.
'   If there are no interfaces, function returns an empty array
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetInterfaceAsStringArray() As String()

    'Error Handling
1798:    If gcfHandleErrors Then On Error GoTo GetInterfaceAsStringArray_Error
1799:    PushCallStack "System.GetInterfaceAsStringArray"
    
1801:    Dim Interfaces() As String, i As Integer
    
    'make sure there is at least one iterface
1804:    If Not IInterfaces(0) Is Nothing Then
        
1806:        ReDim Interfaces(UBound(IInterfaces), 3)
1807:        For i = 0 To UBound(Interfaces)
1808:            Interfaces(i, 0) = IInterfaces(i).InterfaceName
1809:            Interfaces(i, 1) = IInterfaces(i).InterfaceCategory
1810:            Interfaces(i, 2) = IInterfaces(i).InterfaceType
1811:            Interfaces(i, 3) = IInterfaces(i).KnowledgeLevel
1812:        Next
1813:    End If
    
1815:    GetInterfaceAsStringArray = Interfaces

'Exit Point
GetInterfaceAsStringArray_Exit:
1819:    PopCallStack
1820:    On Error GoTo 0
1821:    Exit Function


GetInterfaceAsStringArray_Error:
1825:    GlobalErrHandler
1826:    Erase Interfaces
1827:    GetInterfaceAsStringArray = Interfaces
1828:    Resume GetInterfaceAsStringArray_Exit
 
 
1831: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Public Function GetSkillsAsStringArray() As String()
'
'   Function Builds a string array representing the Skills of the system.
'   If there are no Skills, function returns an empty array
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetSkillsAsStringArray() As String()

    'Error Handling
1843:    If gcfHandleErrors Then On Error GoTo GetSkillsAsStringArray_Error
1844:    PushCallStack "System.GetSkillsAsStringArray"
    
1846:    Dim Skills() As String, i As Integer
    
    'make sure there is at least one skill to show
1849:    If Not SSkills(0) Is Nothing Then
        
1851:        ReDim Skills(UBound(SSkills), 2)
1852:        For i = 0 To UBound(Skills)
1853:            Skills(i, 0) = SSkills(i).SkillName
1854:            Skills(i, 1) = SSkills(i).SkillType
1855:            Skills(i, 2) = SSkills(i).KnowledgeLevel
1856:        Next

1858:    End If
    
1860:    GetSkillsAsStringArray = Skills
    
'Exit Point
GetSkillsAsStringArray_Exit:
1864:    PopCallStack
1865:    On Error GoTo 0
1866:    Exit Function

GetSkillsAsStringArray_Error:
1869:    GlobalErrHandler
1870:    Erase Skills
1871:    GetSkillsAsStringArray = Skills
1872:    Resume GetSkillsAsStringArray_Exit
    
1874: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Public Sub BuildInterfacesArray(data() As Variant)
'
'   Sub receives an 2 dimension array representing the interfaces of the system:
'   1. Interface Name - data(i,1)
'   2. Interface Category - data(i,2)
'   3. Interface Type- data(i,3)
'   4. Interface Knowledge Level - data(i,4)
'
'   sub calls the private method AddInterfaceToClass to add a new interface to
'   the interface array.
'
'   Finally, CListToUpdate is being called to flag a new change that needs to go into the
'   DB
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BuildInterfacesArray(data() As Variant)
            
            
    'Error Handling
1897:    If gcfHandleErrors Then On Error GoTo BuildInterfacesArray_Error
1898:    PushCallStack "System.BuildInterfacesArray"
    
1900:    If UBound(data) = 0 Then GoTo BuildInterfacesArray_Exit
1901:    Dim i As Integer
    
1903:    For i = 1 To UBound(data)
1904:        Call AddInterfaceToClass(PrepareStrToDB(CStr(data(i, 1))), CStr(data(i, 2)), CStr(data(i, 3)), CStr(data(i, 4)))
1905:    Next
    
1907:    cListToUpdate.Add "IInterfaces", IInterfaces
    
'Exit Point
BuildInterfacesArray_Exit:
1911:    PopCallStack
1912:    On Error GoTo 0
1913:    Exit Sub

BuildInterfacesArray_Error:
1916:    GlobalErrHandler
1917:    Resume BuildInterfacesArray_Exit

1919: End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Public Sub BuildSkillsArray(data() As Variant)
'
'   Sub receives an 2 dimension array representing the skills of the system:
'   1. Skill Name - data(i,1)
'   2. Skill Type - data(i,2)
'   3. Skill Knowledge Level - data(i,3)
'
'   sub calls the private method AddSkillToClass to add a new skill to
'   the skills array.
'
'   Finally, CListToUpdate is being called to flag a new change that needs to go into the
'   DB
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BuildSkillsArray(data() As Variant)
    
    'Error Handling
1941:    If gcfHandleErrors Then On Error GoTo BuildSkillsArray_Error
1942:    PushCallStack "System.BuildSkillsArray"
    
1944:    If UBound(data) = 1 Then GoTo BuildSkillsArray_Exit
1945:    Dim i As Integer
    
1947:    For i = 1 To UBound(data)
1948:        Call AddSkillToClass(PrepareStrToDB(CStr(data(i, 1))), CStr(data(i, 2)), -1, CStr(data(i, 3)))
1949:    Next
    
1951:    cListToUpdate.Add "SSkills", SSkills
    
'Exit Point
BuildSkillsArray_Exit:
1955:    PopCallStack
1956:    On Error GoTo 0
1957:    Exit Sub

BuildSkillsArray_Error:
1960:    GlobalErrHandler
1961:    Resume BuildSkillsArray_Exit

1963: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Public Function GetEmployeesListAsString(Employees() As Employee) As String
'
'   Function receives an array of employees (coule be System experts,
'   could be Rosh Anaf etc. and builds a string, representing the names of the Employees
'
'   Function returns the string.
'   In case of an Error - function returns ""
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetEmployeesListAsString(employees() As employee) As String
    
    'Error Handling
1979:    If gcfHandleErrors Then On Error GoTo GetEmployeesListAsString_Error
1980:    PushCallStack "System.GetEmployeesListAsString"
    
    'in case there are no known employees
1983:    If Not IsArrayAllocated(employees) Or employees(0) Is Nothing Then
1984:        GetEmployeesListAsString = ""
1985:        GoTo GetEmployeesListAsString_Exit
1986:    End If
    
1988:    Dim i As Integer, result As String
    
1990:    For i = 0 To UBound(employees)
1991:        result = result & employees(i).FirstName & " " & employees(i).LastName & ","
1992:    Next
    
    'sub string without the last comma
1995:    GetEmployeesListAsString = Left(result, Len(result) - 1)

'Exit Point
GetEmployeesListAsString_Exit:
1999:    PopCallStack
2000:    On Error GoTo 0
2001:    Exit Function

GetEmployeesListAsString_Error:
2004:    GlobalErrHandler
2005:    GetEmployeesListAsString = ""
2006:    Resume GetEmployeesListAsString_Exit

2008: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Public Function GetProgrammingLangsAsString() As String
'
'   Function builds a string, representing the names of the programming languages from
'   the system.
'
'   Function returns the string.
'   In cse of an Error, function returns ""
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetProgrammingLangsAsString() As String
        
    'Error Handling
2024:    If gcfHandleErrors Then On Error GoTo GetProgrammingLangsAsString_Error
2025:    PushCallStack "System.GetProgrammingLangsAsString"
    
    'in case there are no programming languages
2028:    If sProgrammingLang(0) = "" Then GetProgrammingLangsAsString = ""
    
2030:    Dim i As Integer, result As String
    
2032:    For i = 0 To UBound(sProgrammingLang)
2033:         result = result + sProgrammingLang(i) + ", "
2034:    Next
    
    'sub string without the last comma
2037:    GetProgrammingLangsAsString = Left(result, Len(result) - 2)

'Exit Point
GetProgrammingLangsAsString_Exit:
2041:    PopCallStack
2042:    On Error GoTo 0
2043:    Exit Function

GetProgrammingLangsAsString_Error:
2046:    GlobalErrHandler
2047:    GetProgrammingLangsAsString = ""
2048:    Resume GetProgrammingLangsAsString_Exit
    
2050: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Private Sub UpdateSystemInterfaces(data() As Interface)
'
'   Sub receives an array of Interfaces to make sure table is updated
'   and inserts all changes  to tblSystemInterfaces
'   Sub assumes that SysId is well defined.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateSystemInterfaces(data() As Interface)

    'Error Handling
2064:    If gcfHandleErrors Then On Error GoTo UpdateSystemInterfaces_Error
2065:    PushCallStack "System.UpdateSystemInterfaces"
         
    'nothing to update
2068:    If UBound(data) = -1 Then GoTo UpdateSystemInterfaces_Exit
    
    'connect to DB
2071:    Call MyDb.ConnectToDB
    
2073:    Dim iInterfaceId As Integer, iKnowledgeLevel As Integer, iInterfaceTypeId As Integer
    
2075:    Dim i As Integer
2076:    For i = 0 To UBound(data)
2077:        If Not data(i).IsClassEmpty Then
            'check if interface is defined in tblInterfaces
2079:            iInterfaceId = FindGetInterfaceId(data(i))
            
2081:            MyDb.pSQLQuery = "Select SystemId from  tblSystemInterfaces where SystemId=" & Me.SysId & " and InterfaceId =" & iInterfaceId & ";"
2082:            MyDb.ExecuteSelect
            
            
            'need to insert the record
2086:            If IsEmpty(MyDb.pDataArray) Then
2087:                iKnowledgeLevel = FindGetKnowLedgeLevelId(data(i).KnowledgeLevel)
2088:                iInterfaceTypeId = FindGetInterfaceTypeId(data(i).InterfaceType)
2089:                MyDb.pSQLQuery = "insert into tblSystemInterfaces values ( " & Me.SysId & "," & iInterfaceId & "," & iKnowledgeLevel & "," & iInterfaceTypeId & ");"
2090:                MyDb.ExecuteInsert
2091:            End If
2092:        End If
2093:    Next
        
'Exit Point
UpdateSystemInterfaces_Exit:
2097:    PopCallStack
    
2099:    On Error GoTo 0
2100:    Exit Sub

UpdateSystemInterfaces_Error:
2103:    GlobalErrHandler
2104:    Resume UpdateSystemInterfaces_Exit
    
2106: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Private Sub FindGetInterfaceId(data As Interface)
'
'   Sub receives an array of Interfaces to make sure table is updated
'   and inserts all changes  to tblSystemInterfaces
'   Sub assumes that SysId is well defined.
'   On Error function returns -1
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FindGetInterfaceId(data As Interface)
    
    'Error Handling
2121:    If gcfHandleErrors Then On Error GoTo FindGetInterfaceId_Error
2122:    PushCallStack "System.FindGetInterfaceId"
    
    'nothing to update
2125:    If data.IsClassEmpty = True Then
2126:        FindGetInterfaceId = -1
2127:        GoTo FindGetInterfaceId_Exit
2128:    End If
        
    
    'connect to DB
2132:    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "System.FindGetInterfaceId", "DB.ConnectToDB returned FALSE"
    
2135:    Dim iInterfaceId As Integer
    
2137:    Dim i As Integer
         
    'check if interface is defined in tblInterfacesInterfaces
2140:    MyDb.pSQLQuery = "Select InterfaceId from  tblInterfaces where InterfaceDescription='" & data.InterfaceName & "';"
2141:    MyDb.ExecuteSelect
    
    'need to insert the record
2144:    If IsEmpty(MyDb.pDataArray) Then
2145:        Dim iInterfaceCategory As Integer, iInterfaceType As Integer
        
2147:        iInterfaceCategory = FindGetInterfaceCategoryId(data.InterfaceCategory)
2148:        iInterfaceId = InsertNewInterface(data.InterfaceName, iInterfaceCategory)
        
2150:        FindGetInterfaceId = iInterfaceId
2151:    Else
2152:        FindGetInterfaceId = CInt(MyDb.pDataArray(0, 0))
2153:    End If
    
'Exit Point
FindGetInterfaceId_Exit:
2157:    PopCallStack
2158:    On Error GoTo 0
2159:    Exit Function

FindGetInterfaceId_Error:
2162:    GlobalErrHandler
2163:    FindGetInterfaceId = -1
2164:    Resume FindGetInterfaceId_Exit
2165: End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Private Function FindGetInterfaceCategoryId(InterfaceCategory As String) As Integer
'
'   function receives interface category name and looks for categoryId
'   in tblInterfaceCategory. If the category was not found by Description, function
'   will create a new category with that name and assigns a new ID for it.
'   Function returns the id of the new category.
'   In case of an Error, function returns -1
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FindGetInterfaceCategoryId(InterfaceCategory As String) As Integer
    
    'Error Handling
2181:    If gcfHandleErrors Then On Error GoTo FindGetInterfaceCategoryId_Error
2182:    PushCallStack "System.FindGetInterfaceCategoryId"
    
2184:    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "System.FindGetInterfaceCategoryId", "DB.ConnectToDB returned FALSE"
        
2187:    If InterfaceCategory = "" Then
2188:        MyDb.pSQLQuery = "select InterfaceCategoryId from tblInterfaceCategory where InterfaceCategoryDescription is null;"
2189:    Else
2190:        MyDb.pSQLQuery = "select InterfaceCategoryId from tblInterfaceCategory where InterfaceCategoryDescription ='" & InterfaceCategory & "';"
2191:    End If
     
2193:    MyDb.ExecuteSelect
    
2195:    If IsEmpty(MyDb.pDataArray) Then
2196:        FindGetInterfaceCategoryId = InsertNewInterfaceCategory(InterfaceCategory)
2197:    Else
2198:        FindGetInterfaceCategoryId = CInt(MyDb.pDataArray(0, 0))
2199:    End If
     
    
    
'Exit Point
FindGetInterfaceCategoryId_Exit:
2205:    PopCallStack
2206:    On Error GoTo 0
2207:    Exit Function

FindGetInterfaceCategoryId_Error:
2210:    GlobalErrHandler
2211:    FindGetInterfaceCategoryId = -1
2212:    Resume FindGetInterfaceCategoryId_Exit
    
2214: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Private Function InsertNewInterfaceCategory(Name As String) As Integer
'
'   Function receives Category Description and inserts a new Category into the DB.
'   Function returns the ID of the Category entered.
'   Important! Only call this function when you are sure the
'   Category does not exist in the DB already.
'   Function will not look for the Category.
'   Function returns CategoryID
'   In case of an Error, function returns -1
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function InsertNewInterfaceCategory(Name As String) As Integer
    
    'Error Handling
2232:    If gcfHandleErrors Then On Error GoTo InsertNewInterfaceCategory_Error
2233:    PushCallStack "System.InsertNewInterfaceCategory"
    
2235:    Dim NewCategoryNum As Integer
    
    'Step 1: find out what is the maximum value in the DB
2238:    MyDb.pSQLQuery = "select max(InterfaceCategoryId) from tblInterfaceCategory"
2239:    MyDb.ExecuteSelect
    
2241:    NewCategoryNum = CInt(MyDb.pDataArray(0, 0)) + 1
    
    
2244:    MyDb.pSQLQuery = "insert into tblInterfaceCategory values (" & NewCategoryNum & ", '" & Name & "');"
2245:    MyDb.ExecuteInsert
    
2247:    InsertNewInterfaceCategory = NewCategoryNum
    
'Exit Point
InsertNewInterfaceCategory_Exit:
2251:    PopCallStack
2252:    On Error GoTo 0
2253:    Exit Function

InsertNewInterfaceCategory_Error:
2256:    GlobalErrHandler
2257:    InsertNewInterfaceCategory = -1
2258:    Resume InsertNewInterfaceCategory_Exit
2259: End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Private Function FindGetInterfaceTypeId(InterfaceType As String) As Integer
'
'   function receives interface Type name and looks for Type Id
'   in tblInterfaceType. If the Type was not found by Description, function
'   will create a new Type with that name and assigns a new ID for it.
'   Function returns the id of the new Type.
'   In case of Error, function returns -1
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FindGetInterfaceTypeId(InterfaceType As String) As Integer

    'Error Handling
2275:    If gcfHandleErrors Then On Error GoTo FindGetInterfaceTypeId_Error
2276:    PushCallStack "System.FindGetInterfaceTypeId"
    
2278:    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "System.FindGetInterfaceTypeId", "DB.ConnectToDB returned FALSE"
     
    'case of null
2282:    If InterfaceType = "" Then
2283:        MyDb.pSQLQuery = "select InterfaceTypeId from tblInterfaceType where InterfaceTypeDescription is null;"
2284:    Else
2285:        MyDb.pSQLQuery = "select InterfaceTypeId from tblInterfaceType where InterfaceTypeDescription ='" & InterfaceType & "';"
2286:    End If
     
    
2289:    MyDb.ExecuteSelect
    
2291:    If IsEmpty(MyDb.pDataArray) Then
    
2293:        FindGetInterfaceTypeId = InsertNewInterfaceType(InterfaceType)
2294:    Else
2295:        FindGetInterfaceTypeId = CInt(MyDb.pDataArray(0, 0))
2296:    End If
     
    
    
'Exit Point
FindGetInterfaceTypeId_Exit:
2302:    PopCallStack
2303:    On Error GoTo 0
2304:    Exit Function

FindGetInterfaceTypeId_Error:
2307:    GlobalErrHandler
2308:    FindGetInterfaceTypeId = -1
2309:    Resume FindGetInterfaceTypeId_Exit
    
2311: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Private Function InsertNewInterfaceType(Name As String) As Integer
'
'   Function receives Type Description and inserts a new Type into the DB.
'   Function returns the ID of the Type entered.
'   Important! Only call this function when you are sure the
'   type does not exist in the DB already.
'   Function will not look for the type.
'   Function returns TypeId
'   In case of an Error, function returns -1
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function InsertNewInterfaceType(Name As String) As Integer
    
    'Error Handling
2329:    If gcfHandleErrors Then On Error GoTo InsertNewInterfaceType_Error
2330:    PushCallStack "System.InsertNewInterfaceType"
    
2332:    Dim NewTypeNum As Integer
    
    'Step 1: find out what is the maximum value in the DB
2335:    MyDb.pSQLQuery = "select max(InterfaceTypeId) from tblInterfaceType"
2336:    MyDb.ExecuteSelect
    
2338:    NewTypeNum = CInt(MyDb.pDataArray(0, 0)) + 1
        
2340:    MyDb.pSQLQuery = "insert into tblInterfaceType values (" & NewTypeNum & ", '" & Name & "');"
2341:    MyDb.ExecuteInsert
    
2343:    InsertNewInterfaceType = NewTypeNum
    
'Exit Point
InsertNewInterfaceType_Exit:
2347:    PopCallStack
2348:    On Error GoTo 0
2349:    Exit Function

InsertNewInterfaceType_Error:
2352:    GlobalErrHandler
2353:    InsertNewInterfaceType = -1
2354:    Resume InsertNewInterfaceType_Exit
2355: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Private Function InsertNewInterface(Description As String, CategoryId As Integer) As Integer
'
'   Function receives Interface information:
'   1.Description
'   2. CategoryId
'   and inserts a new Type into the DB.
'   Function returns the ID of the interface entered.
'   Important! Only call this function when you are sure the
'   interface does not exist in the DB already.
'   Function will not look for the interface.
'   Function returns InterfaceId
'   In case of an Error, function -1
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function InsertNewInterface(Description As String, CategoryId As Integer) As Integer
    
    'Error Handling
2376:    If gcfHandleErrors Then On Error GoTo InsertNewInterface_Error
2377:    PushCallStack "System.InsertNewInterface"
    
2379:    Dim NewInterfaceNum As Integer
    
    'Step 1: find out what is the maximum value in the DB
2382:    MyDb.pSQLQuery = "Select max(InterfaceId) from  tblInterfaces;"
2383:    MyDb.ExecuteSelect
    
2385:    If IsNull((MyDb.pDataArray(0, 0))) Then
2386:        NewInterfaceNum = 1
2387:    Else
2388:        NewInterfaceNum = CInt(MyDb.pDataArray(0, 0)) + 1
2389:    End If
    
    
2392:    MyDb.pSQLQuery = "insert into tblInterfaces values ( " & NewInterfaceNum & ",'" & Description & "'," & CategoryId & ");"
2393:    MyDb.ExecuteInsert
    
2395:    InsertNewInterface = NewInterfaceNum
    
'Exit Point
InsertNewInterface_Exit:
2399:    PopCallStack
2400:    On Error GoTo 0
2401:    Exit Function

InsertNewInterface_Error:
2404:    GlobalErrHandler
2405:    InsertNewInterface = -1
2406:    Resume InsertNewInterface_Exit
    
2408: End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Private Function FindGetKnowLedgeLevelId(KnowledgeLevel As String) As Integer
'
'   function receives interface Knowledge level name and looks for level Id
'   in tblInterfaceKnowledgeLevel. If the Type was not found by Description, function
'   will create a new Knowledge Level with that Description and assigns a new ID for it.
'   Function returns the id of the new Knowledge Level.
'   In case of Error, function returns -1
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FindGetKnowLedgeLevelId(KnowledgeLevel As String) As Integer

    'Error Handling
2424:    If gcfHandleErrors Then On Error GoTo FindGetKnowLedgeLevelId_Error
2425:    PushCallStack "System.FindGetKnowLedgeLevelId"
    
2427:    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "System.FindGetKnowLedgeLevelId", "DB.ConnectToDB returned FALSE"
        
2430:    If KnowledgeLevel = "" Then
2431:        MyDb.pSQLQuery = "select KnoledgeLevelId from tblInterfaceKnowledgeLevel where KnoledgeLevelDescription is null;"
2432:    Else
2433:        MyDb.pSQLQuery = "select KnoledgeLevelId from tblInterfaceKnowledgeLevel where KnoledgeLevelDescription ='" & KnowledgeLevel & "';"
2434:    End If
         
2436:    MyDb.ExecuteSelect
    
2438:    If IsEmpty(MyDb.pDataArray) Then
2439:        FindGetKnowLedgeLevelId = InsertNewInterfaceKnowledgeLevel(KnowledgeLevel)
2440:    Else
2441:        FindGetKnowLedgeLevelId = CInt(MyDb.pDataArray(0, 0))
2442:    End If
     
    
 
'Exit Point
FindGetKnowLedgeLevelId_Exit:
2448:    PopCallStack
2449:    On Error GoTo 0
2450:    Exit Function

FindGetKnowLedgeLevelId_Error:
2453:    GlobalErrHandler
2454:    FindGetKnowLedgeLevelId = -1
2455:    Resume FindGetKnowLedgeLevelId_Exit
    
2457: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Private Function InsertNewInterfaceKnowledgeLevel(Name As String) As Integer
'
'   Function receives Knowledge level Description and inserts a new Type into the DB.
'   Function returns the ID of the Level entered.
'   Important! Only call this function when you are sure the
'   level does not exist in the DB already.
'   Function will not look for the level.
'   Function returns LevelId
'   In case of an Error
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function InsertNewInterfaceKnowledgeLevel(Name As String) As Integer
    
    'Error Handling
2475:    If gcfHandleErrors Then On Error GoTo InsertNewInterfaceKnowledgeLevel_Error
2476:    PushCallStack "System.InsertNewInterfaceKnowledgeLevel"
    
2478:    Dim NewLevelNum As Integer
    
    'Step 1: find out what is the maximum value in the DB
2481:    MyDb.pSQLQuery = "select max(KnoledgeLevelId) from tblInterfaceKnowledgeLevel"
2482:    MyDb.ExecuteSelect
    
2484:    NewLevelNum = CInt(MyDb.pDataArray(0, 0)) + 1
       
2486:    MyDb.pSQLQuery = "insert into tblInterfaceKnowledgeLevel values (" & NewLevelNum & ", '" & Name & "');"
2487:    MyDb.ExecuteInsert
    
2489:    InsertNewInterfaceKnowledgeLevel = NewLevelNum
    
'Exit Point
InsertNewInterfaceKnowledgeLevel_Exit:
2493:    PopCallStack
    
2495:    On Error GoTo 0
2496:    Exit Function

InsertNewInterfaceKnowledgeLevel_Error:
2499:    GlobalErrHandler
2500:    InsertNewInterfaceKnowledgeLevel = -1
2501:    Resume InsertNewInterfaceKnowledgeLevel_Exit
    
2503: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Private Sub UpdateSystemSkills(data() As Skill)
'
'   Sub receives an array of Skills to make sure table is updated
'   and inserts all changes  to tblSystemSkills
'   Sub assumes that SysId is well defined.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateSystemSkills(data() As Skill)

    'Error Handling
2517:    If gcfHandleErrors Then On Error GoTo UpdateSystemSkills_Error
2518:    PushCallStack "System.UpdateSystemSkills"
         
    'nothing to update
2521:    If UBound(data) = -1 Then GoTo UpdateSystemSkills_Exit
    
    'connect to DB
2524:    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "System.UpdateSystemSkills", "DB.ConnectToDB returned FALSE"
    
2527:    Dim iSkillId As Integer, iKnowledgeLevel As Integer
    
2529:    Dim i As Integer
2530:    For i = 0 To UBound(data)
2531:        If data(i).SkillName <> "" And UCase(data(i).SkillName) <> "N/A" Then
            'check if interface is defined in tblInterfacesInterfaces
2533:            iSkillId = FindGetSkillId(data(i).SkillName, FindGetSkillTypeId(data(i).SkillType))
            
2535:            MyDb.pSQLQuery = "Select SystemId from  tblSystemSkills where SystemId=" & Me.SysId & " and SkillId =" & iSkillId & ";"
2536:            MyDb.ExecuteSelect
            
            
            'need to insert the record
2540:            If IsEmpty(MyDb.pDataArray) Then
2541:                iKnowledgeLevel = FindGetSkillKnowLedgeLevelId(data(i).KnowledgeLevel)
2542:                MyDb.pSQLQuery = "insert into tblSystemSkills values ( " & Me.SysId & "," & iSkillId & "," & iKnowledgeLevel & ");"
2543:                MyDb.ExecuteInsert
2544:            End If
2545:        End If
2546:    Next
    
'Exit Point
UpdateSystemSkills_Exit:
2550:    PopCallStack
2551:    On Error GoTo 0
2552:    Exit Sub

UpdateSystemSkills_Error:
2555:    GlobalErrHandler
2556:    Resume UpdateSystemSkills_Exit
         
2558: End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Private Function FindGetSkillKnowLedgeLevelId(KnowledgeLevel As String) As Integer
'
'   function receives Skills Knowledge level name and looks for level Id
'   in tblSkillsKnowledgeLevel. If the Type was not found by Description, function
'   will create a new Knowledge Level with that Description and assigns a new ID for it.
'   Function returns the id of the new Knowledge Level.
'   In case of Error, function returns -1
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FindGetSkillKnowLedgeLevelId(KnowledgeLevel As String) As Integer

    'Error Handling
2574:    If gcfHandleErrors Then On Error GoTo FindGetSkillKnowLedgeLevelId_Error
2575:    PushCallStack "System.FindGetSkillKnowLedgeLevelId"
    
2577:    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "System.FindGetSkillKnowLedgeLevelId", "DB.ConnectToDB returned FALSE"
    
2580:    If KnowledgeLevel = "" Then
2581:        MyDb.pSQLQuery = "select ID from tblSkillsKnowledgeLevel where SkillKnowledgeLevel is null;"
2582:    Else
2583:        MyDb.pSQLQuery = "select ID from tblSkillsKnowledgeLevel where SkillKnowledgeLevel ='" & KnowledgeLevel & "';"
2584:    End If
         
2586:    MyDb.ExecuteSelect
    
2588:    If IsEmpty(MyDb.pDataArray) Then
    
2590:        FindGetSkillKnowLedgeLevelId = InsertNewSkillsKnowledgeLevel(KnowledgeLevel)
2591:    Else
2592:        FindGetSkillKnowLedgeLevelId = CInt(MyDb.pDataArray(0, 0))
2593:    End If
     
    
 
'Exit Point
FindGetSkillKnowLedgeLevelId_Exit:
2599:    PopCallStack
2600:    On Error GoTo 0
2601:    Exit Function

FindGetSkillKnowLedgeLevelId_Error:
2604:    GlobalErrHandler
2605:    FindGetSkillKnowLedgeLevelId = -1
2606:    Resume FindGetSkillKnowLedgeLevelId_Exit
2607: End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Private Function FindGetSkillTypeId(SkillType As String) As Integer
'
'   function receives Skills Type name and looks for skill type Id
'   in tblSkillsType. If the Type was not found by Description, function
'   will create a new SkillType with that Description and assigns a new ID for it.
'   Function returns the id of the new Skill Type.
'   In case of Error, function returns -1
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FindGetSkillTypeId(SkillType As String) As Integer

    'Error Handling
2623:    If gcfHandleErrors Then On Error GoTo FindGetSkillTypeId_Error
2624:    PushCallStack "System.FindGetSkillTypeId"
    
2626:    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "System.FindGetSkillTypeId", "DB.ConnectToDB returned FALSE"
    
2629:    If SkillType = "" Then
2630:        MyDb.pSQLQuery = "select SkillTypeId from tblSkillsType where SkillTypeDescription is null;"
2631:    Else
2632:        MyDb.pSQLQuery = "select SkillTypeId from tblSkillsType where SkillTypeDescription ='" & SkillType & "';"
2633:    End If
     
    
2636:    MyDb.ExecuteSelect
    
2638:    If IsEmpty(MyDb.pDataArray) Then
    
2640:        FindGetSkillTypeId = InsertNewSkillType(SkillType)
2641:    Else
2642:        FindGetSkillTypeId = CInt(MyDb.pDataArray(0, 0))
2643:    End If
     
    
    
'Exit Point
FindGetSkillTypeId_Exit:
2649:    PopCallStack
2650:    On Error GoTo 0
2651:    Exit Function

FindGetSkillTypeId_Error:
2654:    GlobalErrHandler
2655:    FindGetSkillTypeId = -1
2656:    Resume FindGetSkillTypeId_Exit
    
2658: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Private Function InsertNewSkillsKnowledgeLevel(Name As String) As Integer
'
'   Function receives Knowledge level Description and inserts a new Type into the DB.
'   Function returns the ID of the Knowledge Level entered.
'   Important! Only call this function when you are sure the
'   level does not exist in the DB already.
'   Function will not look for the level.
'   Function returns LevelId
'   In case of an Error, function returns -1
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function InsertNewSkillsKnowledgeLevel(Name As String) As Integer
    
    'Error Handling
2676:    If gcfHandleErrors Then On Error GoTo InsertNewSkillsKnowledgeLevel_Error
2677:    PushCallStack "System.InsertNewSkillsKnowledgeLevel"
    
2679:    Dim NewLevelNum As Integer
    
    'Step 1: find out what is the maximum value in the DB
2682:    MyDb.pSQLQuery = "select max(ID) from tblSkillsKnowledgeLevel"
2683:    MyDb.ExecuteSelect
    
2685:    NewLevelNum = CInt(MyDb.pDataArray(0, 0)) + 1
    
2687:    MyDb.pSQLQuery = "insert into tblSkillsKnowledgeLevel values (" & NewLevelNum & ", '" & Name & "');"
2688:    MyDb.ExecuteInsert
    
2690:    InsertNewSkillsKnowledgeLevel = NewLevelNum
    
'Exit Point
InsertNewSkillsKnowledgeLevel_Exit:
2694:    PopCallStack
2695:    On Error GoTo 0
2696:    Exit Function

InsertNewSkillsKnowledgeLevel_Error:
2699:    GlobalErrHandler
2700:    InsertNewSkillsKnowledgeLevel = -1
2701:    Resume InsertNewSkillsKnowledgeLevel_Exit
2702: End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Private Function InsertNewSkillType(Name As String) As Integer
'
'   Function receives skill type Description and inserts a new Type into the DB.
'   Function returns the ID of the type entered.
'   Important! Only call this function when you are sure the
'   type does not exist in the DB already.
'   Function will not look for the type.
'   Function returns LevelId
'   In case of an Error, function returns -1
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function InsertNewSkillType(Name As String) As Integer
    
    'Error Handling
2720:    If gcfHandleErrors Then On Error GoTo InsertNewSkillType_Error
2721:    PushCallStack "System.InsertNewSkillType"
     
2723:    Dim NewTypeNum As Integer
    
    'Step 1: find out what is the maximum value in the DB
2726:    MyDb.pSQLQuery = "select max(ID) from tblSkillsType"
2727:    MyDb.ExecuteSelect
    
2729:    NewTypeNum = CInt(MyDb.pDataArray(0, 0)) + 1
    
2731:    MyDb.pSQLQuery = "insert into tblSkillsType values (" & NewTypeNum & ", '" & Name & "');"
2732:    MyDb.ExecuteInsert
    
2734:    InsertNewSkillType = NewTypeNum
    
'Exit Point
InsertNewSkillType_Exit:
2738:    PopCallStack
2739:    On Error GoTo 0
2740:    Exit Function

InsertNewSkillType_Error:
2743:    GlobalErrHandler
2744:    InsertNewSkillType = -1
2745:    Resume InsertNewSkillType_Exit
    
    
2748: End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Public Sub ArchiveSystem()
'
'   Move system to Archive. Set status = false in tblSystems.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ArchiveSystem()
    
    'Error Handling
2759:    If gcfHandleErrors Then On Error GoTo ArchiveSystem_Error
2760:    PushCallStack "System.ArchiveSystem"
    
    
2763:    MyDb.pSQLQuery = "update tblSystems set status=false where systemid = " & Me.SysId & ";"
2764:    MyDb.ExecuteUpdate
    
'Exit Point
ArchiveSystem_Exit:
2768:    PopCallStack
2769:    On Error GoTo 0
2770:    Exit Sub

ArchiveSystem_Error:
2773:    GlobalErrHandler
2774:    Resume ArchiveSystem_Exit
       
2776: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               Public Sub PermanentDelete()
'
'   Sub permanantely deletes a system from the DB.
'   Note the order of the deletion!
'   records from tblSystemDescription and tblSystems cannot be deleted
'   before system is deleted from all other tables.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PermanentDelete()
            
    'Error Handling
2791:    If gcfHandleErrors Then On Error GoTo PermanentDelete_Error
2792:    PushCallStack "System.PermanentDelete"
    
2794:    If SysId = -1 Then _
        ThrowError CustomError.MANDATORY_ID_IS_MISSING, "System.PermanentDelete", "system Id is mandatory at this stage"
    
    'tblSystemDBTypes
2798:    MyDb.pSQLQuery = "delete from tblSystemDBTypes where systemid = " & Me.SysId & ";"
2799:    MyDb.ExecuteDelete
    
    'tblSystemDevLang
2802:    MyDb.pSQLQuery = "delete from tblSystemDevLang where systemid = " & Me.SysId & ";"
2803:    MyDb.ExecuteDelete
    
    'tblSystemEmployees
2806:    MyDb.pSQLQuery = "delete from tblSystemEmployees where systemid = " & Me.SysId & ";"
2807:    MyDb.ExecuteDelete
    
    'tblSystemInterfaces
2810:    MyDb.pSQLQuery = "delete from tblSystemInterfaces where systemid = " & Me.SysId & ";"
2811:    MyDb.ExecuteDelete
        
    'tblSystemSkills
2814:    MyDb.pSQLQuery = "delete from tblSystemSkills where systemid = " & Me.SysId & ";"
2815:    MyDb.ExecuteDelete
    
    'tblSystemDescription
2818:    MyDb.pSQLQuery = "delete from tblSystemDescription where systemid = " & Me.SysId & ";"
2819:    MyDb.ExecuteDelete
    
    'tblSystems
2822:    MyDb.pSQLQuery = "delete from tblSystems where systemid = " & Me.SysId & ";"
2823:    MyDb.ExecuteDelete
   
'Exit Point
PermanentDelete_Exit:
2827:    PopCallStack
2828:    On Error GoTo 0
2829:    Exit Sub

PermanentDelete_Error:
2832:    GlobalErrHandler
2833:    MsgBox "אירעה שגיאה במחיקת מערכת", vbCritical, "שגיאה"
2834:    Resume PermanentDelete_Exit
    
2836: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Start of Properties Section
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get SysId() As Integer
2844:    SysId = iSysId
2845: End Property

Public Property Let SysId(id As Integer)
2848:    If id > 0 Then
2849:        iSysId = id
2850:    Else
2851:        iSysId = -1
2852:    End If
2853: End Property

Public Property Get SubSystemName() As String
2856:    SubSystemName = sSubSystemName
2857: End Property

Public Property Let SubSystemName(Name As String)
2860:    If Name <> "" Then
2861:        sSubSystemName = Name
        
2863:        If cListToUpdate.Exists("SubSystemName") Then cListToUpdate.Remove ("SubSystemName")
2864:        cListToUpdate.Add "SubSystemName", sSubSystemName
2865:    End If
2866: End Property

Public Property Get SystemDescription() As String
2869:    SystemDescription = sSystemDescription
2870: End Property

Public Property Let SystemDescription(SystemDescription As String)
2873:    If SystemDescription <> "" Then
2874:        sSystemDescription = SystemDescription
        
2876:        If cListToUpdate.Exists("SystemDescription") Then cListToUpdate.Remove ("SystemDescription")
2877:        cListToUpdate.Add "SystemDescription", SystemDescription
2878:    End If
2879: End Property

Public Property Get KnowledgeConsumer() As String
2882:    KnowledgeConsumer = sKnowledgeConsumer
2883: End Property

Public Property Let KnowledgeConsumer(value As String)
2886:    If value <> "" Then
2887:        sKnowledgeConsumer = value
        
2889:        If cListToUpdate.Exists("KnowledgeConsumers") Then cListToUpdate.Remove ("KnowledgeConsumers")
2890:        cListToUpdate.Add "KnowledgeConsumers", sKnowledgeConsumer
2891:    End If
2892: End Property

Public Property Get RoshAnaf() As employee()
2895:    RoshAnaf = eRoshAnaf
2896: End Property

Public Property Let RoshAnaf(Emp() As employee)
    
2900:    eRoshAnaf = Emp
    
2902:    If cListToUpdate.Exists("eRoshAnaf") Then cListToUpdate.Remove ("eRoshAnaf")
2903:    cListToUpdate.Add "eRoshAnaf", eRoshAnaf
2904: End Property

Public Property Get RoshAnafBachir() As employee()
2907:    RoshAnafBachir = eRoshAnafBachir
2908: End Property

Public Property Let RoshAnafBachir(Emp() As employee)
    
2912:    eRoshAnafBachir = Emp
    
2914:    If cListToUpdate.Exists("eRoshAnafBachir") Then cListToUpdate.Remove ("eRoshAnafBachir")
2915:    cListToUpdate.Add "eRoshAnafBachir", eRoshAnafBachir
2916: End Property


Public Property Get RoshMador() As employee()
2920:    RoshMador = eRoshMador
2921: End Property

Public Property Let RoshMador(Emp() As employee)

2925:    eRoshMador = Emp
    
2927:    If cListToUpdate.Exists("eRoshMador") Then cListToUpdate.Remove ("eRoshMador")
2928:    cListToUpdate.Add "eRoshMador", eRoshMador
2929: End Property


Public Property Get Workers() As employee()
2933:    Workers = eWorkers
2934: End Property

Public Property Let Workers(Emp() As employee)

2938:    eWorkers = Emp
    
2940:    If cListToUpdate.Exists("eWorkers") Then cListToUpdate.Remove ("eWorkers")
2941:    cListToUpdate.Add "eWorkers", eWorkers
2942: End Property
 
Public Property Get IsInfrastructure() As Integer
2945:    IsInfrastructure = iIsInfrastructure
2946: End Property

Public Property Let IsInfrastructure(value As Integer)
2949:    iIsInfrastructure = value
    
2951:    If cListToUpdate.Exists("IsInfrastructure") Then cListToUpdate.Remove ("IsInfrastructure")
2952:    cListToUpdate.Add "IsInfrastructure", iIsInfrastructure
2953: End Property

Public Property Get BizDev() As String
2956:    BizDev = sBizDev
2957: End Property

Public Property Let BizDev(value As String)
2960:    If value <> "" Then
2961:        sBizDev = value
        
2963:        If cListToUpdate.Exists("BizEnv") Then cListToUpdate.Remove ("BizEnv")
2964:        cListToUpdate.Add "BizEnv", sBizDev
2965:    End If
2966: End Property

Public Property Get DevEnv() As String
2969:    DevEnv = sDevEnv
2970: End Property

Public Property Let DevEnv(value As String)
2973:    If value <> "" Then
2974:        sDevEnv = value
        
2976:        If cListToUpdate.Exists("DevEnv") Then cListToUpdate.Remove ("DevEnv")
2977:        cListToUpdate.Add "DevEnv", sDevEnv
2978:    End If
2979: End Property

Public Property Get TechEnv() As String
2982:    TechEnv = sTechEnv
2983: End Property

Public Property Let TechEnv(value As String)
2986:    If value <> "" Then
2987:        sTechEnv = value
        
2989:        If cListToUpdate.Exists("TechEnv") Then cListToUpdate.Remove ("TechEnv")
2990:        cListToUpdate.Add "TechEnv", sTechEnv
2991:    End If
2992: End Property

Public Property Get IsCore() As Boolean
2995:    IsCore = bIsCore
2996: End Property

Public Property Let IsCore(value As Boolean)
2999:    bIsCore = value
    
3001:    If cListToUpdate.Exists("IsCore") Then cListToUpdate.Remove ("IsCore")
3002:    cListToUpdate.Add "IsCore", bIsCore
3003: End Property

Public Property Get IsWebBased() As Boolean
3006:    IsWebBased = bIsWebBased
3007: End Property

Public Property Let IsWebBased(value As Boolean)
3010:    bIsWebBased = value
    
3012:    If cListToUpdate.Exists("IsWebBased") Then cListToUpdate.Remove ("IsWebBased")
3013:    cListToUpdate.Add "IsWebBased", bIsWebBased
3014: End Property

Public Property Get NumOfInterfaces() As Integer
3017:    NumOfInterfaces = iNumOfInterfaces
3018: End Property

Public Property Let NumOfInterfaces(value As Integer)
3021:    If value >= 0 Then
3022:        iNumOfInterfaces = value
3023:    Else
3024:        iNumOfInterfaces = -1
3025:    End If
    
3027:    If cListToUpdate.Exists("NumOfInterfaces") Then cListToUpdate.Remove ("NumOfInterfaces")
3028:    cListToUpdate.Add "NumOfInterfaces", iNumOfInterfaces
3029: End Property

Public Property Get NumOfCriticalInterfaces() As Integer
3032:    NumOfCriticalInterfaces = iNumOfCriticalInterfaces
3033: End Property

Public Property Let NumOfCriticalInterfaces(value As Integer)
3036:    If value >= 0 Then
3037:        iNumOfCriticalInterfaces = value
3038:    Else
3039:        iNumOfCriticalInterfaces = -1
3040:    End If
    
3042:    If cListToUpdate.Exists("NumOfCriticalInterfaces") Then cListToUpdate.Remove ("NumOfCriticalInterfaces")
3043:    cListToUpdate.Add "NumOfCriticalInterfaces", iNumOfCriticalInterfaces
3044: End Property

Public Property Get SystemRisks() As String
3047:    SystemRisks = sSystemRisks
3048: End Property

Public Property Let SystemRisks(value As String)
3051:    If value <> "" Then
3052:        sSystemRisks = value
        
3054:        If cListToUpdate.Exists("SystemRisks") Then cListToUpdate.Remove ("SystemRisks")
3055:        cListToUpdate.Add "SystemRisks", sSystemRisks
3056:    End If
3057: End Property

Public Property Get PreservationTopics() As String
3060:    PreservationTopics = sPreservationTopics
3061: End Property

Public Property Let PreservationTopics(value As String)
3064:    If value <> "" Then
3065:        sPreservationTopics = value
        
3067:        If cListToUpdate.Exists("PreservationTopics") Then cListToUpdate.Remove ("PreservationTopics")
3068:        cListToUpdate.Add "PreservationTopics", sPreservationTopics
3069:    End If
3070: End Property

Public Property Get PreservationSuggestions() As String
3073:    PreservationSuggestions = sPreservationSuggestions
3074: End Property

Public Property Let PreservationSuggestions(value As String)
3077:    If value <> "" Then
3078:        sPreservationSuggestions = value
        
3080:        If cListToUpdate.Exists("PreservationSuggestions") Then cListToUpdate.Remove ("PreservationSuggestions")
3081:        cListToUpdate.Add "PreservationSuggestions", sPreservationSuggestions
3082:    End If
3083: End Property

Public Property Get IsChangeManagement() As Integer
3086:    IsChangeManagement = iIsChangeManagement
3087: End Property

Public Property Let IsChangeManagement(value As Integer)
3090:    iIsChangeManagement = value
    
3092:    If cListToUpdate.Exists("IsChangeManagement") Then cListToUpdate.Remove ("IsChangeManagement")
3093:    cListToUpdate.Add "IsChangeManagement", iIsChangeManagement
3094: End Property

Public Property Get IsPlannedToClose() As Integer
3097:    IsPlannedToClose = iIsPlannedToClose
3098: End Property

Public Property Let IsPlannedToClose(value As Integer)
3101:    iIsPlannedToClose = value
    
3103:    If cListToUpdate.Exists("IsPlannedToClose") Then cListToUpdate.Remove ("IsPlannedToClose")
3104:    cListToUpdate.Add "IsPlannedToClose", iIsPlannedToClose
3105: End Property

Public Property Get IsManagementBrief() As Integer
3108:    IsManagementBrief = iIsManagementBrief
3109: End Property

Public Property Let IsManagementBrief(value As Integer)
3112:    iIsManagementBrief = value
    
3114:    If cListToUpdate.Exists("IsManagementBrief") Then cListToUpdate.Remove ("IsManagementBrief")
3115:    cListToUpdate.Add "IsManagementBrief", iIsManagementBrief
3116: End Property

Public Property Get IsApplicableDocuments() As Integer
3119:    IsApplicableDocuments = iIsApplicableDocuments
3120: End Property

Public Property Let IsApplicableDocuments(value As Integer)
3123:    iIsApplicableDocuments = value
    
3125:    If cListToUpdate.Exists("IsApplicableDocuments") Then cListToUpdate.Remove ("IsApplicableDocuments")
3126:    cListToUpdate.Add "IsApplicableDocuments", iIsApplicableDocuments
3127: End Property

Public Property Get IsArchitectureDoc() As Integer
3130:    IsArchitectureDoc = iIsArchitectureDoc
3131: End Property

Public Property Let IsArchitectureDoc(value As Integer)
3134:    iIsArchitectureDoc = value
    
3136:    If cListToUpdate.Exists("IsArchitectureDoc") Then cListToUpdate.Remove ("IsArchitectureDoc")
3137:    cListToUpdate.Add "IsArchitectureDoc", iIsArchitectureDoc
3138: End Property

Public Property Get IsShobDoc() As Integer
3141:    IsShobDoc = iIsShobDoc
3142: End Property

Public Property Let IsShobDoc(value As Integer)
3145:    iIsShobDoc = value
    
3147:    If cListToUpdate.Exists("IsShobDoc") Then cListToUpdate.Remove ("IsShobDoc")
3148:    cListToUpdate.Add "IsShobDoc", iIsShobDoc
3149: End Property

Public Property Get IsTesting() As Integer
3152:    IsTesting = iIsTesting
3153: End Property

Public Property Let IsTesting(value As Integer)
3156:    iIsTesting = value
    
3158:    If cListToUpdate.Exists("IsTesting") Then cListToUpdate.Remove ("IsTesting")
3159:    cListToUpdate.Add "IsTesting", iIsTesting
3160: End Property
Public Property Get IsSupportAndOp() As Integer
3162:    IsSupportAndOp = iIsSupportAndOp
3163: End Property

Public Property Let IsSupportAndOp(value As Integer)
3166:    iIsSupportAndOp = value
    
3168:    If cListToUpdate.Exists("IsSupportAndOp") Then cListToUpdate.Remove ("IsSupportAndOp")
3169:    cListToUpdate.Add "IsSupportAndOp", iIsSupportAndOp
3170: End Property

Public Property Get QaPassFail() As Integer
3173:    QaPassFail = iQaPassFail
3174: End Property

Public Property Let QaPassFail(value As Integer)
3177:    iQaPassFail = value
    
3179:    If cListToUpdate.Exists("QaPassFail") Then cListToUpdate.Remove ("QaPassFail")
3180:    cListToUpdate.Add "QaPassFail", iQaPassFail
3181: End Property

Public Property Get Db() As LeumiDB
3184:    Db = MyDb
3185: End Property

Public Property Get SystemName() As String
3188:    SystemName = sSystemName
3189: End Property

Public Property Let SystemName(value As String)
3192:    If value <> "" Then sSystemName = value
3193: End Property

Public Property Get Comments() As String
3196:    Comments = sComments
3197: End Property

Public Property Let Comments(value As String)
3200:    If value <> "" Then sComments = value
    
3202:    If cListToUpdate.Exists("Comments") Then cListToUpdate.Remove ("Comments")
3203:    cListToUpdate.Add "Comments", sComments
3204: End Property


Public Property Get Log() As Log
3208:    Log = MyLog
3209: End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   End of Properties Section
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''


