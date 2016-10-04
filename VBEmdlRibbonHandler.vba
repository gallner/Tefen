Attribute VB_Name = "mdlRibbonHandler"
' Description:  This module handles the ribbon functions.
'               It is the entry point for the entire application.
'               Eich click on the ribbon fires up the associated
'               function in the ribbon and starts the logic path
'               of the application.
'               Efforts were made to make this module as simple UI
'               as possible, and write all the logic else where.
'
' Authors:      Nir Gallner, nir@verisoft.co
'
'
' Date                 Comment
' --------------------------------------------------------------
' 9/25/2016            Initial version
'
Option Explicit
Dim FormDb As LeumiDB

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Sub RbnNewForm(control As IRibbonControl)
'
'   Callback for NewSystemForm onAction
'
'   Sub asks for a new system name.
'   If the system already exist in the DB, System information is loaded.
'   If the system does not exist in the DB, a new form with system name is loaded.
'   Sub copies Sheet: "Template" and assigns the new system name into it
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbnNewForm(control As IRibbonControl)
     
    'Error Handling
33:    If gcfHandleErrors Then On Error GoTo RbnNewForm_Error
34:    PushCallStack "mdlRibbonHandler.RbnNewForm"
     
36:   Call PerfromanceEnabled(True)
        
38:    Dim MyDb As LeumiDB
39:    Set MyDb = New LeumiDB: Call MyDb.InitClass
    
41:    Dim MyLog As Log
42:    Set MyLog = New Log: Call MyLog.InitClass
    
44:    Dim MySystem As System
45:    Set MySystem = New System: Call MySystem.InitClass
    
47:    Dim strName As String
48:    Dim OriginalIndex As Integer
     
    'get the system name
51:    strName = InputBox("אנא הזן את שם המערכת", "הזנת מערכת חדשה")
52:    If strName = vbNullString Then
53:        MsgBox "לא הוזן ערך", vbCritical, "שגיאה בהזנת ערך"
54:        GoTo RbnNewForm_Exit
55:    End If
    
    'create a new sheet
58:    Application.ScreenUpdating = False
59:    With Sheets("Template")
60:        .Visible = True
61:        .Copy After:=Sheets(Sheets.count)
62:        .Visible = False
63:    End With
    
65:    ActiveSheet.Name = strName
66:    Application.ScreenUpdating = True
    
    'search for system in DB
69:    Dim SystemId As Integer
70:    SystemId = MySystem.FindSystemId(Trim(strName))
    
    'found system in DB
73:    If SystemId > 0 Then
74:        MySystem.SysId = SystemId
75:        Call MySystem.FromDbToClass
76:        Call FromClassToForm(MySystem)
77:        MsgBox "שם מערכת כבר קיים, נתוני מערכת נטענו", vbInformation, "הוספת מערכת"
78:    Else
        'focus the cursor on the newly created system sheet
80:        Worksheets(strName).Activate
81:        Worksheets(strName).Cells(2, "C").value = strName
82:        Worksheets(strName).Cells(1, 1).Select
83:    End If
    
85:    Call MyDb.FinalizeClass: Set MyDb = Nothing
86:    Call MyLog.FinalizeClass: Set MyLog = Nothing
87:    Call MySystem.FinalizeClass: Set MySystem = Nothing
    
89:    Call PerfromanceEnabled(False)

'Exit Point
RbnNewForm_Exit:
93:    PopCallStack
94:    On Error GoTo 0
95:    Exit Sub

RbnNewForm_Error:
98:    GlobalErrHandler
99:    Resume RbnNewForm_Exit
    
101: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Sub RbnInsertToDb(control As IRibbonControl)
'
'   Callback for InsertToDB onAction
'
'   Sub Inserts a new system name.
'   If the system already exist in the DB, System updates the system info in the DB.
'   sub is identical to sub RbnUpdateSysInDB
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbnInsertToDb(control As IRibbonControl)
    
    'Error Handling
117:    If gcfHandleErrors Then On Error GoTo RbnInsertToDb_Error
118:    PushCallStack "mdlRibbonHandler.RbnInsertToDb"
    
120:    If Not isValidSystemForm() Then _
        ThrowError CustomError.INVALID_FORM_INPUT, "mdlRibbonHandler.RbnUpdateSysInDB", "Form is not a valid system map form"
    
123:    Dim NewSystem As System
124:    Set NewSystem = New System: Call NewSystem.InitClass
    
126:    Set NewSystem = FromFormToClass()
127:    Call NewSystem.FromClassToDb
    
129:    Call NewSystem.FinalizeClass: Set NewSystem = Nothing
    
131:    MsgBox "מערכת הוזנה למסד הנתונים" & vbCrLf & "יש לטעון מחדש את רשימת המערכות", vbOKOnly, "הזנת מערכת"
    
    
'Exit Point
RbnInsertToDb_Exit:
136:    PopCallStack
137:    On Error GoTo 0
138:    Exit Sub

RbnInsertToDb_Error:
141:    GlobalErrHandler
142:    MsgBox "אירעה שגיאה בהזנת מערכת", vbCritical, "שגיאה"
143:    Resume RbnInsertToDb_Exit
    
145: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Sub RbnUpdateSysInDB(control As IRibbonControl)
'
'   Callback for UpdateSystemToDB onAction
'
'   Sub Inserts a new system name.
'   If the system already exist in the DB, System updates the system info in the DB.
'   sub is identical to sub RbnInsertToDb
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbnUpdateSysInDB(control As IRibbonControl)

    'Error Handling
161:    If gcfHandleErrors Then On Error GoTo RbnUpdateSysInDB_Error
162:    PushCallStack "mdlRibbonHandler.RbnUpdateSysInDB"
    
164:    If Not isValidSystemForm() Then _
        ThrowError CustomError.INVALID_FORM_INPUT, "mdlRibbonHandler.RbnUpdateSysInDB", "Form is not a valid system map form"
    
167:    Dim NewSystem As System
168:    Set NewSystem = New System: Call NewSystem.InitClass
    
170:    Set NewSystem = FromFormToClass()
171:    Call NewSystem.FromClassToDb
    
173:    Call NewSystem.FinalizeClass: Set NewSystem = Nothing
    
175:    MsgBox "מערכת עודכנה במסד הנתונים" & vbCrLf & "יש לטעון מחדש את מפת המערכות", vbOKOnly, "עדכון מערכת"
    
'Exit Point
RbnUpdateSysInDB_Exit:
179:    PopCallStack
180:    On Error GoTo 0
181:    Exit Sub

RbnUpdateSysInDB_Error:
184:    GlobalErrHandler
185:    MsgBox "אירעה שגיאה בעת עדכון המערכת", vbCritical, "שגיאה"
186:    Resume RbnUpdateSysInDB_Exit
    
188: End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Sub RbnPermanentDelete(control As IRibbonControl)
'
'   Callback for PermanentDelete onAction
'
'   Sub Permanantely deletes a system (activesheet) from the DB
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbnPermanentDelete(control As IRibbonControl)

    'Error Handling
201:    If gcfHandleErrors Then On Error GoTo RbnPermanentDelete_Error
202:    PushCallStack "mdlRibbonHandler.RbnPermanentDelete"
    
204:    If Not isValidSystemForm() Then _
        ThrowError CustomError.INVALID_FORM_INPUT, "mdlRibbonHandler.RbnUpdateSysInDB", "Form is not a valid system map form"
        
207:    Dim SystemToDelete As System
   
209:   Call PerfromanceEnabled(True)
    
211:    Set SystemToDelete = New System: SystemToDelete.InitClass
    
213:    Dim SystemName As String
214:    SystemName = ActiveWorkbook.ActiveSheet.Range("c2").value
    
    'are you sure?????
217:    If vbYes = MsgBox("האם למחוק מערכת " & SystemName & "? לא ניתן לשנות לאחר מחיקה", vbYesNo, "מחיקת מערכת") Then
218:        SystemToDelete.SysId = SystemToDelete.FindSystemId(SystemName)
        
        'no system found
221:        If SystemToDelete.SysId = -1 Then
222:            MsgBox "לא נמצאה מערכת למחיקה", vbOKOnly, "מחיקת מערכת"
223:            GoTo RbnPermanentDelete_Exit
224:        Else
        
            'System going for PERMANENT DELETE!
227:            Call SystemToDelete.PermanentDelete
            
229:            Application.DisplayAlerts = False
230:            ActiveWorkbook.ActiveSheet.Delete
231:            Application.DisplayAlerts = True
232:        End If
        
234:    Else
235:        GoTo RbnPermanentDelete_Exit
236:    End If
    
238:    MsgBox "מחיקה הסתיימה בהצלחה" & vbCrLf & "יש לטעון מחדש את רשימת המערכות", vbOKOnly, "מחיקת מערכת"
239:    Call SystemToDelete.FinalizeClass: Set SystemToDelete = Nothing
    
241:   Call PerfromanceEnabled(False)
    
'Exit Point
RbnPermanentDelete_Exit:
245:    PopCallStack
246:    On Error GoTo 0
247:    Exit Sub

RbnPermanentDelete_Error:
250:    GlobalErrHandler
251:    MsgBox "אירעה שגיאה בעת מחיקת המערכת", vbCritical, "שגיאה"
252:    Resume RbnPermanentDelete_Exit
    
254: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Sub RbnArchiveSystem(control As IRibbonControl)
'
'   Callback for ArchiveSystem onAction
'
'   Sub archives a system (activesheet) from the DB
'   and deletes the activeSheet
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbnArchiveSystem(control As IRibbonControl)
    
    'Error Handling
269:    If gcfHandleErrors Then On Error GoTo RbnArchiveSystem_Error
270:    PushCallStack "mdlRibbonHandler.RbnArchiveSystem"
    
272:    If Not isValidSystemForm() Then _
        ThrowError CustomError.INVALID_FORM_INPUT, "mdlRibbonHandler.RbnUpdateSysInDB", "Form is not a valid system map form"
    
275:    Dim SystemToArchive As System
276:    Set SystemToArchive = New System: SystemToArchive.InitClass
    
278:    Dim SystemName As String
279:    SystemName = ActiveWorkbook.ActiveSheet.Range("c2").value
    
    'are you sure?????
282:    If vbYes = MsgBox("האם להעביר מערכת " & SystemName & "לארכיון?", vbYesNo, "מחיקת מערכת") Then
283:        SystemToArchive.SysId = SystemToArchive.FindSystemId(SystemName)
        
        'no system found
286:        If SystemToArchive.SysId = -1 Then
287:            MsgBox "לא נמצאה מערכת", vbOKOnly, "מחיקת מערכת"
288:            GoTo RbnArchiveSystem_Exit
289:        Else
        
            'System going for ARCHIVE
292:            Call SystemToArchive.ArchiveSystem
            
294:            Application.DisplayAlerts = False
295:            ActiveWorkbook.ActiveSheet.Delete
296:            Application.DisplayAlerts = True
297:        End If
        
299:    Else
300:        GoTo RbnArchiveSystem_Exit
301:    End If
    
303:    MsgBox "מערכת הועברה לארכיון" & vbCrLf & "יש לטעון מחדש את רשימת המערכות", vbOKOnly, "סיום תהליך"
304:    Call SystemToArchive.FinalizeClass: Set SystemToArchive = Nothing
    
'Exit Point
RbnArchiveSystem_Exit:
308:    PopCallStack
309:    On Error GoTo 0
310:    Exit Sub

RbnArchiveSystem_Error:
313:    GlobalErrHandler
314:    MsgBox "אירעה שגיאה בעת העברת מערכת לארכיון", vbCritical, "שגיאה"
315:    Resume RbnArchiveSystem_Exit
    
317: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Sub RbnSearchSystem(control As IRibbonControl)
'
'   Callback for SearchSystem onAction
'
'   sub loads a new instance of frmSystems form with a flag
'   to show ACTIVE systems (isActiveSystem = True)
'   sub is identical to RbnSearchArchive, but with an activeSystem flag = true
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbnSearchSystem(control As IRibbonControl)
    
    'Error Handling
332:    If gcfHandleErrors Then On Error GoTo RbnSearchSystem_Error
333:    PushCallStack "mdlRibbonHandler.RbnSearchSystem"
    
335:    Dim oSht As Worksheet


338:    Dim myForm As frmSystems
339:    Set myForm = New frmSystems
    
341:    frmSystems.isActiveSystem = True
342:    frmSystems.Show
    
'Exit Point
RbnSearchSystem_Exit:
346:    PopCallStack
347:    On Error GoTo 0
348:    Exit Sub

RbnSearchSystem_Error:
351:    GlobalErrHandler
352:    Resume RbnSearchSystem_Exit
    
354: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Sub RbnSearchArchive(control As IRibbonControl)
'
'   Callback for SearchArchive onAction
'
'   sub loads a new instance of frmSystems form with a flag
'   to show ARCHIVED systems (isActiveSystem = False)
'   sub is identical to RbnSearchSystem, but with an activeSystem flag = False
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbnSearchArchive(control As IRibbonControl)

    'Error Handling
370:    If gcfHandleErrors Then On Error GoTo RbnSearchArchive_Error
371:    PushCallStack "mdlRibbonHandler.RbnSearchArchive"
    
373:    Dim myForm As frmSystems
374:    Set myForm = New frmSystems
    
376:    frmSystems.isActiveSystem = False
377:    frmSystems.Show
    
'Exit Point
RbnSearchArchive_Exit:
381:    PopCallStack
382:    On Error GoTo 0
383:    Exit Sub

RbnSearchArchive_Error:
386:    GlobalErrHandler
387:    Resume RbnSearchArchive_Exit
    
389: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Sub RbnConnectToDb(control As IRibbonControl)
'
'   Callback for Connect onAction
'   Test Connection to the DB
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbnConnectToDb(control As IRibbonControl)

    'Error Handling
401:    If gcfHandleErrors Then On Error GoTo RbnConnectToDb_Error
402:    PushCallStack "mdlRibbonHandler.RbnConnectToDb"
    
404:    Dim MyDb As LeumiDB
405:    Set MyDb = New LeumiDB
    
407:    If MyDb.ConnectToDB Then
408:        MsgBox "חיבור למסד הנתונים הצליח", vbOKOnly, "חיבור למסד הנתונים"
409:        Else
410:            MsgBox "חיבור למסד הנתונים נכשל", vbCritical, "חיבור למסד הנתונים"
411:        End If
    
413:    Call MyDb.FinalizeClass: Set MyDb = Nothing
    
415:    Set MyDb = Nothing
       
'Exit Point
RbnConnectToDb_Exit:
419:    PopCallStack
420:    On Error GoTo 0
421:    Exit Sub

RbnConnectToDb_Error:
424:    GlobalErrHandler
425:    Resume RbnConnectToDb_Exit
   
    
    
429: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Sub RbnAddSkill(control As IRibbonControl)
'
'   Callback for AddSkill onAction
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbnAddSkill(control As IRibbonControl)
438:    Call cmdAddNewItemToList("M")
439: End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Sub RbnAddInrerfaceCategory(control As IRibbonControl)
'
'   Callback for AddInterfaceCategory onAction
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbnAddInrerfaceCategory(control As IRibbonControl)
447:    Call cmdAddNewItemToList("A")
448: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Sub RbnAddInterfaceType(control As IRibbonControl)
'
'   Callback for AddSInterfceType onAction
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbnAddInterfaceType(control As IRibbonControl)
457:    Call cmdAddNewItemToList("C")
458: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Sub RbnAddInterfaceKnowledgeLevel(control As IRibbonControl)
'
'   Callback for AddInterfaceKnowledge onAction
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbnAddInterfaceKnowledgeLevel(control As IRibbonControl)
467:    Call cmdAddNewItemToList("E")
468: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Sub RbnAddSkillType(control As IRibbonControl)
'
'   Callback for AddSkillType onAction
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbnAddSkillType(control As IRibbonControl)
477:    Call cmdAddNewItemToList("G")
478: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Sub RbnAddSkillKnowledgeLevel(control As IRibbonControl)
'
'   Callback for AddSkillKnowledgeLevel onAction
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbnAddSkillKnowledgeLevel(control As IRibbonControl)
487:    Call cmdAddNewItemToList("I")
488: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Sub RbnViewSystemMap(control As IRibbonControl)
'
'   Callback for ViewSystemMap onAction
'   View the system map sheet called:  "מפת המערכת"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbnViewSystemMap(control As IRibbonControl)

    'Error Handling
500:    If gcfHandleErrors Then On Error GoTo RbnViewSystemMap_Error
501:    PushCallStack "mdlRibbonHandler.RbnViewSystemMap"
    
503:    Sheets("מפת המערכת").Visible = True
504:    Sheets("מפת המערכת").Activate
505:    Sheets("מפת המערכת").Range("a1").Select
    
'Exit Point
RbnViewSystemMap_Exit:
509:    PopCallStack
510:    On Error GoTo 0
511:    Exit Sub

RbnViewSystemMap_Error:
514:    GlobalErrHandler
515:    Resume RbnViewSystemMap_Exit
516: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Sub RbnViewArchiveMap(control As IRibbonControl)
'
'   Callback for ViewArchiveMap onAction
'   View the archived system map sheet called:  "ארכיון"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbnViewArchiveMap(control As IRibbonControl)
    
    'Error Handling
528:    If gcfHandleErrors Then On Error GoTo RbnViewArchiveMap_Error
529:    PushCallStack "mdlRibbonHandler.RbnViewArchiveMap"
    
531:    Sheets("ארכיון").Visible = True
532:    Sheets("ארכיון").Activate
533:    Sheets("ארכיון").Range("a1").Select
    
'Exit Point
RbnViewArchiveMap_Exit:
537:    PopCallStack
538:    On Error GoTo 0
539:    Exit Sub

RbnViewArchiveMap_Error:
542:    GlobalErrHandler
543:    Resume RbnViewArchiveMap_Exit
    
545: End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'           Sub RbnSaveAndExit(control As IRibbonControl)
'
'   Callback for SaveAndExit onAction
'   Save and Exit. Note!! There is code running when save and exit in
'   This WorkBook. You should take a look there, since there is the
'   real logic of save and closes the workbook.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbnSaveAndExit(control As IRibbonControl)

    'Error Handling
558:    If gcfHandleErrors Then On Error GoTo RbnSaveAndExit_Error
559:    PushCallStack "mdlRibbonHandler.RbnSaveAndExit"
    
561:   Call PerfromanceEnabled(True)
562:    ActiveWorkbook.Save
563:   Call PerfromanceEnabled(False)
564:    ActiveWorkbook.Close
    
'Exit Point
RbnSaveAndExit_Exit:
568:    PopCallStack
569:    On Error GoTo 0
570:    Exit Sub

RbnSaveAndExit_Error:
573:    GlobalErrHandler
574:    Resume RbnSaveAndExit_Exit

576: End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'           Sub RbnSaveNoExit(control As IRibbonControl)
'
'   Callback for Save onAction
'   Save workbook, no exit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbnSaveNoExit(control As IRibbonControl)
    
    'Error Handling
587:    If gcfHandleErrors Then On Error GoTo RbnSaveNoExit_Error
588:    PushCallStack "mdlRibbonHandler.RbnSaveNoExit"
    
590:   Call PerfromanceEnabled(True)
591:    ActiveWorkbook.Save
592:    Call PerfromanceEnabled(False)
    
'Exit Point
RbnSaveNoExit_Exit:
596:    PopCallStack
597:    On Error GoTo 0
598:    Exit Sub

RbnSaveNoExit_Error:
601:    GlobalErrHandler
602:    Resume RbnSaveNoExit_Exit

604: End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   Sub RbnGetDBPath(control As IRibbonControl)
'
'   Sub opens a dialog box to select the .accdb (access Database)
'   location. Sub stores the location path into SystemLog sheet
'   custom property.
'
'   Callback for Bind onAction
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbnGetDBPath(control As IRibbonControl)

    'Error Handling
619:    If gcfHandleErrors Then On Error GoTo RbnGetDBPath_Error
620:    PushCallStack "mdlRibbonHandler.RbnGetDBPath"
    
    
623:    Dim fd As FileDialog
624:    Dim selectedPath As String
625:    Dim i As Integer
    
    'set a dialog to select the DB
628:    Set fd = Application.FileDialog(msoFileDialogFilePicker)
629:    With fd     'Configure dialog box
630:        .AllowMultiSelect = False
631:        .Title = "Select DataBase"
632:        .InitialFileName = ""
633:        .Filters.Clear
634:        .Filters.Add "Access DB", "*.accdb"
        
        'Show the dialog and collect file paths selected by the user
637:        If .Show = -1 Then   'User clicked Open
638:            selectedPath = .SelectedItems(1)
639:        Else
640:           GoTo RbnGetDBPath_Exit
641:        End If
642:    End With
643:    Set fd = Nothing
    
645:     Dim wksSheet1 As Worksheet

647:    Set wksSheet1 = Sheets("SystemLog")

    ' Add metadata to worksheet.
650:    On Error Resume Next
651:    wksSheet1.CustomProperties(1).Delete
652:    Err.Clear
653:    On Error GoTo 0
654:    wksSheet1.CustomProperties.Add Name:="selectedPath", value:=selectedPath
    
    'clean up
657:    Set fd = Nothing
658:    Set wksSheet1 = Nothing
    
'Exit Point
RbnGetDBPath_Exit:
662:    PopCallStack
663:    On Error GoTo 0
664:    Exit Sub

RbnGetDBPath_Error:
667:    GlobalErrHandler
668:    Resume RbnGetDBPath_Exit
    
670: End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Sub RbnReloadSystemMap(control As IRibbonControl)
'
'   Callback for ReloadSystemMap onAction
'   Sub deletes the Map list from the sheet and reloads the data
'   from the DB
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbnReloadSystemMap(control As IRibbonControl)

    'Error Handling
683:    If gcfHandleErrors Then On Error GoTo RbnReloadSystemMap_Error
684:    PushCallStack "mdlRibbonHandler.RbnReloadSystemMap"

686:    Call PerfromanceEnabled(True)

    'step 1: clean map sheet
689:    Sheets("מפת המערכת").Activate
    
691:    Dim MyRange As String, LastRow As Integer
692:    LastRow = Sheets("מפת המערכת").Cells(Sheets("מפת המערכת").Rows.count, "A").End(xlUp).row + 1
    
694:    MyRange = "A3:AG" & LastRow
695:    ActiveSheet.Range(MyRange).value = ""
    
    'step 2: build map
698:    Call BuildMap("System")
    

    
702:    MsgBox "טעינת מערכות הסתיימה בהצלחה", vbOKOnly, "טעינת מערכות"
        
704:        Call PerfromanceEnabled(False)

'Exit Point
RbnReloadSystemMap_Exit:
708:    PopCallStack
709:    On Error GoTo 0
710:    Exit Sub

RbnReloadSystemMap_Error:
713:    GlobalErrHandler
714:    Resume RbnReloadSystemMap_Exit
    
716: End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Sub RbReloadArchiveMap(control As IRibbonControl)
'
'   Callback for ReloadArchiveMap onAction
'   Sub deletes the Archive Map list from the sheet and reloads the data
'   from the DB
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbReloadArchiveMap(control As IRibbonControl)
     
    'Error Handling
729:    If gcfHandleErrors Then On Error GoTo RbReloadArchiveMap_Error
730:    PushCallStack "mdlRibbonHandler.RbReloadArchiveMap"
    
732:     Call PerfromanceEnabled(True)
     
     'step 1: clean map sheet
735:    Sheets("ארכיון").Activate
    
737:    Dim MyRange As String, LastRow As Integer
738:    LastRow = Sheets("ארכיון").Cells(Sheets("ארכיון").Rows.count, "A").End(xlUp).row + 1
    
740:    MyRange = "A3:AG" & LastRow
741:    ActiveSheet.Range(MyRange).value = ""
    
    'step 2: build map
744:    Call BuildMap("Archive")
    
746:    MsgBox "טעינת מערכות לארכיון הסתיימה בהצלחה", vbOKOnly, "טעינת מערכות"
747:   Call PerfromanceEnabled(False)
    
'Exit Point
RbReloadArchiveMap_Exit:
751:    PopCallStack
752:    On Error GoTo 0
753:    Exit Sub

RbReloadArchiveMap_Error:
756:    GlobalErrHandler
757:    Resume RbReloadArchiveMap_Exit
    
759: End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Sub RbReloadLists(control As IRibbonControl)
'
'   Callback for ReloaLists onAction
'   Sub deletes the lists of values from the list spreadsheet
'   and reloads the new list from the DB
'   TODO! retrieve the new informatin from the DB
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RbReloadLists(control As IRibbonControl)

    'Error Handling
773:    If gcfHandleErrors Then On Error GoTo RbReloadLists_Error
774:    PushCallStack "mdlRibbonHandler.RbReloadLists"

776:    Call BuildListsSheet

'Exit Point
RbReloadLists_Exit:
780:    PopCallStack
781:    On Error GoTo 0
782:    Exit Sub

RbReloadLists_Error:
785:    GlobalErrHandler
786:    Resume RbReloadLists_Exit
    
788: End Sub





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Private Sub cmdAddNewItemToList(col As String)
'
'   General sub, receives a col letter, which represents a list to add value to
'   in the lists sheet (called: "רשימות") and goto next in line
'   in order to add a value.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAddNewItemToList(col As String)
    
    'Error Handling
804:    If gcfHandleErrors Then On Error GoTo cmdAddNewItemToList_Error
805:    PushCallStack "mdlRibbonHandler.cmdAddNewItemToList"
    
    
808:    Dim LastRow As Integer
    
810:    Sheets("רשימות").Visible = True
811:    Sheets("רשימות").Activate
812:    LastRow = Sheets("רשימות").Range(col & Rows.count).End(xlUp).row
813:    Sheets("רשימות").Range(col & LastRow + 1).Select
    
'Exit Point
cmdAddNewItemToList_Exit:
817:    PopCallStack
818:    On Error GoTo 0
819:    Exit Sub

cmdAddNewItemToList_Error:
822:    GlobalErrHandler
823:    Resume cmdAddNewItemToList_Exit
    
825: End Sub

'Callback for FindEmployeeBySkill onAction
Sub RbnFindEmployeeBySkill(control As IRibbonControl)
829:    Load frmFindEmpBySkill
830:    frmFindEmpBySkill.Show
831: End Sub

'Callback for SkillGap onAction
Sub RbnSkillGap(control As IRibbonControl)
835:    Load frmSkillsGap
836:    frmSkillsGap.Show
837: End Sub

'Callback for AddEmployee onAction
Sub RbnAddNewEmployee(control As IRibbonControl)
841:    frmEmployee.Show
842: End Sub

'Callback for FindEmployee onAction
Sub RbnFindEmployee(control As IRibbonControl)
846: End Sub

'Callback for ArchiveEmployee onAction
Sub RbnArchiveEmployee(control As IRibbonControl)
850: End Sub

'Callback for DeleteEmployee onAction
Sub RbnDeleteEmployee(control As IRibbonControl)
854: End Sub
