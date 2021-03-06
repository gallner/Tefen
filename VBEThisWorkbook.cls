VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True




Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Call CleanUpSheets
End Sub

'make sure workbook opens
Private Sub Workbook_Open()
    On Error Resume Next
    
    If Err Then
        Dim MyLog As Log
        Set MyLog = New Log
        
        Call MyLog.Log("error: " & Err.Description, "workbook.workbook_open", MyLog.ERROR)
        MsgBox "����� ����� ���� ����� �����. ��� ���� ������ ���� ����� ������ ���� ������ ���� �� ����� ���� ������.", vbInformation, "����� ���� ����� �����"
        
        Application.SetOption "Error Trapping", 2
        
        Set MyLog = Nothing
    End If
End Sub

'delete all sheets except application defined sheets
Sub CleanUpSheets()
    
    Dim i As Integer
    
    Application.DisplayAlerts = False
    
    Dim sht As Variant
    For Each sht In ActiveWorkbook.Worksheets
        
            Select Case sht.Name
            Case "������"
                sht.Visible = False
            Case "SystemLog"
                sht.Visible = False
            Case "��� �������"
                sht.Visible = False
            Case "��� ������"
                sht.Visible = True
            Case "������"
                sht.Visible = False
            Case "Template"
                sht.Visible = False
            Case Else
                Sheets(sht.Name).Delete
            End Select
       
    Next
    
    Application.DisplayAlerts = True
End Sub
