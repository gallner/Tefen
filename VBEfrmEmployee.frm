VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEmployee 
   Caption         =   "����� ������"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13020
   OleObjectBlob   =   "VBEfrmEmployee.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyDb As LeumiDB


Private Sub cmdAddEmployee_Click()
    
    'error handling
    If gcfHandleErrors Then On Error GoTo cmdAddEmployee_Click_Error
    PushCallStack "frmEmployee.cmdAddEmployee_Click"
    
    'case of update
    If "����� �����" = cmdAddEmployee.Caption Then
        Call updateEmployeesDetails
        GoTo cmdAddEmployee_Click_Exit
    End If
    
    
    
    'check that the name is not empty
    If Len(txtFirstName.Text) = 0 Or Len(txtLastName.Text) = 0 Then
        MsgBox "���� ����� �� ���� ��� �����", vbCritical, "�����"
        GoTo cmdAddEmployee_Click_Exit
    End If
    
    'connect to db
    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "frmEmployee.cmdAddEmployee_Click", "LeumiDB.ConnectToDB returned FALSE"
    
    'step 1: make sure there is no other employee with the same name
    Dim Query As String
    Query = "select * from tblEmployees where FirstName='" & txtFirstName.Text & "' and LastName ='" & txtLastName.Text & "';"
    
    MyDb.pSQLQuery = Query
    MyDb.ExecuteSelect
    
    'found the employee
    If IsArrayAllocated(MyDb.pDataArray) Then
        txtEmpId.Text = CInt(MyDb.pDataArray(0, 0))
        txtFirstName.Text = MyDb.pDataArray(1, 0)
        txtLastName.Text = MyDb.pDataArray(2, 0)
        Me.cboRole.Text = mdlEnums.GetPositionById(CInt(MyDb.pDataArray(3, 0)))
        Me.cmdAddEmployee.Caption = "����� �����"
        MsgBox "���� ��� ���� ������", vbInformation, "���� ���� ���"
    Else
        Query = "insert into tblEmployees values('" & txtFirstName.Text & "','" & txtLastName.Text & "');"
        MyDb.pSQLQuery = Query
        MyDb.ExecuteInsert
        Me.cmdAddEmployee.Caption = "����� �����"
        MsgBox "���� ���� ������", vbInformation, "���� ���� ���"
    End If


'exit point
cmdAddEmployee_Click_Exit:
    Call PopCallStack
    Exit Sub
    
cmdAddEmployee_Click_Error:
    Call GlobalErrHandler
    Resume cmdAddEmployee_Click_Exit
    
End Sub

Public Sub updateEmployeesDetails()
    
    'error handling
    If gcfHandleErrors Then On Error GoTo updateEmployeesDetails_Error
    PushCallStack "frmEmployee.updateEmployeesDetails"
    
    'check that the name is not empty
    If Len(txtFirstName.Text) = 0 Or Len(txtLastName.Text) = 0 Then
        MsgBox "���� ����� �� ���� ��� �����", vbCritical, "�����"
        GoTo updateEmployeesDetails_Exit
    End If
    
    'connect to db
    If Not MyDb.ConnectToDB Then _
        ThrowError CustomError.CONNECTION_TO_DB_FAIL, "frmEmployee.cmdAddEmployee_Click", "LeumiDB.ConnectToDB returned FALSE"
    
    'step 1: get the id
    Dim Query As String
    Query = "select employeeid from tblEmployees where FirstName='" & txtFirstName.Text & "' and LastName ='" & txtLastName.Text & "';"
    
    MyDb.pSQLQuery = Query
    MyDb.ExecuteSelect
    
    'found the employee
    If Not IsArrayAllocated(MyDb.pDataArray) Then
        MsgBox "���� �� ���� ������", vbCritical, "�����"
    Else
        Query = "update tblEmployees set firstname = '" & txtFirstName.Text & "', lastname = '" & txtLastName.Text & _
                "', positionid = " & GetPositionIdByName(cboRole.Text) & " where employeeid = " & txtEmpId.Text
        Debug.Print Query
        MyDb.pSQLQuery = Query
        MyDb.ExecuteUpdate
        MsgBox "����� ������ ������", vbInformation, "����� ���� ����"
    End If
    
    
'exit point
updateEmployeesDetails_Exit:
    Call PopCallStack
    Exit Sub
    
updateEmployeesDetails_Error:
    Call GlobalErrHandler
    Resume updateEmployeesDetails_Exit
End Sub


Private Sub UserForm_Initialize()
    Set MyDb = New LeumiDB
    Application.DefaultSheetDirection = xlLTR
End Sub
