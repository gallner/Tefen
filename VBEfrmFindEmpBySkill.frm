VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFindEmpBySkill 
   Caption         =   "מיפוי מערכות"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10185
   OleObjectBlob   =   "VBEfrmFindEmpBySkill.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFindEmpBySkill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Description:  This form is the interface for finding
'               employees that has specific skills set
'               defined by the user.
'               User inserts must have skills and nice to have skills
'               and the system finds employees that matches this skill set.
'               The form uses module: mdlFindEmpBySkill as the logic
'               module for this form
'
'
' Authors:      Nir Gallner, nir@verisoft.co
'
'
' Date                 Comment
' --------------------------------------------------------------
' 9/25/2016            Initial version

Option Explicit
Dim MyDb As LeumiDB

Private Sub ReloadCboSkills()
     
23:    Dim i As Integer
24:    Dim itemToAdd As String
    
    'connect to DB
27:    Call MyDb.ConnectToDB
    
    '''''''''''remove old items''''''
    
    'in case combo box is full- empty it first.
32:    If cboMustHaveSkills.ListCount > 0 Then
33:        Do While cboMustHaveSkills.ListCount > 0
34:            cboMustHaveSkills.RemoveItem 0
35:        Loop
36:    End If
    
    'in case combo box is full- empty it first.
39:    If cboNiceHaveSkills.ListCount > 0 Then
40:        Do While cboNiceHaveSkills.ListCount > 0
41:            cboNiceHaveSkills.RemoveItem 0
42:        Loop
43:    End If
    
         
    'load Skills infromation
47:    MyDb.pSQLQuery = "SELECT SkillDescription FROM tblSkills ;"
    
49:    MyDb.ExecuteSelect
    
51:    If IsArrayAllocated(MyDb.pDataArray) Then
52:        For i = 0 To UBound(MyDb.pDataArray, 2)
53:            itemToAdd = MyDb.pDataArray(0, i)
54:            cboMustHaveSkills.AddItem itemToAdd
55:            cboNiceHaveSkills.AddItem itemToAdd
56:        Next
57:    End If
    
   

61: End Sub

Private Sub cmdAddMustSkillToList_Click()
64:    lstMustHaveSkills.AddItem cboMustHaveSkills.value
65: End Sub

Private Sub cmdAddNiceHaveSkillToList_Click()
68:    lstNiceHaveSkills.AddItem cboNiceHaveSkills.value
69: End Sub

Private Sub cmdCalculate_Click()
    
     'Error handling - identical to all subs and functions
74:    If gcfHandleErrors Then On Error GoTo cmdCalculate_Click_Error
75:    PushCallStack "frmFindEmpBySkill.cmdCalculate_Click"
    
77:    Dim i As Integer
78:    Dim sMustHaveSkills() As String
79:    Dim sNiceToHaveSkills() As String
80:    Dim sResult() As String
    
    'get the must have skills
83:    ReDim sMustHaveSkills(lstMustHaveSkills.ListCount - 1)
84:    For i = 0 To UBound(sMustHaveSkills)
85:        sMustHaveSkills(i) = lstMustHaveSkills.List(i)
86:    Next i
    
    'get the nice to have skills
89:    ReDim sNiceToHaveSkills(lstNiceHaveSkills.ListCount - 1)
90:    For i = 0 To UBound(sNiceToHaveSkills)
91:        sNiceToHaveSkills(i) = lstNiceHaveSkills.List(i)
92:    Next i
    
94:    sResult = FindTheRightEmployee(sMustHaveSkills, sNiceToHaveSkills)
     
96:    Call DisplayArray2(sResult)
97:    Me.Hide

'Exit Point
cmdCalculate_Click_Exit:
101:    PopCallStack
102:    On Error GoTo 0
103:    Exit Sub
     
     
     
'Error Handling
cmdCalculate_Click_Error:
109:        GlobalErrHandler
110:        Resume cmdCalculate_Click_Exit
111: End Sub
    
    

Private Sub cmdClearForm_Click()
116:    lstMustHaveSkills.Clear
117:    lstNiceHaveSkills.Clear
    
119: End Sub

Private Sub UserForm_Initialize()

123:    Set MyDb = New LeumiDB: Call MyDb.InitClass
124:    Call ReloadCboSkills
125: End Sub


Private Sub UserForm_Terminate()
129:    Call MyDb.FinalizeClass: Set MyDb = Nothing
130: End Sub
