Attribute VB_Name = "mdlFindEmpBySkill"
Option Explicit

Private Type EmployeeKnowledge
    employee As employee
    knowledge As Dictionary
End Type


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Public Function FindTheRightEmployee(sMustHaveSkills() As String, sNiceToHaveSkills() As String) As String()
'
'   Function receives 2 arrays:
'   1. must have skills
'   2. nice to have skills
'
'   Function looks for employees who has all the must have skills and builds an array with the knowledge level
'   of the must have skills and the nice to have skills.
'
'   This function is built according to the requirement specified in the Aug 2nd, 2016 status
'   presentation presented to the bank by Tefen
'
'   Function returns an array ready to be presented to the user with all the necessary information
'   according to requirements.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FindTheRightEmployee(sMustHaveSkills() As String, sNiceToHaveSkills() As String) As String()

    'Error handling - identical to all subs and functions
29:    If gcfHandleErrors Then On Error GoTo FindTheRightEmployee_Error
30:    PushCallStack "mdlFindEmpBySkill.FindTheRightEmployee"

    'setup
33:    Dim i As Integer, j As Integer
34:    Dim Query As String
     
    'connect to DB
37:    Dim MyDb As LeumiDB
38:    Set MyDb = New LeumiDB: MyDb.InitClass
39:    If Not MyDb.ConnectToDB Then GoTo FindTheRightEmployee_Error 'add exception here
    
    'part a - get the list of employee id's with the must haves
            
43:    Query = "SELECT DISTINCT [FirstName]+' '+[LastName] AS Expr2,a.employeeid " & _
            " FROM (select tblemployeeskill.employeeid, tblemployeeskill.skillid " & _
            " from tblemployeeskill  inner join tblskills on tblemployeeskill.skillid " & _
            " = tblskills.skillid where tblskills.skilldescription  in ('" & Join(sMustHaveSkills, "','") & "'))" & _
            " AS a INNER JOIN tblEmployees ON a.employeeid = tblEmployees.EmployeeID " & _
            " GROUP BY [FirstName]+' '+[LastName], a.employeeid HAVING Count(a.skillid)>=" & _
             CInt(UBound(sMustHaveSkills) + 1) & ";"

    
52:    MyDb.pSQLQuery = Query
53:    MyDb.ExecuteSelect
        
    'no one fits
56:    If Not IsArrayAllocated(MyDb.pDataArray) Then GoTo FindTheRightEmployee_Exit
    
    'part b - get all the employees in the list and their skills
59:    Dim employees() As EmployeeKnowledge
60:    ReDim employees(UBound(MyDb.pDataArray, 2))
    
    'get the list of the employees
63:    For i = 0 To UBound(employees)
64:        Set employees(i).employee = New employee: employees(i).employee.InitClass
65:        Call employees(i).employee.PopulateClass(CStr(MyDb.pDataArray(0, i)), CInt(MyDb.pDataArray(1, i)))
66:    Next i
    
    'get their list of skills
69:    For i = 0 To UBound(employees)
70:        Set employees(i).knowledge = New Dictionary
71:        Query = "SELECT tblSkills.SkillDescription , tblEmployeeSkill.KnowledgeLevel FROM tblSkills INNER JOIN " & _
            " (tblEmployees INNER JOIN tblEmployeeSkill ON tblEmployees.employeeid = tblEmployeeSkill.employeeid)" & _
            " ON tblSkills.SkillId= tblEmployeeSkill.SkillId where tblemployees.employeeid =" & employees(i).employee.EmpId & ";"
        
75:        MyDb.pSQLQuery = Query
76:        MyDb.ExecuteSelect
        
78:        For j = 0 To UBound(MyDb.pDataArray, 2)
79:            employees(i).knowledge.Add MyDb.pDataArray(0, j), MyDb.pDataArray(1, j)
80:        Next j
81:    Next i
    
    'present the data
84:    Dim finalArr() As String
85:    ReDim finalArr(UBound(employees) + 1, CInt(UBound(sMustHaveSkills) + 1 + UBound(sNiceToHaveSkills) + 1))
86:    finalArr(0, 0) = "employees / Skills"
    
    'list of skills
89:    For i = 0 To UBound(sMustHaveSkills)
90:        finalArr(0, i + 1) = sMustHaveSkills(i)
91:    Next i
    
93:    For j = 0 To UBound(sNiceToHaveSkills)
94:        finalArr(0, j + i + 1) = sNiceToHaveSkills(j)
95:    Next j
    
    
    'list of employees
99:    For i = 0 To UBound(employees)
100:        finalArr(i + 1, 0) = employees(i).employee.FirstName & " " & employees(i).employee.LastName
101:        For j = 1 To UBound(finalArr, 2)
102:            If employees(i).knowledge.Exists(finalArr(0, j)) Then
103:                finalArr(i + 1, j) = employees(i).knowledge(finalArr(0, j))
104:            End If
105:        Next j
        
107:    Next i

    'return data
110:    FindTheRightEmployee = finalArr
    
    
'Exit Point
FindTheRightEmployee_Exit:
115:    PopCallStack
116:    On Error GoTo 0
117:    Exit Function
     
'Error Handling
FindTheRightEmployee_Error:
121:        GlobalErrHandler
122:        Resume FindTheRightEmployee_Exit

124: End Function
