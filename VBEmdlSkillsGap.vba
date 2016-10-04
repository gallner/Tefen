Attribute VB_Name = "mdlSkillsGap"
Option Explicit

Private Type EmployeeKnowledge
    employee As employee
    knowledge As Dictionary
End Type


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Public Function BuildSkillGapMap(sSystems() As String, sEmpNames() As String) As String()
'
'   Function does the following:
'   1. Builds a list of needed skills - a join of all skills in the list provided
'   2. Figures out (by query) how much is needed of each skill
'   3. Checks the availability of skills knowledge within the selected employees.
'   4. For each employee, in the right place, specifies his level of experties in the required skill
'
'   Function receives a list of systems and a list of employees, and calculate the gap in needed skills
'   This function is built according to the requirement specified in the Aug 2nd, 2016 status
'   presentation presented to the bank by Tefen
'
'   Function returns an array ready to be presented to the user with all the necessary information
'   according to requirements.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BuildSkillGapMap(sSystems() As String, sEmpNames() As String) As String()

    'Error handling - identical to all subs and functions
29:    If gcfHandleErrors Then On Error GoTo BuildSkillGapMap_Error
30:    PushCallStack "mdlSkillsGap.BuildSkillGapMap"
    
    'setup
33:    Dim i As Integer, j As Integer
34:    Dim Query As String
     
    'connect to DB
37:    Dim MyDb As LeumiDB
38:    Set MyDb = New LeumiDB: MyDb.InitClass
39:    If Not MyDb.ConnectToDB Then GoTo BuildSkillGapMap_Error 'add exception here
    
    
    'Part A: Find out how many skills we have and how many knowledge subjects needed to meet requirements.
         
    'step 1: get the system id's
45:    Dim sys As System, sysIds() As String
46:    ReDim sysIds(UBound(sSystems))
    
48:    Set sys = New System: sys.InitClass
49:    For i = 0 To UBound(sSystems)
50:        sysIds(i) = sys.FindGetSystemId(sSystems(i))
51:    Next i

     
    'step 2: get the number of skills and count of needs from the DB
55:     Query = "SELECT tblSkills.SkillDescription, Count(tblSkills.SkillId) AS CountOfSkillId" & _
        " FROM tblSkills INNER JOIN tblSystemSkills ON tblSkills.SkillId = tblSystemSkills.SkillId" & _
        " WHERE tblSystemSkills.SystemId in (" & Join(sysIds, ",") & ")" & _
        " GROUP BY tblSkills.SkillDescription" & _
        " order by tblSkills.SkillDescription;"
    
61:    MyDb.pSQLQuery = Query
62:    MyDb.ExecuteSelect
    
    'build the requirements dictionary
65:    Dim dRequiredSkills As Dictionary
66:    Set dRequiredSkills = New Dictionary
67:    For i = 0 To UBound(MyDb.pDataArray, 2)
68:        dRequiredSkills.Add MyDb.pDataArray(0, i), MyDb.pDataArray(1, i)
69:    Next i
    
    
    'part B - get the current state - how many skills are missing

    'step 1: build employees array
75:    Dim EmployeeIds As String 'will hold all the id's
76:    Dim employees() As EmployeeKnowledge
77:    ReDim employees(UBound(sEmpNames))
    
79:    For i = 0 To UBound(employees)
80:        Set employees(i).employee = New employee: employees(i).employee.InitClass
81:        employees(i).employee.PopulateClass (sEmpNames(i))
82:        EmployeeIds = EmployeeIds & employees(i).employee.EmpId & ","
83:    Next i
    
    'step 2 - get the totals of knowledge
86:    Query = "SELECT tblSkills.SkillDescription, Count(tblEmployeeSkill.EmployeeId) AS CountOfEmployeeId" & _
            " FROM tblEmployees INNER JOIN (tblSkills INNER JOIN tblEmployeeSkill ON tblSkills.SkillId = " & _
            "tblEmployeeSkill.SkillId) ON tblEmployees.EmployeeID = tblEmployeeSkill.EmployeeId where " & _
            "tblEmployees.EmployeeId in (" & Left(EmployeeIds, Len(EmployeeIds) - 1) & ")" & _
            " GROUP BY tblSkills.SkillDescription;"
    
92:    MyDb.pSQLQuery = Query
93:    MyDb.ExecuteSelect
    
    'build the sum of knowledge dictionary
96:    Dim dExistSkills As Dictionary
97:    Set dExistSkills = New Dictionary
98:    For i = 0 To UBound(MyDb.pDataArray, 2)
99:        dExistSkills.Add MyDb.pDataArray(0, i), MyDb.pDataArray(1, i)
100:    Next i

    
    'part C - get the knowledge level of each employees
    
    'step 1: get all the skills knowledge for each employee
106:    For i = 0 To UBound(employees)
107:        Set employees(i).knowledge = New Dictionary
108:        Query = "SELECT tblSkills.SkillDescription , tblEmployeeSkill.KnowledgeLevel FROM tblSkills INNER JOIN " & _
            " (tblEmployees INNER JOIN tblEmployeeSkill ON tblEmployees.employeeid = tblEmployeeSkill.employeeid)" & _
            " ON tblSkills.SkillId= tblEmployeeSkill.SkillId where tblemployees.employeeid =" & employees(i).employee.EmpId & ";"
        
112:        MyDb.pSQLQuery = Query
113:        MyDb.ExecuteSelect
        
115:        For j = 0 To UBound(MyDb.pDataArray, 2)
116:            employees(i).knowledge.Add MyDb.pDataArray(0, j), MyDb.pDataArray(1, j)
117:        Next j
118:    Next i
    
    
    'part D - arrange the data
122:    Dim finalArr() As String
123:    ReDim finalArr(dRequiredSkills.count, UBound(employees) + 4)
    
    'headlines
    'finalArr(0, 0) = "Skill Name"
    'finalArr(0, 1) = "Amount Needed"
    'finalArr(0, 2) = "Amount Exist"
    'finalArr(0, 3) = "Gap"
    
    'summary data
132:    For i = 0 To UBound(finalArr, 1) - 1
133:        finalArr(i + 1, 0) = dRequiredSkills.Keys(i)
134:        finalArr(i + 1, 1) = dRequiredSkills.Items(i)
        
136:        If dExistSkills.Exists(dRequiredSkills.Keys(i)) Then
137:            finalArr(i + 1, 2) = dExistSkills(dRequiredSkills.Keys(i))
138:        Else
139:            finalArr(i + 1, 2) = 0
140:        End If
        
142:        finalArr(i + 1, 3) = CInt(finalArr(i + 1, 2)) - CInt(finalArr(i + 1, 1))
143:    Next i
    
    'employees data
146:    For i = 0 To UBound(employees)
147:        finalArr(0, i + 4) = employees(i).employee.FirstName & " " & employees(i).employee.LastName
        
149:        For j = 0 To UBound(finalArr, 1) - 1
150:            If employees(i).knowledge.Exists(finalArr(j + 1, 0)) Then
151:                finalArr(j + 1, i + 4) = employees(i).knowledge(finalArr(j + 1, 0))
152:            Else
153:                finalArr(j + 1, i + 4) = 0
154:            End If
155:        Next j
156:    Next i
    
    'return resuly
159:    BuildSkillGapMap = finalArr
    
        
'Exit Point
BuildSkillGapMap_Exit:
164:    PopCallStack
165:    On Error GoTo 0
166:    Exit Function
     
'Error Handling
BuildSkillGapMap_Error:
170:        GlobalErrHandler
171:        Resume BuildSkillGapMap_Exit

173: End Function


