Attribute VB_Name = "mdlEnums"
'
' Description:  Basic Enums for the system.
'
' Authors:      Nir Gallner, nir@verisoft.co
'
'
' Date                 Comment
' --------------------------------------------------------------
' 9/25/2016            Initial version
'
Option Explicit


Enum EmployeePosition
    RoshAnafBachir = 1
    RoshAnaf = 2
    RoshMador = 3
    TeamLeader = 4
    ProjectManager = 5
    SystemEngineer = 6
    Programmer = 7
    KnowledgeExpert = 8
    Undefined = 9
End Enum

Enum InterfaceCategory
    InterfacedSystem = 1
    SupportedItem = 2
    Technology = 3
    Unknown = 4
End Enum

Enum InterfaceKnowledgeLevel
    BasicUser = 1
    AdvancedUser = 2
    Admin = 3
    System = 4
    InterfaceKnowledgeOnly = 5
    Undefined = 8
End Enum

Enum InterfaceType
    Inupt = 1
    Output = 2
    Dual = 3
End Enum

Enum SkillKnowledgeLevel
    Low = 1
    Meium = 2
    High = 3
End Enum

Enum SkillType
    Osh = 1
    Technology = 2
    Product = 3
    Interface = 4
    Database = 5
    BankSystem = 6
    OperatingSystem = 7
    ProgrammingLang = 8
    BankProcedure = 9
    BusinessField = 10
    Infrastructure = 11
    Undefined = 12
End Enum

Public Function GetPositionById(PositionId As Integer) As String
    
71:    Dim result As String
    
73:    Select Case PositionId
        Case EmployeePosition.KnowledgeExpert:
75:            result = "����� ���"
76:        Case EmployeePosition.Programmer:
77:            result = "�������"
78:        Case EmployeePosition.ProjectManager:
79:            result = "���� ������"
80:        Case EmployeePosition.RoshAnaf:
81:            result = "��� ���"
82:        Case EmployeePosition.RoshAnafBachir:
83:            result = "��� ��� ����"
84:        Case EmployeePosition.RoshMador:
85:            result = "��� ����"
86:        Case EmployeePosition.SystemEngineer:
87:            result = "���� ������"
88:        Case EmployeePosition.TeamLeader:
89:            result = "��� ����"
90:        Case Else:
91:            result = "�� �����"
92:    End Select
    
94:    GetPositionById = result
    
96: End Function

Public Function GetPositionIdByName(PositionName As String) As Integer
    
100:    Dim result As String
    
102:    Select Case PositionName
        Case "����� ���":
104:            result = EmployeePosition.KnowledgeExpert
105:        Case "�������":
106:            result = EmployeePosition.Programmer
107:        Case "���� ������":
108:            result = EmployeePosition.ProjectManager
109:        Case "��� ���":
110:            result = EmployeePosition.RoshAnaf
111:        Case "��� ��� ����":
112:            result = EmployeePosition.RoshAnafBachir
113:        Case "��� ����":
114:            result = EmployeePosition.RoshMador
115:        Case "���� ������":
116:            result = EmployeePosition.SystemEngineer
117:        Case "��� ����":
118:            result = EmployeePosition.TeamLeader
119:        Case Else:
120:            result = EmployeePosition.Undefined
121:    End Select
    
123:    GetPositionIdByName = result
    
125: End Function





