VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Skill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Description:  This class is a container for a single skill
'               description of the Leumi project.
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

Private sSkillName As String
Private sSkillType As String
Private iSkillType As Integer
Private sKnowledgeLevel As String


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                               Public Sub InitClass
'   Sub serves as sort of constructor for the class.
'   It initializes relevant parmeters:
'   1. Strings will be assigned value ""
'   2. Integers are set to -1
'   3. Boolean values are set to False
'   4. Arrays are set to 2 spots (redim (1)).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitClass()
36:    sSkillName = ""
37:    sSkillType = ""
38:    iSkillType = -1
39:    sKnowledgeLevel = ""
40: End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       Public Sub FinalizeClass()
'   Sub serves as sort of destructor for the class.
'   Currently not doing any thing but for future features it's ready
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FinalizeClass()
    
50: End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   Public Function IsClassEmpty()
'
'   Function checks if the class is empty:
'   1. All string equals to ""
'   2. All integers equals to -1
'
'   Function returns a boolean number representing whether the class is empty or not.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsClassEmpty() As Boolean

    'Error Handling
65:    If gcfHandleErrors Then On Error GoTo IsClassEmpty_Error
66:    PushCallStack "Skill.IsClassEmpty"
    
68:    Dim result As Boolean
    
70:    If Me.KnowledgeLevel = "" And Me.SkillType = "" And Me.SkillName = "" And intSkillType = -1 Then
71:        result = True
72:    Else
73:       result = False
74:    End If
    
76:    IsClassEmpty = result
    
'Exit Point
IsClassEmpty_Exit:
80:    PopCallStack
81:    On Error GoTo 0
82:    Exit Function

IsClassEmpty_Error:
85:    GlobalErrHandler
86:    IsClassEmpty = False
87:    Resume IsClassEmpty_Exit
    
89: End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Public Sub PopulateClass(Name As String, Category As String, IType As String, KnowledgeLevel As String)
'
'   function receives Name, Category,Type and Knowledge level
'   as parameter, and populates the class with the data
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateClass(Name As String, sType As String, iType As Integer, Optional Level As String)

    'Error Handling
102:    If gcfHandleErrors Then On Error GoTo PopulateSkillClass_Error
103:    PushCallStack "Skill.PopulateSkillClass"
    
    'if class is populated, delete current data
106:    If Not Me.IsClassEmpty Then Call Me.InitClass
    
    'skill should contain at least skill name
109:    If Len(Name) = 0 Then _
        ThrowError CustomError.STRING_IS_EMPTY, "Skill.PopulateSkillClass", "Skill name cannot be empty at this poing"
    
112:    Me.SkillName = Name
113:    Me.SkillType = sType
114:    intSkillType = iType 'This is a private property. Me will not work here
115:    Me.KnowledgeLevel = Level
 
'Exit Point
PopulateSkillClass_Exit:
119:    PopCallStack
120:    On Error GoTo 0
121:    Exit Sub

PopulateSkillClass_Error:
124:    GlobalErrHandler
125:    Resume PopulateSkillClass_Exit
      
127: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                               Start of Properties Section
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Let SkillName(Name As String)
135:    If Name <> "" Then sSkillName = Name
136: End Property

Public Property Get SkillName() As String
139:    SkillName = sSkillName
140: End Property

Public Property Let SkillType(NewSkillType As String)
143:    If NewSkillType <> "" Then sSkillType = NewSkillType
144: End Property

Public Property Get SkillType() As String
147:    SkillType = sSkillType
148: End Property

Public Property Let KnowledgeLevel(Level As String)
151:    If Level <> "" Then sKnowledgeLevel = Level
152: End Property

Public Property Get KnowledgeLevel() As String
155:    KnowledgeLevel = sKnowledgeLevel
156: End Property

Private Property Let intSkillType(iType As Integer)
159:    If iType <> -1 Then iSkillType = iType
160: End Property

Private Property Get intSkillType() As Integer
163:    intSkillType = iSkillType
164: End Property



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                               End of Properties Section
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



