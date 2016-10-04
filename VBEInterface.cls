VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Description:  This class is a container for a single interface
'               description of the Leumi project.
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

Private sInterfaceName As String
Private sInterfaceCategory As String
Private sInterfaceType As String
Private sKnowledgeLevel As String


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                               Public Sub InitClass
'   Sub serves as sort of constructor for the class.
'   It initializes relevant parmeters:
'   1. Strings will be assigned value """
'   2. Integers are set to -1
'   3. Boolean values are set to False
'   4. Arrays are set to 2 spots (redim (1)).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitClass()
35:    sInterfaceName = ""
36:    sInterfaceCategory = ""
37:    sInterfaceType = ""
38:    sKnowledgeLevel = ""
39: End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                               Public Sub FinilizeClass()
'   Sub serves as sort of destructor for the class.
'   Currently not doing anything but here for best practice
'   and future features.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FinilizeClass()
    
49: End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   Public Function IsClassEmpty()
'
'   Function checks if the class is empty: all string equals to ""
'
'   Function returns a boolean number representing whether the class is empty or not.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsClassEmpty() As Boolean

61:    Dim result As Boolean
    
63:    If InterfaceCategory = "" And InterfaceName = "" And InterfaceType = "" And KnowledgeLevel = "" Then
64:        result = True
65:    Else
66:       result = False
67:    End If
    
69:    IsClassEmpty = result
    
'Error Handling
IsClassEmpty_Error:
    
74: End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Public Sub PopulateClass(Name As String, Category As String, IType As String, KnowledgeLevel As String)
'
'   function receives Name, Category,Type and Knowledge level
'   as parameter, and populates the class with the data
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateClass(Name As String, Category As String, iType As String, KnowledgeLevel As String)

    'Error Handling
86:    If gcfHandleErrors Then On Error GoTo PopulateInterfaceClass_Error
87:    PushCallStack "Interface.PopulateInterfaceClass"
    
    'if class is populated, delete current data
90:    If Not Me.IsClassEmpty Then Call Me.InitClass
        
    'at least the name of the interface must contain a value
93:    If Name = "" Then ThrowError CustomError.STRING_IS_EMPTY, "Interface.PopulateInterfaceClass", _
                                "String cannot be empty at this point"
    
96:    Me.InterfaceName = Name
97:    Me.InterfaceCategory = Category
98:    Me.InterfaceType = iType
99:    Me.KnowledgeLevel = KnowledgeLevel
    
'Exit Point
PopulateInterfaceClass_Exit:
103:    PopCallStack
104:    On Error GoTo 0
105:    Exit Sub

PopulateInterfaceClass_Error:
108:    GlobalErrHandler
109:    Resume PopulateInterfaceClass_Exit
    
111: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                               Start of Properties Section
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Let InterfaceName(Interface As String)
119:    If Interface <> "" Then sInterfaceName = Interface
120: End Property

Public Property Get InterfaceName() As String
123:    InterfaceName = sInterfaceName
124: End Property

Public Property Let InterfaceCategory(Category As String)
127:    If Category <> "" Then sInterfaceCategory = Category
128: End Property

Public Property Get InterfaceCategory() As String
131:    InterfaceCategory = sInterfaceCategory
132: End Property

Public Property Let InterfaceType(sType As String)
135:    If sType <> "" Then sInterfaceType = sType
136: End Property

Public Property Get InterfaceType() As String
139:    InterfaceType = sInterfaceType
140: End Property

Public Property Let KnowledgeLevel(Level As String)
143:    If Level <> "" Then sKnowledgeLevel = Level
144: End Property

Public Property Get KnowledgeLevel() As String
147:    KnowledgeLevel = sKnowledgeLevel
148: End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                               End of Properties Section
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


