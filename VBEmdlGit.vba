Attribute VB_Name = "mdlGit"


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   Sub SaveCodeModules(dir As String)
'   Extracts all the VBA modules into a directory.
'   code downloaded from: http://stackoverflow.com/questions/131605/best-way-to-do-version-control-for-ms-excel
'
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SaveCodeModules(dir As String)

'This code Exports all VBA modules
Dim moduleName As String
Dim vbaType As Integer
Dim i As Integer
With ThisWorkbook.VBProject
    For i = 1 To .VBComponents.count
        If .VBComponents(i).CodeModule.CountOfLines > 0 Then
            moduleName = .VBComponents(i).CodeModule.Name
            vbaType = .VBComponents(i).Type

            If vbaType = 1 Then
                .VBComponents(i).Export dir & moduleName & ".vba"
            ElseIf vbaType = 3 Then
                .VBComponents(i).Export dir & moduleName & ".frm"
            ElseIf vbaType = 100 Then
                .VBComponents(i).Export dir & moduleName & ".cls"
            End If

        End If
    Next i
End With

End Sub


Sub ImportCodeModules(dir As String)

Dim modList(0 To 0) As String
Dim vbaType As Integer

' delete all forms, modules, and code in MEOs
With ThisWorkbook.VBProject
    For Each comp In .VBComponents

        moduleName = comp.CodeModule.Name

        vbaType = .VBComponents(moduleName).Type

        If moduleName <> "DevTools" Then
            If vbaType = 1 Or _
                vbaType = 3 Then

                .VBComponents.Remove .VBComponents(moduleName)

            ElseIf vbaType = 100 Then

                ' we can't simply delete these objects, so instead we empty them
                .VBComponents(moduleName).CodeModule.DeleteLines 1, .VBComponents(moduleName).CodeModule.CountOfLines

            End If
        End If
    Next comp
End With

' make a list of files in the target directory
Set FSO = CreateObject("Scripting.FileSystemObject")
Set dirContents = FSO.GetFolder(dir) ' figure out what is in the directory we're importing

' import modules, forms, and MEO code back into workbook
With ThisWorkbook.VBProject
    For Each moduleName In dirContents.Files

        ' I don't want to import the module this script is in
        If moduleName.Name <> "DevTools.vba" Then

            ' if the current code is a module or form
            If Right(moduleName.Name, 4) = ".vba" Or _
                Right(moduleName.Name, 4) = ".frm" Then

                ' just import it normally
                .VBComponents.Import dir & moduleName.Name

            ' if the current code is a microsoft excel object
            ElseIf Right(moduleName.Name, 4) = ".cls" Then
                Dim count As Integer
                Dim fullmoduleString As String
                Open moduleName.Path For Input As #1

                count = 0              ' count which line we're on
                fullmoduleString = ""  ' build the string we want to put into the MEO
                Do Until EOF(1)        ' loop through all the lines in the file

                    Line Input #1, moduleString  ' the current line is moduleString
                    If count > 8 Then            ' skip the junk at the top of the file

                        ' append the current line `to the string we'll insert into the MEO
                        fullmoduleString = fullmoduleString & moduleString & vbNewLine

                    End If
                    count = count + 1
                Loop

                ' insert the lines into the MEO
                .VBComponents(Replace(moduleName.Name, ".cls", "")).CodeModule.InsertLines .VBComponents(Replace(moduleName.Name, ".cls", "")).CodeModule.CountOfLines + 1, fullmoduleString

                Close #1

            End If
        End If

    Next moduleName
End With

End Sub


