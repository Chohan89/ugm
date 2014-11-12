Attribute VB_Name = "VersionControl"
Sub SaveCodeModules()
    Dim i%, sName$, wName$, wPath$
    With ThisWorkbook.VBProject
        For i% = 1 To .VBComponents.Count
            If .VBComponents(i%).CodeModule.CountOfLines > 0 Then
                sName$ = .VBComponents(i%).CodeModule.Name
                wName$ = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
                wPath$ = ThisWorkbook.Path & "\" & wName$ & "\"
                On Error Resume Next
                MkDir wPath$
                On Error GoTo 0
                .VBComponents(i%).Export wPath$ & sName$ & ".vba"
            End If
        Next i
    End With

End Sub

Sub ImportCodeModules()
    With ThisWorkbook.VBProject
        For i% = 1 To .VBComponents.Count
            ModuleName = .VBComponents(i%).CodeModule.Name
            If .VBComponents(i%).Type = vbext_ct_StdModule And ModuleName <> "VersionControl" Then
                wName$ = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
                wPath$ = ThisWorkbook.Path & "\" & wName$ & "\"
                If Dir(wPath$ & ModuleName & ".vba") <> "" Then
                    .VBComponents.Remove .VBComponents(ModuleName)
                    .VBComponents.Import wPath$ & ModuleName & ".vba"
                End If
            End If
        Next i
    End With
End Sub
