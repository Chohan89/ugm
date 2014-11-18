Attribute VB_Name = "Deprecated"
'Callback for NewSheetBtn onAction
Private Sub NewSheetAction(control As IRibbonControl)
    Call Copy
    Call Cleanup
End Sub

'Callback for ProperNameBtn onAction
Private Sub ProperNameAction(control As IRibbonControl)
    Call Propername
End Sub

Private Function SelectFile() As String
    Dim intChoice As Integer
    Dim strPath As String
    'only allow the user to select one file
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    'make the file dialog visible to the user
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    'determine what choice the user made
    If intChoice <> 0 Then
        'get the file path selected by the user
        strPath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
        'return the file path
        SelectFile = strPath
    End If
End Function

Private Sub OpenFile()
    Workbooks.Open SelectFile()
End Sub

Private Sub GetXML()
    Dim myFile As String, text As String, textline As String, posLat As Integer, posLong As Integer
    Open Application.UserLibraryPath & "UserGroupManager.exportedUI" For Input As #1
    Do Until EOF(1)
    Line Input #1, textline
        text = text & textline
    Loop
    Close #1
    ActiveProject.SetCustomUI (ribbonXml)
End Sub

Private Sub SortCountryPhone()

    Set Country = Cells.Find(What:="Country")
    Set Phone = Cells.Find(What:="Phone")
    If Country Is Nothing Then
        MsgBox ("No column found with header ""Country""")
        Exit Sub
    End If
    If Phone Is Nothing Then
        MsgBox ("No column found with header ""Phone""")
        Exit Sub
    End If
    firstHeader = Cells(Country.Row, Country.Column).End(xlToLeft).Column
    lastheader = Cells(Country.Row, Application.ActiveSheet.Columns.Count).End(xlToLeft).Column
    Range(Cells(Country.Row, firstHeader), Cells(Country.Row, lastheader)).Sort key1:=Country, key2:=Phone, order1:=xlAscending, Header:=xlYes
    'Application.ActiveSheet.Rows.Count

End Sub

Private Sub main()
    Call Copy
    If sName = "False" Then
        MsgBox "Canceled"
        Exit Sub
    End If
    Call Cleanup
    Call Propername
    Call EmailCleanup
    Call Countries
    Call State
    Call DelnMoveCol
    Call ColHeaderConfig
    'Call SortCountryPhone
End Sub
