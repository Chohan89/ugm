Attribute VB_Name = "DaveUGM"
Public Sub CreateMeeting()
    If Not CopyForMeeting() Then
        MsgBox "Canceled"
        Exit Sub
    End If
    Call ResetFilters
    Call DeleteAndMoveColumns
    Call ColumnHeaderConfig
    Call PhoneConfig
    Call FilterCountries
    Call FilterState
End Sub

Public Sub CreateMasterSheet()
    Set oldSheet = ActiveSheet
    If Not CopyForMaster() Then
        MsgBox "Canceled"
        Exit Sub
    End If
    Application.DisplayAlerts = False
    oldSheet.Delete
    Application.DisplayAlerts = True
    Call ResetFilters
    Call Cleanup
    Call Propername
    Call EmailCleanup
    
    'Note: Master Sheet will be Sorted by who has most OPEN ISSUES (Opened Issues - Closed Issues)'
End Sub

Private Function WksExists(wksName As String) As Boolean
    On Error Resume Next
    WksExists = CBool(Len(Worksheets(wksName).Name) > 0)
End Function

Private Function FindColumn(ColNam As String) As Range
    On Error Resume Next
    Lastcol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column      'find last column
    Dim col As Range
    Set col = Range(Cells(1, 1), Cells(1, Lastcol)).Find(ColNam)          'find target column
    lastrow = Cells(Application.ActiveSheet.Rows.Count, col.Column).End(xlUp).Row
    Set FindColumn = Range(Cells(col.Row, col.Column), Cells(lastrow, col.Column))
    If FindColumn Is Nothing Then
        MsgBox (ColNam + " column not found")
    End If
End Function

Private Function RenameCell(Target As String, Chnge As String)
    On Error Resume Next
    FindColumn(Target)(1)(1).Value = Chnge
End Function

Private Function CreateColumn(Anchor As String, ColNam As String, Colour As Integer)
    On Error Resume Next
    Lastcol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column      'find last column
    Set Anch = Range(Cells(1, 1), Cells(1, Lastcol)).Find(Anchor)           'find Anchor column
    Columns(Anch.Column).Select
    Selection.Insert shift:=xlToRight, copyorigin:=xlformatfromleftofabove
    If Not Colour <> "" Then
        ActiveCell.Interior.Color = Colour
    End If
    ActiveCell.Value = ColNam
End Function

Private Function CopyForMaster() As Boolean
    Dim i As Integer
    
    If WksExists("Master Sheet") Then
        MsgBox "Master Sheet already exists"
        CopyForMaster = False
        Exit Function
    End If
    
    i = ActiveWorkbook.Worksheets.Count
    Worksheets(1).Copy After:=Worksheets(i)
    ActiveSheet.Name = "Master Sheet"
    CopyForMaster = True
End Function

Private Function CopyForMeeting() As Boolean
    Dim sName$
    
    If Not WksExists("Master Sheet") Then
        Call CopyForMaster
    End If
     
    Do While sName$ = ""
        sName$ = Application.InputBox("Enter a name for the new sheet, preferably the site of the meeting" & vbLf & "(Name should include atleast one letter)", "Create Site Worksheet")

        If WksExists(sName$) Then
            MsgBox "Duplicate Sheet Name"
            sName$ = ""
        End If
        
        If sName$ = "False" Then
            CopyForMeeting = False
            Exit Function
        End If
    Loop
    
    Worksheets("Master Sheet").Copy After:=Worksheets(ActiveWorkbook.Worksheets.Count)
    ActiveSheet.Name = sName$
    CopyForMeeting = True
End Function

Private Sub Cleanup() 'We dont use FindColumn function because it searches row 1, whereas Issues Opened can be anywhere
    Dim x As Integer
    Dim y As Integer
    
    Dim Zcell As Range
    
    Set Zcell = ActiveSheet.Cells.Find("Issues Opened ")
    
    If IsNumeric(Zcell.Row) Then
        For x = 1 To Zcell.Row + 1
            Rows(1).Delete            'Delete uneeded Rows'
        Next x
    End If
    
    Set Zcell = ActiveSheet.Cells.Find("Issues Opened")
    
    If IsNumeric(Zcell.Column) Then
        For y = 2 To Zcell.Column
            Columns(1).Delete             'Delete uneeded Columns'
        Next y
    End If
End Sub

Private Sub FilterCountries()
    'Filter for US and Canada'
    
    Set Ccell = FindColumn("CON")
    If Ccell Is Nothing Then
        Exit Sub
    End If

    ActiveSheet.Range(Ccell.Address).AutoFilter Field _
        :=Ccell.Column, Criteria1:="=CA", Operator:=xlOr, Criteria2:="=US"
End Sub

Private Sub EmailCleanup()
    'Remove Duplicates via Email Column
    
    Dim Ecell As Range
    Dim rLastCell As Range
    
    Set rLastCell = ActiveSheet.Cells.Find(What:="*", After:=ActiveSheet.Cells(1, 1), LookIn:=xlFormulas, LookAt:= _
    xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)                                        'find last column'
    
    Columns(rLastCell.Column).Select                                                                                        'create temp column'
    Selection.Insert shift:=xlToRight, copyorigin:=xlFormatFromRightOrAbove
    
    Set Ecell = FindColumn("Column1")
    If Ecell Is Nothing Then
        Exit Sub
    End If
    
    If Ecell.Column <> 16 Then
        MsgBox ("New temp column is not at position 16")
        Exit Sub
    End If
        
    Range("p2").FormulaR1C1 = "=[@[Issues Opened]]-[@[Issues Closed]]"
    'Range(lstcol.Offset(0, 1)).FormulaR1C1 = "=[@[Issues Opened]]-[@[Issues Closed]]"
    Header = Cells(Ecell.Row, Ecell.Column).End(xlToLeft).Column
    lastheader = Cells(Ecell.Row, Application.ActiveSheet.Columns.Count).End(xlToLeft).Column
    
    Range(Cells(Ecell.Row, Header), Cells(Ecell.Row, lastheader)).Sort key1:=Ecell, order1:=xlDescending
    
    Columns(Ecell.Column).EntireColumn.Delete                                                                            'Delete Temp column'

    Set Ecell = FindColumn("Email")
    If Ecell Is Nothing Then
        Exit Sub
    End If
    ActiveSheet.Range(Ecell.Address).RemoveDuplicates Columns:=Ecell.Column, Header:=xlYes
End Sub

Private Sub FilterState()
    Dim Scell As Range

    Set Scell = FindColumn("State/Region")
    If Scell Is Nothing Then
        Exit Sub
    End If
    
    ActiveSheet.Range(Scell.Address).AutoFilter Field _
        :=Scell.Column, Criteria1:=Array("DC", "DE", "MD", "NJ", "NY", "PA"), Operator:= _
        xlFilterValues
End Sub

Private Sub DeleteAndMoveColumns()
    Dim i As Long
    Dim k As Long
    
    Set Acell = FindColumn("City")
    If Acell Is Nothing Then
        Exit Sub
    End If
    
    For i = 1 To Acell.Column
        
        If Cells(1, i).Value <> "" And Not IsEmpty(Cells(i, 1).Value) Then        'If cells are not empty then color in Blue'
        Cells(1, i).Interior.Color = RGB(0, 112, 192)                             'this is before we make new columns in red'
        Range(Acell.Address).Borders(xlEdgeLeft).LineStyle = xlContinuous
        End If
    Next i

    'Delete unused cols and Add new cols
    
    Dim Gcell As Range
    Dim Pcell As Range

    Set Gcell = FindColumn("Backlog")
    If Gcell Is Nothing Then
        Exit Sub
    End If
    
    Columns(Gcell.Column).EntireColumn.Delete              'Delete Backlog column'
    
    Set Gcell = FindColumn("Site Name")
    If Gcell Is Nothing Then
        Exit Sub
    End If
    
    Set Pcell = FindColumn("Phone")
    If Pcell Is Nothing Then
        Exit Sub
    End If
    
    Columns(Gcell.Column).Select                               'Move Site Name to the left of Phone'
    Selection.Cut
    Columns(Pcell.Column).Select
    Selection.Insert shift:=xlToRight
        
    Set Gcell = FindColumn("ZIP Code")
    If Gcell Is Nothing Then
        Exit Sub
    End If
    
    Columns(Pcell.Column).Select
    Selection.Cut
    Columns(Gcell.Column).Select
    Selection.Insert shift:=xlToRight, copyorigin:=xlformatfromleftofabove
        
    'create new column and name it Attend'
    Call CreateColumn("Email", "Attend", 255)
        
    'create new column and name it Office Site'
    Call CreateColumn("Email", ActiveSheet.Name, 255)
        
    'create new column and name it Response Details'
    Call CreateColumn("Email", "Response Details", 255)
    
    'create new column and name it P'
    Call CreateColumn("Email", "P", 255)
    
    'create new column and name it Area'
    Call CreateColumn("Site ID", "Area", 255)
        
    'create new column and name it Area Code State'
    Call CreateColumn("Site ID", "Area Code State", 255)
    
    'create new column and name it Local'
    Call CreateColumn("Site ID", "Local", 255)
End Sub

Private Sub ColumnHeaderConfig()
    'Find Issues Opened Col to rename it'
    Call RenameCell("Issues Opened", "OPN")
        
    'Find Issues Closed Col to rename it'
    Call RenameCell("Issues Closed", "CLOSE")
    
    'Find Release Col to rename it'
    Call RenameCell("Release", "REL")
    
    'Find Country Col to rename it'
    Call RenameCell("Country", "CON")
    
    Call AutoFitColumns
    
    Rows(1).RowHeight = 53
End Sub

Private Sub SortCountryPhone() 'NOT BEING USED AS OF YET
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

Private Sub PhoneConfig()
    Set Phone = FindColumn("Phone")
    If Phone Is Nothing Then
        Exit Sub
    End If
    
    Set areanum = FindColumn("Area")
    If areanum Is Nothing Then
        Exit Sub
    End If
    
    lastphone = Cells(Application.ActiveSheet.Rows.Count, Phone.Column).End(xlUp).Row
    Set firstcell = Cells(Phone.Row, Phone.Column)
    Set lastcell = Cells(lastphone, Phone.Column)
    Set temprng = Range(firstcell, lastcell)
    Dim a As Range
    Dim tmpnum As String
    Dim g As String
    area_temp = areanum.Value

    For Each a In temprng
        If Not IsNumeric(a.Cells) Then                                      'Red = non Numeric AND OR < 10 digit length
            a.Cells.Interior.Color = RGB(255, 0, 0)
        ElseIf Len(a.Cells) < 10 Then
            a.Cells.Interior.Color = RGB(255, 0, 0)
        ElseIf Len(a) = 10 Then
            a.Cells.Interior.Color = RGB(0, 255, 0)                           'Green = 10 digit length
            g = Left(a.Cells, 3)
            area_temp(a.Row, 1) = g
        ElseIf Len(a) > 10 Then
            a.Cells.Interior.Color = RGB(255, 255, 0)                         'Yellow = >10 digit length
            tmpnum = Right(a.Cells, 10)
            g = Left(tmpnum, 3)
            area_temp(a.Row, 1) = g
        End If
    Next a

    areanum.Value = area_temp
End Sub

Private Sub AutoFitColumns()
    Lastcol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column

    For i = 1 To Lastcol
        Columns(i).EntireColumn.AutoFit                                              'Autofit all columns'
    Next i
End Sub


