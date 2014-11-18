Attribute VB_Name = "DaveUGM"
Public sName As String
Public cName As Range
Public RC As Integer

Sub main()
Call Copy
        If sName = "False" Then
        MsgBox "Canceled"
        Exit Sub
        End If
Call State
Call DelnMoveCol
Call ColHeaderConfig
Call PhoneConfig
End Sub

Sub MasterSheet()
RC = 0                               'Reset Return Code
Call OnlyCopy
        If RC = 8 Then
            Exit Sub
        End If
Call Cleanup
Call Propername
Call EmailCleanup
Call Countries

'Note: Master Sheet will be Sorted by who has most OPEN ISSUES (Opened Issues - Closed Issues)'

End Sub

Function WksExists(wksName As String) As Boolean
    On Error Resume Next
    WksExists = CBool(Len(Worksheets(wksName).Name) > 0)
End Function

Function ValCol(ColNam As String) As Boolean
    On Error Resume Next
    Lastcol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column      'find last column
    Set cName = Range(Cells(1, 1), Cells(1, Lastcol)).Find(ColNam)          'find target column
    lastrow = Cells(Application.ActiveSheet.Rows.Count, cName.Column).End(xlUp).Row
    Set cName = Range(Cells(cName.Row, cName.Column), Cells(lastrow, cName.Column))
    'MsgBox (x.Address)
    If cName Is Nothing Then
        ValCol = False
        MsgBox (ColNam + " column not found")
    Else
        ValCol = True
    End If
End Function

Function RenameCell(Target As String, Chnge As String)
On Error Resume Next
    Lastcol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column      'find last column
    Set Trget = Range(Cells(1, 1), Cells(1, Lastcol)).Find(Target)          'find target column
    Columns(Trget.Column).Select
    ActiveCell.Value = Chnge
End Function

Function MoveCol(Anchor As String, ColNam As String, Colour As Integer)
On Error Resume Next
    Lastcol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column      'find last column
    Set Anch = Range(Cells(1, 1), Cells(1, Lastcol)).Find(Anchor)           'find Anchor column
    Columns(Anch.Column).Select
    Selection.Insert shift:=xlToRight, copyorigin:=xlformatfromleftofabove
    Set GetCell = Range(Cells(1, 1), Cells(1, Lastcol)).Find("Column1")
    Range(GetCell.Address).Select
        If Not Colour <> "" Then
        ActiveCell.Interior.Color = Colour
        End If
        If ColNam = "sName" Then
            ActiveCell.Value = sName
        End If
    ActiveCell.Value = ColNam
End Function
Sub OnlyCopy()
Dim i As Integer

If WksExists("Master Sheet") Then
    MsgBox "Master Sheet already exists"
    RC = 8
    Exit Sub
End If

i = ActiveWorkbook.Worksheets.Count
Worksheets(1).Copy After:=Worksheets(i)
ActiveSheet.Name = "Master Sheet"
End Sub

Function Copy()
sName = ""

Dim i As Integer
i = ActiveWorkbook.Worksheets.Count
 
Do While sName = ""
sName = Application.InputBox("Enter a name for the new sheet, preferably the site of the meeting" & vbLf & "(Name should include atleast one letter)", "Create Site Worksheet")
If WksExists(sName) Then
    MsgBox "Duplicate Sheet Name"
    sName = ""
End If

If sName = "False" Then
    Exit Function
End If

Loop

If Not WksExists("Master Sheet") Then
    MsgBox "Create a Master Sheet First"
    sName = "False"
    Exit Function
Else
Worksheets("Master Sheet").Copy After:=Worksheets(i)
ActiveSheet.Name = sName
End If

End Function

Sub Cleanup()                                                   'We dont use ValCol function because it searches row 1, whereas Issues Opened can be anywhere
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
Columns(1).Delete             'Delete uneeded Cols'
Next y
End If

End Sub

Sub Countries()

'Filter for US and Canada'

'Dim Ccell As Range
'Set Ccell = ActiveSheet.Cells.Find("Country")

If Not ValCol("Country") Then
    Exit Sub
Else: Set Ccell = cName
End If

    ActiveSheet.Range(Ccell.Address).AutoFilter Field _
    :=Ccell.Column, Criteria1:="=CA", Operator:=xlOr, Criteria2:="=US"
   
End Sub

Sub EmailCleanup()

'Remove Duplicates via Email Column

Dim Ecell As Range
Dim rLastCell As Range

Set rLastCell = ActiveSheet.Cells.Find(What:="*", After:=ActiveSheet.Cells(1, 1), LookIn:=xlFormulas, LookAt:= _
xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)                                        'find last column'

Columns(rLastCell.Column).Select                                                                                        'create temp column'
Selection.Insert shift:=xlToRight, copyorigin:=xlFormatFromRightOrAbove

If Not ValCol("Column1") Then
    Exit Sub
Else: Set Ecell = cName
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
    
    
'Set Ecell = ActiveSheet.Cells.Find("Email")
If Not ValCol("Email") Then
    Exit Sub
Else: Set Ecell = cName
End If
    ActiveSheet.Range(Ecell.Address).RemoveDuplicates Columns:=Ecell.Column, Header:=xlYes

End Sub

Sub State()

'Filter States

Dim Scell As Range

If Not ValCol("State/Region") Then
    Exit Sub
Else: Set Scell = cName
End If

    ActiveSheet.Range(Scell.Address).AutoFilter Field _
        :=Scell.Column, Criteria1:=Array("DC", "DE", "MD", "NJ", "NY", "PA"), Operator:= _
        xlFilterValues
End Sub
Sub DelnMoveCol()


Dim i As Long
Dim k As Long

Set Acell = ActiveSheet.Cells.Find("City")

For i = 1 To Acell.Column
    
    If Cells(1, i).Value <> "" And Not IsEmpty(Cells(i, 1).Value) Then        'If cells are not empty then color in Blue'
    Cells(1, i).Interior.Color = RGB(0, 112, 192)                             'this is before we make new columns in red'
    Range(Acell.Address).Borders(xlEdgeLeft).LineStyle = xlContinuous
    End If
Next i


'Delete unused cols and Add new cols

Dim Gcell As Range
Dim Pcell As Range
 
 
Set Gcell = ActiveSheet.Cells.Find("Backlog")

    Columns(Gcell.Column).EntireColumn.Delete              'Delete Backlog column'

Set Gcell = ActiveSheet.Cells.Find("Site Name")
Set Pcell = ActiveSheet.Cells.Find("Phone")

Columns(Gcell.Column).Select                               'Move Site Name to the left of Phone'
    Selection.Cut
    Columns(Pcell.Column).Select
    Selection.Insert shift:=xlToRight
    
Set Gcell = ActiveSheet.Cells.Find("Phone")                'Move Phone to the right spot'
Set Pcell = ActiveSheet.Cells.Find("ZIP code")

Columns(Gcell.Column).Select
    Selection.Cut
    Columns(Pcell.Column).Select
    Selection.Insert shift:=xlToRight
    
'create new column and name it Attend'
Call MoveCol("Email", "Attend", 255)
    
'create new column and name it Office Site'
Call MoveCol("Email", sName, 255)
    
'create new column and name it Response Details'
Call MoveCol("Email", "Response Details", 255)

'create new column and name it P'
Call MoveCol("Email", "P", 255)

'create new column and name it Area'
Call MoveCol("Site ID", "Area", 255)
    
'create new column and name it Area Code State'
Call MoveCol("Site ID", "Area Code State", 255)

'create new column and name it Local'
Call MoveCol("Site ID", "Local", 255)
    
End Sub

Sub ColHeaderConfig()
'Find Issues Opened Col to rename it'
Call RenameCell("Issues Opened", "OPN")
    
'Find Issues Closed Col to rename it'
Call RenameCell("Issues Closed", "CLOSE")

'Find Release Col to rename it'
Call RenameCell("Release", "REL")

'Find Country Col to rename it'
Call RenameCell("Country", "CO")

Call AutoFitCol

Rows(1).RowHeight = 53

End Sub

Sub SortCountryPhone() 'NOT BEING USED AS OF YET
    
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

Sub PhoneConfig()
Call ResetFilters

If Not ValCol("Phone") Then
    Exit Sub
Else: Set Phone = cName
End If

If Not ValCol("Area") Then
    Exit Sub
Else: Set areanum = cName
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
      'Cells(a.Row, areanum.Column).Value = g
      area_temp(a.Row, 1) = g
  
  ElseIf Len(a) > 10 Then
        a.Cells.Interior.Color = RGB(255, 255, 0)                         'Yellow = >10 digit length
        tmpnum = Right(a.Cells, 10)
        g = Left(tmpnum, 3)
       ' Cells(a.Row, areanum.Column).Value = g
        area_temp(a.Row, 1) = g
  End If

Next a

areanum.Value = area_temp

End Sub
Sub AutoFitCol()

Lastcol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column

For i = 1 To Lastcol
        Columns(i).EntireColumn.AutoFit                                              'Autofit all columns'
Next i

End Sub


