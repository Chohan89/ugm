Attribute VB_Name = "DaveUGM"
Public sName As String

Sub MasterSheet()

Call OnlyCopy
Call Cleanup
Call Propername
Call EmailCleanup
Call FilterCountries

'Note: Master Sheet will be Sorted by who has most OPEN ISSUES (Opened Issues - Closed Issues)'

End Sub

Public Function WksExists(wksName As String) As Boolean
    On Error Resume Next
    WksExists = CBool(Len(Worksheets(wksName).Name) > 0)
End Function

Function ValCol(ColNam As String) As Boolean
    On Error Resume Next
    Set x = ActiveSheet.Cells.Find(ColNam)
    If x Is Nothing Then
        ValCol = False
        MsgBox (ColNam + " column not found")
    Else
        ValCol = True
    End If
End Function

Public Sub OnlyCopy()
    Dim i As Integer
    
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
    
    Else
        Worksheets(2).Copy After:=Worksheets(i)
        ActiveSheet.Name = sName
    End If

End Function

Sub Cleanup()
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

Sub FilterCountries()

    'Filter for US and Canada'
    
    Dim Ccell As Range
    Set Ccell = ActiveSheet.Cells.Find("Country")
    
    If Not ValCol("Country") Then
        Exit Sub
    End If
    
    ActiveSheet.Range(Ccell.Address).AutoFilter Field _
        :=Ccell.Column, Criteria1:="=CA", Operator:=xlOr, Criteria2:="=US"
   
End Sub

Sub EmailCleanup()
    
    'Remove Duplicates via Email Column
    
    Dim Ecell As Range
    
    Set Opn = ActiveSheet.Cells.Find("Issues Opened")
    Set Clse = ActiveSheet.Cells.Find("Issues Closed")
    
    If Not ValCol("Issues Opened") Then
        Exit Sub
    End If
    
    If Not ValCol("Issues Closed") Then
        Exit Sub
    End If
    
    Dim rLastCell As Range
    
    Set rLastCell = ActiveSheet.Cells.Find(What:="*", After:=ActiveSheet.Cells(1, 1), LookIn:=xlFormulas, LookAt:= _
    xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)                                        'find last column'
    
    Columns(rLastCell.Column).Select                                                                                        'create temp column'
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrAbove
    
    Set Ecell = ActiveSheet.Cells.Find("Column1")
        If Ecell.Column <> 16 Then
            MsgBox ("New temp column is not at position 16")
            Exit Sub
        End If
        
        
        Range("p2").FormulaR1C1 = "=[@[Issues Opened]]-[@[Issues Closed]]"
        
        Header = Cells(Ecell.Row, Ecell.Column).End(xlToLeft).Column
        lastheader = Cells(Ecell.Row, Application.ActiveSheet.Columns.Count).End(xlToLeft).Column
        
        Range(Cells(Ecell.Row, Header), Cells(Ecell.Row, lastheader)).Sort key1:=Ecell, order1:=xlDescending
        
        Columns(Ecell.Column).EntireColumn.Delete                                                                            'Delete Temp column'
        
        
    Set Ecell = ActiveSheet.Cells.Find("Email")
        ActiveSheet.Range(Ecell.Address).RemoveDuplicates Columns:=Ecell.Column, Header:=xlYes

End Sub

Public Sub State()

    'Filter States
    
    Dim Scell As Range
    
    Set Scell = ActiveSheet.Cells.Find("State/Region")

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
        Selection.Insert Shift:=xlToRight
        
    Set Gcell = ActiveSheet.Cells.Find("Phone")                'Move Phone to the right spot'
    Set Pcell = ActiveSheet.Cells.Find("ZIP code")
    
    Columns(Gcell.Column).Select
        Selection.Cut
        Columns(Pcell.Column).Select
        Selection.Insert Shift:=xlToRight
        
    Set Gcell = ActiveSheet.Cells.Find("Email")
    
        Columns(Gcell.Column).Select                                                 'create new column and name it Attend'
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set Pcell = ActiveSheet.Cells.Find("Column1")
        Range(Pcell.Address).Select
        ActiveCell.Interior.Color = 255
        ActiveCell.Value = "Attend"
        
        
        Columns(Gcell.Column).Select                                                 'create new column and name it Office Site'
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set Pcell = ActiveSheet.Cells.Find("Column1")
        Range(Pcell.Address).Select
        ActiveCell.Value = sName                                                     'get value from another sub (copy)'
        
        Columns(Gcell.Column).Select                                                 'create new column and name it Response Details'
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set Pcell = ActiveSheet.Cells.Find("Column1")
        Range(Pcell.Address).Select
        ActiveCell.Value = "Response Details"
        
        Columns(Gcell.Column).Select                                                 'create new column and name it P'
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set Pcell = ActiveSheet.Cells.Find("Column1")
        Range(Pcell.Address).Select
        ActiveCell.Value = "P"
        
    Set Gcell = ActiveSheet.Cells.Find("Site ID")                                     'New column for reference (before it was Email)'
    
        Columns(Gcell.Column).Select                                                 'create new column and name it Area'
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set Pcell = ActiveSheet.Cells.Find("Column1")
        Range(Pcell.Address).Select
        ActiveCell.Interior.Color = 255
        ActiveCell.Value = "Area"
        
        Columns(Gcell.Column).Select                                                 'create new column and name it Area Code State'
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set Pcell = ActiveSheet.Cells.Find("Column1")
        Range(Pcell.Address).Select
        ActiveCell.Interior.Color = 255
        ActiveCell.Value = "Area Code State"
        
        Columns(Gcell.Column).Select                                                 'create new column and name it Local'
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set Pcell = ActiveSheet.Cells.Find("Column1")
        Range(Pcell.Address).Select
        ActiveCell.Interior.Color = 255
        ActiveCell.Value = "Local"
        
    
End Sub

Sub ColHeaderConfig()
    Dim Acell As Range
    Dim i As Long
    Dim k As Long
    
    
    Set Acell = ActiveSheet.Cells.Find("Issues Opened")                                'Find Issues Opened Col to rename it'
        Columns(Acell.Column).Select
        ActiveCell.Value = "OPN"
        
    Set Acell = ActiveSheet.Cells.Find("Issues Closed")                                'Find Issues Closed Col to rename it'
        Columns(Acell.Column).Select
        ActiveCell.Value = "CLOSE"
        
    Set Acell = ActiveSheet.Cells.Find("Release")                                       'Find Release Col to rename it'
        Columns(Acell.Column).Select
        ActiveCell.Value = "REL"
        
    Set Acell = ActiveSheet.Cells.Find("Country")                                        'Find Country Col to rename it'
        Columns(Acell.Column).Select
        ActiveCell.Value = "CO"
    
    
    Set Acell = ActiveSheet.Cells.Find("City")
    
    Call AutoFitCol
    
    Rows(1).RowHeight = 53

End Sub

Sub PhoneConfig()

    Set Phone = ActiveSheet.Cells.Find("Phone")
    Set areanum = ActiveSheet.Cells.Find("Area")
    
    If Not ValCol("Phone") Then
        Exit Sub
    End If
    
    If Not ValCol("Area") Then
        Exit Sub
    End If
    
    lastphone = Cells(Application.ActiveSheet.Rows.Count, Phone.Column).End(xlUp).Row
    Dim rng As Range, cell As Range
    Dim firstcell As Variant, lastcell As Variant, firstcella As Variant
    Set firstcell = Cells(Phone.Row, Phone.Column)
    Set lastcell = Cells(lastphone, Phone.Column)
    Set firstcella = Phone.Offset(1, 0)
    Set temprng = Range(firstcella, lastcell)
    Dim a As Range
    Dim tmpnum As String
    Dim g As String
    For Each a In temprng
           
      If Not IsNumeric(a.Cells) Then                                      'Red = non Numeric AND OR < 10 digit length
          a.Cells.Interior.Color = RGB(255, 0, 0)
       
      ElseIf Len(a.Cells) < 10 Then
          a.Cells.Interior.Color = RGB(255, 0, 0)
          
      ElseIf Len(a) = 10 Then
          a.Cells.Interior.Color = RGB(0, 255, 0)                           'Green = 10 digit length
          g = Left(a.Cells, 3)
          Cells(a.Row, areanum.Column).Value = g
      
      ElseIf Len(a) > 10 Then
          a.Cells.Interior.Color = RGB(255, 255, 0)                         'Yellow = >10 digit length
          tmpnum = Right(a.Cells, 10)
          g = Left(tmpnum, 3)
          Cells(a.Row, areanum.Column).Value = g
      End If
    
    Next a
        
    End Sub
    
    Sub AutoFitCol()
    
    Lastcol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    
    For i = 1 To Lastcol
            Columns(i).EntireColumn.AutoFit                                              'Autofit all columns'
    Next i

End Sub

