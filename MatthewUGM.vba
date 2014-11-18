Attribute VB_Name = "MatthewUGM"
Public Sub Propername()
    Set fname = Cells.Find(What:="First Name")
    Set lname = Cells.Find(What:="Last Name")
    Set thename = Cells.Find(What:="Full Name")
    If fname Is Nothing Then
        MsgBox ("No column found with header ""Full Name""")
        Exit Sub
    End If
    If lname Is Nothing Then
        MsgBox ("No column found with header ""Last Name""")
        Exit Sub
    End If
    If thename Is Nothing Then
        fname.EntireColumn.Insert
        Cells(fname.Column - 1)(fname.Row).Value = "Full Name"
    End If
    lastfname = Cells(Application.ActiveSheet.Rows.Count, fname.Column).End(xlUp).Row
    lastlname = Cells(Application.ActiveSheet.Rows.Count, lname.Column).End(xlUp).Row
    Set fnamerange = Range(Cells(fname.Row, fname.Column), Cells(lastfname, fname.Column))
    Set lnamerange = Range(Cells(lname.Row, lname.Column), Cells(lastlname, lname.Column))
    Set toProper = Union(fnamerange, lnamerange)
    Dim temp_range As Variant
    temp_range = toProper.Value
    For i = LBound(temp_range) To UBound(temp_range)
        temp_range(i, 1) = StrConv(temp_range(i, 1), vbProperCase)
        temp_range(i, 2) = StrConv(temp_range(i, 2), vbProperCase)
    Next i
    toProper.Value = temp_range
    Cells(fname.Column - 1)(fname.Row + 1).Value = "=CONCATENATE([@[First Name]],"" "",[@[Last Name]])"
    Range(Cells(1, fname.Column - 1), Cells(Application.ActiveSheet.Rows.Count, fname.Column - 1)).Columns.AutoFit
    
    'RestoreListObjectFilters wks, varFilterCache
End Sub

Public Sub ResetFilters()
    If ActiveSheet.FilterMode = True Then
        ActiveSheet.ShowAllData
    End If
End Sub
