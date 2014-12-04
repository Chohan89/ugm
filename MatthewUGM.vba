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
End Sub

Public Sub ResetFilters()
    If ActiveSheet.FilterMode = True Then
        ActiveSheet.ShowAllData
    End If
End Sub

Public Sub ShowForm()
    Set myFrm = New FilterForm
    myArr = UniqueItems(FindColumn("Area Code State"), False)
    myArr(0) = "" 'Removes header column
    myFrm.SetData (myArr)
    myFrm.Show
End Sub

Function UniqueItems(ArrayIn, Optional Count As Variant) As Variant
'   Accepts an array or range as input
'   If Count = True or is missing, the function returns the number of unique elements
'   If Count = False, the function returns a variant array of unique elements
    Dim Unique() As Variant ' array that holds the unique items
    Dim Element As Variant
    Dim i As Integer
    Dim FoundMatch As Boolean
'   If 2nd argument is missing, assign default value
    If IsMissing(Count) Then Count = True
'   Counter for number of unique elements
    NumUnique = 0
'   Loop thru the input array
    For Each Element In ArrayIn
        FoundMatch = False
'       Has item been added yet?
        For i = 1 To NumUnique
            If Element = Unique(i) Then
                FoundMatch = True
                Exit For '(exit loop)
            End If
        Next i
AddItem:
'       If not in list, add the item to unique list
        If Not FoundMatch And Not IsEmpty(Element) Then
            NumUnique = NumUnique + 1
            ReDim Preserve Unique(NumUnique)
            Unique(NumUnique) = Element
        End If
    Next Element
'   Assign a value to the function
    If Count Then UniqueItems = NumUnique Else UniqueItems = Unique
End Function

