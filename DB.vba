Attribute VB_Name = "DB"
' ADO / (DAO)

Function GetAreaCodeDB() As Collection
    Dim database As database
    Dim rst As Recordset
    
    Dim col As Collection
    Set col = New Collection

    Set database = OpenDatabase(Application.UserLibraryPath & "\UserGroupManager.mdb")
    Set rst = database.OpenRecordset("Geodata")

    Do Until rst.EOF
        Dim key As String
        Dim geo As Geodata
        Set geo = New Geodata
        key = rst.Fields("Country") & rst.Fields("AreaCode")
        geo.State = rst.Fields("State")
        geo.Region = rst.Fields("Region")
        col.Add geo, key
        rst.MoveNext
    Loop

    'close the objects
    rst.Close
    database.Close

    'destroy the variables
    Set rst = Nothing
    Set database = Nothing
    
    Set GetAreaCodeDB = col
End Function

Public Function Contains(col As Collection, key As Variant) As Boolean
    Dim obj As Variant
    On Error GoTo err
        Contains = True
        obj = col(key).State
        Exit Function
err:
        Contains = False
End Function
