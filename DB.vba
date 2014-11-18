Attribute VB_Name = "DB"
' ADO / (DAO)

Function GetAreaCodeDB() As Collection
    Dim database As database
    Dim rst As Recordset
    
    Set GetAreaCodeDB = New Collection

    Set database = OpenDatabase(Application.UserLibraryPath & "\UserGroupManager.mdb")
    Set rst = database.OpenRecordset("State")

    Do Until rst.EOF
        dictionary.Add rst.Fields("Name"), rst.Fields("Country") & rst.Fields("AreaCode")
        rst.MoveNext
    Loop

    'close the objects
    rst.Close
    database.Close

    'destroy the variables
    Set rst = Nothing
    Set database = Nothing
End Function
