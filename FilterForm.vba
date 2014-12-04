VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FilterForm 
   Caption         =   "Filter"
   ClientHeight    =   3084
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   5844
   OleObjectBlob   =   "FilterForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FilterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BTN_moveAllLeft_Click()

    Dim iCtr As Long

    For iCtr = 0 To Me.ListBox2.ListCount - 1
        Me.ListBox1.AddItem Me.ListBox2.List(iCtr)
    Next iCtr

    Me.ListBox2.Clear
    
    SortData ListBox1

End Sub

Private Sub BTN_moveAllRight_Click()

    Dim iCtr As Long

    For iCtr = 0 To Me.ListBox1.ListCount - 1
        Me.ListBox2.AddItem Me.ListBox1.List(iCtr)
    Next iCtr

    Me.ListBox1.Clear
    
    SortData ListBox2

End Sub

Private Sub BTN_MoveSelectedLeft_Click()

    Dim iCtr As Long

    For iCtr = 0 To Me.ListBox2.ListCount - 1
        If Me.ListBox2.Selected(iCtr) = True Then
            Me.ListBox1.AddItem Me.ListBox2.List(iCtr)
        End If
    Next iCtr

    For iCtr = Me.ListBox2.ListCount - 1 To 0 Step -1
        If Me.ListBox2.Selected(iCtr) = True Then
            Me.ListBox2.RemoveItem iCtr
        End If
    Next iCtr
    
    SortData ListBox1

End Sub

Private Sub BTN_MoveSelectedRight_Click()

    Dim iCtr As Long

    For iCtr = 0 To Me.ListBox1.ListCount - 1
        If Me.ListBox1.Selected(iCtr) = True Then
            Me.ListBox2.AddItem Me.ListBox1.List(iCtr)
        End If
    Next iCtr

    For iCtr = Me.ListBox1.ListCount - 1 To 0 Step -1
        If Me.ListBox1.Selected(iCtr) = True Then
            Me.ListBox1.RemoveItem iCtr
        End If
    Next iCtr
    
    SortData ListBox2

End Sub

Private Sub BTN_ok_Click()
    Unload Me
End Sub

Private Sub SortData(myListBox)
    With myListBox
        For i = 0 To .ListCount - 2
            For j = i + 1 To .ListCount - 1
                If .List(i) > .List(j) Then
                    Temp = .List(j)
                    .List(j) = .List(i)
                    .List(i) = Temp
                End If
            Next j
        Next i
    End With
End Sub

Public Sub SetData(data)

    Me.ListBox1.Clear
    Me.ListBox2.Clear

    Dim iCtr As Long

    For iCtr = LBound(data) To UBound(data)
        If Not data(iCtr) = "" Then
            Me.ListBox1.AddItem data(iCtr)
        End If
    Next iCtr
    
    SortData ListBox1

End Sub
