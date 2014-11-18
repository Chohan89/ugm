Attribute VB_Name = "UGM"
Public Sub main()
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

Public Sub CreateMasterSheetAction(control As IRibbonControl)
    Call CreateMasterSheet
End Sub

Sub CreateMeetingAction(control As IRibbonControl)
    CreateMeeting
End Sub
