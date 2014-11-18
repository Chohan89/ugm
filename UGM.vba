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
    Call MasterSheet
End Sub

Sub CreateMeetingAction(control As IRibbonControl)
End Sub