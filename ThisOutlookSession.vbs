' ============================================================
' ThisOutlookSession
' Register / unregister context menu entry
' ============================================================

Private Sub Application_Startup()
    Call AddContextMenu
End Sub

Private Sub Application_Quit()
    Call RemoveContextMenu
End Sub

Private Sub AddContextMenu()
    Dim objBar As Office.CommandBar
    Dim objBtn As Office.CommandBarButton

    Call RemoveContextMenu

    For Each objBar In Application.ActiveExplorer.CommandBars
        If objBar.Name = "Context Menu" Then
            Set objBtn = objBar.Controls.Add(Type:=msoControlButton, Temporary:=True)
            objBtn.Caption = "Quick Reply"
            objBtn.OnAction = "QuickReplyMail"
            objBtn.BeginGroup = True
        End If
    Next objBar
End Sub

Private Sub RemoveContextMenu()
    Dim objBar As Office.CommandBar
    Dim objCtrl As Office.CommandBarControl

    On Error Resume Next
    For Each objBar In Application.ActiveExplorer.CommandBars
        For Each objCtrl In objBar.Controls
            If objCtrl.Caption = "Quick Reply" Then
                objCtrl.Delete
            End If
        Next objCtrl
    Next objBar
    On Error GoTo 0
End Sub
