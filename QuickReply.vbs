' ============================================================
' QuickReply Module
' Function: Quick reply to emails with "Hi Sender" + template
' ============================================================

' ============================================================
' [Manage your templates here]
' Format: Templates(0) = "template content", starting from 0
' ============================================================
Public Function GetTemplates() As String()
    Dim t(4) As String
    t(0) = "Received, I will handle it as soon as possible. Thank you!"
    t(1) = "Sure, no problem. Let me confirm on my end and get back to you."
    t(2) = "Thanks for your email! I need to check with the team on this and will follow up shortly."
    t(3) = "Received, currently following up. I will notify you as soon as there is an update."
    t(4) = "Thanks for the reminder! I have taken care of it on my end. Please check."
    GetTemplates = t
End Function

' ============================================================
' Core reply function
' ============================================================
Public Sub QuickReplyMail()
    Dim objMail As MailItem
    Dim objExplorer As Explorer
    Dim objSelection As Selection

    Set objExplorer = Application.ActiveExplorer
    Set objSelection = objExplorer.Selection

    If objSelection.Count = 0 Then
        MsgBox "Please select an email first!", vbExclamation, "Quick Reply"
        Exit Sub
    End If

    If Not TypeOf objSelection.Item(1) Is MailItem Then
        MsgBox "Please select a valid email!", vbExclamation, "Quick Reply"
        Exit Sub
    End If

    Set objMail = objSelection.Item(1)

    frmQuickReply.LoadMail objMail
    frmQuickReply.Show
End Sub

' ============================================================
' Extract sender display name
' e.g. "John Smith <john@xx.com>" -> "John Smith"
' ============================================================
Public Function GetSenderName(objMail As MailItem) As String
    Dim fullName As String
    fullName = objMail.SenderName

    If InStr(fullName, "<") > 0 Then
        fullName = Trim(Left(fullName, InStr(fullName, "<") - 1))
    End If

    GetSenderName = fullName
End Function
