' ============================================================
' frmQuickReply - Template selection form
' ============================================================

Private m_Mail As MailItem

Public Sub LoadMail(objMail As MailItem)
    Set m_Mail = objMail
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Quick Reply - Select Template"
    Me.Width = 400
    Me.Height = 320

    Dim lblHint As MSForms.Label
    Set lblHint = Me.Controls.Add("Forms.Label.1", "lblHint")
    lblHint.Caption = "Select a template to send:"
    lblHint.Left = 10
    lblHint.Top = 10
    lblHint.Width = 360
    lblHint.Height = 20

    Dim lstTemplates As MSForms.ListBox
    Set lstTemplates = Me.Controls.Add("Forms.ListBox.1", "lstTemplates")
    lstTemplates.Left = 10
    lstTemplates.Top = 35
    lstTemplates.Width = 360
    lstTemplates.Height = 180

    Dim btnSend As MSForms.CommandButton
    Set btnSend = Me.Controls.Add("Forms.CommandButton.1", "btnSend")
    btnSend.Caption = "Send Reply"
    btnSend.Left = 10
    btnSend.Top = 225
    btnSend.Width = 100
    btnSend.Height = 30

    Dim btnCancel As MSForms.CommandButton
    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel")
    btnCancel.Caption = "Cancel"
    btnCancel.Left = 120
    btnCancel.Top = 225
    btnCancel.Width = 80
    btnCancel.Height = 30

    Dim templates() As String
    templates = GetTemplates()
    Dim i As Integer
    For i = 0 To UBound(templates)
        lstTemplates.AddItem templates(i)
    Next i

    lstTemplates.ListIndex = 0
End Sub

Private Sub btnSend_Click()
    Dim lstTemplates As MSForms.ListBox
    Set lstTemplates = Me.Controls("lstTemplates")

    If lstTemplates.ListIndex = -1 Then
        MsgBox "Please select a template!", vbExclamation
        Exit Sub
    End If

    Dim selectedTemplate As String
    selectedTemplate = lstTemplates.Value

    Dim senderName As String
    senderName = GetSenderName(m_Mail)

    Dim replyBody As String
    replyBody = "Hi " & senderName & "," & vbCrLf & vbCrLf & selectedTemplate

    Dim objReply As MailItem
    Set objReply = m_Mail.Reply()
    objReply.Body = replyBody & vbCrLf & vbCrLf & "---" & vbCrLf & objReply.Body
    objReply.Display

    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub
