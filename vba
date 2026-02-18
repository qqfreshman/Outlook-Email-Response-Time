' Check_Email_Response.vb
' Global variable declarations at the top of your module
Public teamSenders As Variant
Public filteredMails As Object ' Global variable for filtered mails

' Initialize the global variables - FIXED VERSION
Sub InitializeTeamSenders()
    ' Split the array initialization into smaller chunks to avoid "too many continuations" error
    Dim tempArray(32) As String
    
    tempArray(0) = "Kobe, Bryant"
    tempArray(1) = "Kevin, Durant"
    
    teamSenders = tempArray
End Sub

' Diagnostic subroutine to check folder and date ranges
Sub DiagnoseEmailFolder()
    Dim olFolder As Outlook.MAPIFolder
    Dim cutoffDate As Date
    
    ' Prompt user to select the mailbox/folder
    Set olFolder = Application.Session.PickFolder
    If olFolder Is Nothing Then 
        MsgBox "No folder selected. Operation cancelled.", vbExclamation, "Folder Selection"
        Exit Sub
    End If
    
    cutoffDate = Date - 14
    
    Debug.Print "=== EMAIL FOLDER DIAGNOSIS ==="
    Debug.Print "Current date: " & Format(Date, "yyyy-mm-dd")
    Debug.Print "Cutoff date (2 weeks ago): " & Format(cutoffDate, "yyyy-mm-dd")
    
    Call DiagnoseFolderContents(olFolder, cutoffDate)
    
    MsgBox "Diagnosis complete. Check Immediate Window for details.", vbInformation, "Diagnosis Complete"
    
    Set olFolder = Nothing
End Sub

' Optimized procedure to initialize filtered mails for last 2 weeks - FIXED VERSION
Sub InitializeFilteredMails(olFolder As Outlook.MAPIFolder)
    Dim olMail As Object
    Dim cutoffDate As Date
    Dim mailIndex As Long
    Dim itemCount As Long
    Dim processedCount As Long
    Dim futureEmailCount As Long
    
    On Error GoTo ErrorHandler
    
    ' FIXED: Always clear and recreate the filteredMails dictionary
    Call ClearFilteredMails

    ' Set cutoff date to 2 weeks ago
    cutoffDate = Date - 14
    
    ' Create collection object to store filtered mails
    Set filteredMails = CreateObject("Scripting.Dictionary")
    mailIndex = 0
    itemCount = olFolder.Items.Count
    processedCount = 0
    futureEmailCount = 0
    
    ' Sort items by ReceivedTime in descending order (newest first)
    olFolder.Items.Sort "[ReceivedTime]", True
    
    Debug.Print "Processing " & itemCount & " items in folder..."
    Debug.Print "Cutoff date: " & Format(cutoffDate, "yyyy-mm-dd hh:nn:ss")
    Debug.Print "Today's date: " & Format(Date, "yyyy-mm-dd")
    
    For Each olMail In olFolder.Items
        If TypeOf olMail Is MailItem Then
            processedCount = processedCount + 1
            
            ' Debug: Show first few email dates to understand the data
            If processedCount <= 5 Then
                Debug.Print "Email " & processedCount & " - Sender: " & olMail.SenderName & ", Time: " & Format(olMail.ReceivedTime, "yyyy-mm-dd hh:nn:ss")
            End If
            
            ' FIXED: Check for future emails (likely system clock issues)
            If olMail.ReceivedTime > Date + 1 Then
                futureEmailCount = futureEmailCount + 1
                If futureEmailCount <= 3 Then ' Only log first few
                    Debug.Print "WARNING: Future email detected - " & Format(olMail.ReceivedTime, "yyyy-mm-dd hh:nn:ss")
                End If
                ' Skip future emails but continue processing
                GoTo NextEmail
            End If
            
            ' FIXED: Include emails within our date range (don't exit early)
            If olMail.ReceivedTime >= cutoffDate Then
                filteredMails.Add mailIndex, olMail
                mailIndex = mailIndex + 1
            End If
            
            ' FIXED: Only exit if we've processed a reasonable number of old emails
            ' This prevents early exit due to mixed timestamps
            If olMail.ReceivedTime < cutoffDate Then
                ' Count consecutive old emails
                Static consecutiveOldEmails As Long
                consecutiveOldEmails = consecutiveOldEmails + 1
                
                ' Exit only after finding 50 consecutive old emails
                ' This handles mixed timestamps better
                If consecutiveOldEmails >= 50 Then
                    Debug.Print "Found 50 consecutive emails older than cutoff. Stopping collection."
                    Debug.Print "Last processed - Sender: " & olMail.SenderName & ", Time: " & Format(olMail.ReceivedTime, "yyyy-mm-dd hh:nn:ss")
                    Exit For
                End If
            Else
                ' Reset counter if we find a recent email
                consecutiveOldEmails = 0
            End If
        End If
        
NextEmail:
    Next
    
    Debug.Print "Processed " & processedCount & " total items"
    Debug.Print "Found " & futureEmailCount & " emails with future timestamps"
    Debug.Print "Filtered " & filteredMails.Count & " mails from last 2 weeks"
    
    ' Additional validation
    If filteredMails.Count = 0 Then
        Debug.Print "No emails found in date range. Checking folder contents..."
        Call DiagnoseFolderContents(olFolder, cutoffDate)
    End If
    
    GoTo CleanupExit
    
ErrorHandler:
    Debug.Print "Error in InitializeFilteredMails: " & Err.Description
    Resume CleanupExit
    
CleanupExit:
    Set olMail = Nothing
    On Error GoTo 0
End Sub

' New diagnostic subroutine to help troubleshoot folder contents
Private Sub DiagnoseFolderContents(olFolder As Outlook.MAPIFolder, cutoffDate As Date)
    Dim olMail As Object
    Dim sampleCount As Long
    Dim totalEmails As Long
    Dim emailsInRange As Long
    Dim oldestEmail As Date
    Dim newestEmail As Date
    
    On Error Resume Next
    
    Debug.Print "=== FOLDER DIAGNOSIS ==="
    Debug.Print "Folder name: " & olFolder.Name
    Debug.Print "Total items: " & olFolder.Items.Count
    Debug.Print "Looking for emails newer than: " & Format(cutoffDate, "yyyy-mm-dd hh:nn:ss")
    
    oldestEmail = Date + 365 ' Initialize to future date
    newestEmail = Date - 365  ' Initialize to past date
    sampleCount = 0
    totalEmails = 0
    emailsInRange = 0
    
    ' Sample first 20 emails to understand date distribution
    For Each olMail In olFolder.Items
        If TypeOf olMail Is MailItem Then
            totalEmails = totalEmails + 1
            sampleCount = sampleCount + 1
            
            If olMail.ReceivedTime < oldestEmail Then oldestEmail = olMail.ReceivedTime
            If olMail.ReceivedTime > newestEmail Then newestEmail = olMail.ReceivedTime
            
            If olMail.ReceivedTime >= cutoffDate Then emailsInRange = emailsInRange + 1
            
            If sampleCount <= 10 Then
                Debug.Print "Sample " & sampleCount & ": " & Format(olMail.ReceivedTime, "yyyy-mm-dd hh:nn:ss") & " - " & Left(olMail.Subject, 30)
            End If
            Debug.Print "sender " & olMail.SenderName & " subject " & olMail.Subject & " received " & Format(olMail.ReceivedTime, "yyyy-mm-dd hh:nn:ss")
            If sampleCount >= 20 Then Exit For
        End If
    Next
    
    Debug.Print "Total emails found: " & totalEmails
    Debug.Print "Emails in date range: " & emailsInRange
    Debug.Print "Oldest email: " & Format(oldestEmail, "yyyy-mm-dd hh:nn:ss")
    Debug.Print "Newest email: " & Format(newestEmail, "yyyy-mm-dd hh:nn:ss")
    Debug.Print "========================"
    
    Set olMail = Nothing
    On Error GoTo 0
End Sub

' Optimized main procedure with improved error handling
Sub Mail_By_Response()
    Dim olFolder As Outlook.MAPIFolder
    Dim olMail As Object
    Dim dictDetails As Object
    Dim OutMail As Outlook.MailItem
    Dim cutoffDate As Date
    Dim errorCount As Long
    Dim processedConversations As Object
    Dim mailKey As Variant
    Dim msg As String
    Dim startTime As Date
    
    startTime = Now
    On Error GoTo ErrorHandler

    ' FIXED: Always clear previous data first
    Call ClearFilteredMails
    ' Initialize global team senders if not already done
    If IsEmpty(teamSenders) Then
        Call InitializeTeamSenders
    End If
    
    Set dictDetails = CreateObject("Scripting.Dictionary")
    Set processedConversations = CreateObject("Scripting.Dictionary")
    cutoffDate = Date - 7 ' Last 7 days for processing (but we have 2 weeks of data)
    errorCount = 0

    ' Prompt user to select the mailbox/folder
    Set olFolder = Application.Session.PickFolder
    If olFolder Is Nothing Then 
        MsgBox "No folder selected. Operation cancelled.", vbExclamation, "Folder Selection"
        GoTo CleanupAndExit
    End If
    
    ' Initialize filtered mails for last 2 weeks
    Call InitializeFilteredMails(olFolder)
    
    ' Check if any mails were found
    If filteredMails Is Nothing Or filteredMails.Count = 0 Then
        MsgBox "No emails found in the last 2 weeks.", vbInformation, "No Data"
        GoTo CleanupAndExit
    End If
    
    Debug.Print "Starting analysis of " & filteredMails.Count & " emails..."
    
    ' Process filtered mails
    For Each mailKey In filteredMails.Keys
        On Error Resume Next
        Set olMail = filteredMails(mailKey)
        
        ' Check if there was an error accessing the item
        If Err.Number <> 0 Then
            errorCount = errorCount + 1
            Debug.Print "Error accessing mail item: " & Err.Description
            Err.Clear
            GoTo NextItem
        End If
        
        On Error GoTo ErrorHandler
        
        ' Only process emails within our analysis window
        If olMail.ReceivedTime >= cutoffDate Then
            Call ProcessEmailConversation(olMail, dictDetails, processedConversations, cutoffDate)
        End If
        
NextItem:
        Set olMail = Nothing
    Next

    Debug.Print "Analysis completed in " & DateDiff("s", startTime, Now) & " seconds"
    
    ' Build and send summary
    msg = BuildSummaryMessageWithTable(dictDetails, errorCount)
    Call SendSummaryEmail(msg)
    
    MsgBox "Summary sent!" & vbCrLf & "Processed " & filteredMails.Count & " mails from last 2 weeks" & vbCrLf & _
           "Analysis time: " & DateDiff("s", startTime, Now) & " seconds", vbInformation, "Analysis Complete"
    
    GoTo CleanupAndExit

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description & vbCrLf & "Error Number: " & Err.Number, vbCritical, "Error"
    Debug.Print "Error in Mail_By_Response: " & Err.Description

CleanupAndExit:
    ' Complete cleanup
    Set dictDetails = Nothing
    Set processedConversations = Nothing
    Set OutMail = Nothing
    Set olFolder = Nothing
    Set olMail = Nothing
    Call ClearFilteredMails
    On Error GoTo 0
End Sub

' New helper subroutine to process individual email conversations
Private Sub ProcessEmailConversation(olMail As Object, dictDetails As Object, processedConversations As Object, cutoffDate As Date)
    Dim convID As String
    Dim mail_subject As String
    Dim firstTeamReply As Date
    Dim firstUserEmail As Date
    Dim firstTeamSender As String
    Dim hasTeamReply As Boolean
    Dim responseMinutes As Long
    
    convID = olMail.ConversationID
    mail_subject = olMail.Subject
    
    ' Skip if conversation already processed or no conversation ID
    If convID = "" Or processedConversations.Exists(convID) Then
        Exit Sub
    End If
    
    ' Mark this conversation as processed
    processedConversations.Add convID, True
    
    Debug.Print "Processing: " & Left(mail_subject, 50) & "..."
    
    ' Get first team reply info
    hasTeamReply = GetTeam1stReplyInfo(convID, cutoffDate, firstTeamReply, firstUserEmail, firstTeamSender)
    
    If hasTeamReply And firstUserEmail > 0 Then
        responseMinutes = DateDiff("n", firstUserEmail, firstTeamReply)
    Else
        responseMinutes = -1 ' No team reply found or no user email found
    End If

    ' Store the conversation details
    Call StoreConversationDetails(dictDetails, convID, mail_subject, firstUserEmail, firstTeamReply, responseMinutes, firstTeamSender)
End Sub

' Helper subroutine to store conversation details
Private Sub StoreConversationDetails(dictDetails As Object, convID As String, mail_subject As String, _
                                   firstUserEmail As Date, firstTeamReply As Date, responseMinutes As Long, firstTeamSender As String)
    Dim replyTimeStr As String
    Dim userEmailTimeStr As String
    Dim safeSubject As String
    
    ' Format time strings
    If firstTeamReply > 0 Then
        replyTimeStr = Format(firstTeamReply, "yyyy-mm-dd hh:nn")
    Else
        replyTimeStr = "No Reply"
    End If
    
    If firstUserEmail > 0 Then
        userEmailTimeStr = Format(firstUserEmail, "yyyy-mm-dd hh:nn")
    Else
        userEmailTimeStr = "N/A"
    End If
    
    ' Clean subject for HTML display
    safeSubject = CleanSubjectForHTML(mail_subject)
    
    ' Store details: subject|userTime|replyTime|responseMinutes|teamSender
    dictDetails.Add convID, safeSubject & "|" & userEmailTimeStr & "|" & replyTimeStr & "|" & responseMinutes & "|" & firstTeamSender
End Sub

' Helper function to clean subject for HTML display
Private Function CleanSubjectForHTML(subject As String) As String
    Dim cleanSubject As String
    cleanSubject = subject
    cleanSubject = Replace(cleanSubject, "&", "&amp;")
    cleanSubject = Replace(cleanSubject, "<", "&lt;")
    cleanSubject = Replace(cleanSubject, ">", "&gt;")
    cleanSubject = Replace(cleanSubject, """", "&quot;")
    cleanSubject = Replace(cleanSubject, "'", "&#39;")
    cleanSubject = Replace(cleanSubject, vbCrLf, " ")
    cleanSubject = Replace(cleanSubject, vbCr, " ")
    cleanSubject = Replace(cleanSubject, vbLf, " ")
    CleanSubjectForHTML = cleanSubject
End Function

' Separate subroutine for sending email
Private Sub SendSummaryEmail(msg As String)
    Dim OutMail As Outlook.MailItem
    
    On Error GoTo EmailError
    
    Set OutMail = Application.CreateItem(olMailItem)
    With OutMail
        .To = "admin@abc.com"
        .Subject = "Email Response Time Analysis - " & Format(Date, "yyyy-mm-dd")
        .HTMLBody = msg
        .Send
    End With
    
    GoTo EmailCleanup

EmailError:
    MsgBox "Error sending email: " & Err.Description, vbCritical, "Email Error"
    
EmailCleanup:
    Set OutMail = Nothing
    On Error GoTo 0
End Sub

' Optimized function to get team first reply info
Function GetTeam1stReplyInfo(conversationID As String, cutoffDate As Date, ByRef firstTeamReply As Date, _
                           ByRef firstUserEmail As Date, ByRef firstTeamSender As String) As Boolean
    Dim olItem As Object
    Dim tempFirstReply As Date
    Dim tempFirstUser As Date
    Dim temp1stTeamSender As String
    Dim mailKey As Variant
    Dim conversationEmails As Object
    
    On Error GoTo ErrorHandler
    
    ' Initialize return values
    tempFirstReply = #12/31/9999#
    tempFirstUser = #12/31/9999#
    temp1stTeamSender = ""
    
    ' Create a collection of emails in this conversation for better processing
    Set conversationEmails = CreateObject("Scripting.Dictionary")
    
    ' First pass: collect all emails in this conversation
    For Each mailKey In filteredMails.Keys
        Set olItem = filteredMails(mailKey)
        If TypeOf olItem Is MailItem Then
            If olItem.ConversationID = conversationID And olItem.ReceivedTime >= cutoffDate Then
                conversationEmails.Add olItem.ReceivedTime, olItem
            End If
        End If
        Set olItem = Nothing
    Next

    ' Second pass: find first team reply
    For Each mailKey In conversationEmails.Keys
        Set olItem = conversationEmails(mailKey)
        If IsTeamSender(olItem.SenderName) Then
            If olItem.ReceivedTime < tempFirstReply Then
                tempFirstReply = olItem.ReceivedTime
                temp1stTeamSender = olItem.categories
            End If
        End If
        Set olItem = Nothing
    Next

    ' Third pass: if team reply found, find the latest user email before it
    If tempFirstReply < #12/31/9999# Then
        For Each mailKey In conversationEmails.Keys
            Set olItem = conversationEmails(mailKey)
            If olItem.ReceivedTime < tempFirstReply And Not IsTeamSender(olItem.SenderName) Then
                If olItem.ReceivedTime > tempFirstUser Or tempFirstUser = #12/31/9999# Then
                    tempFirstUser = olItem.ReceivedTime
                End If
            End If
            Set olItem = Nothing
        Next
        
        firstTeamReply = tempFirstReply
        firstUserEmail = tempFirstUser
        firstTeamSender = temp1stTeamSender
        GetTeam1stReplyInfo = True
    Else
        firstTeamReply = 0
        firstUserEmail = 0
        firstTeamSender = ""
        GetTeam1stReplyInfo = False
    End If
    
    GoTo CleanupExit

ErrorHandler:
    Debug.Print "Error in GetTeam1stReplyInfo: " & Err.Description
    GetTeam1stReplyInfo = False
    
CleanupExit:
    Set conversationEmails = Nothing
    Set olItem = Nothing
    On Error GoTo 0
End Function

' Optimized team sender check function
Private Function IsTeamSender(senderName As String) As Boolean
    Dim sender As Variant
    Dim senderLower As String
    
    IsTeamSender = False
    
    ' Input validation
    If senderName = "" Or IsNull(senderName) Then Exit Function
    
    senderLower = LCase(Trim(senderName))
    If senderLower = "" Then Exit Function
    
    ' Check against team senders list
    For Each sender In teamSenders
        If Not IsNull(sender) And Trim(CStr(sender)) <> "" Then
            If InStr(senderLower, LCase(Trim(CStr(sender)))) > 0 Then
                IsTeamSender = True
                Exit Function
            End If
        End If
    Next sender
End Function

' Helper function to build summary message with HTML table (unchanged but with better error handling)
Private Function BuildSummaryMessageWithTable(dictDetails As Object, errorCount As Long) As String
    Dim msg As String
    Dim i As Long
    Dim detailKeys As Variant
    Dim emailDetails As Variant
    Dim responseTimeStr As String
    Dim responseClass As String
    
    On Error GoTo ErrorHandler
    
    ' Start HTML document
    msg = BuildHTMLHeader()
    
    ' Add summary information
    msg = msg & BuildSummarySection(errorCount)
    
    If dictDetails.Count > 0 Then
        ' Add statistics
        msg = msg & BuildResponseTimeStats(dictDetails)
        
        ' Add detailed table
        msg = msg & BuildDetailedTable(dictDetails)
    Else
        msg = msg & BuildNoDataSection()
    End If
    
    ' Add footer and close HTML
    msg = msg & BuildHTMLFooter()
    
    BuildSummaryMessageWithTable = msg
    GoTo ExitFunction

ErrorHandler:
    Debug.Print "Error in BuildSummaryMessageWithTable: " & Err.Description
    BuildSummaryMessageWithTable = "<html><body><h1>Error generating report</h1><p>" & Err.Description & "</p></body></html>"
    
ExitFunction:
    On Error GoTo 0
End Function

' Helper functions for building HTML sections
Private Function BuildHTMLHeader() As String
    Dim html As String
    html = "<!DOCTYPE html>" & vbCrLf
    html = html & "<html>" & vbCrLf
    html = html & "<head>" & vbCrLf
    html = html & "<style>" & vbCrLf
    html = html & "body { font-family: Arial, sans-serif; margin: 20px; }" & vbCrLf
    html = html & "h1 { color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px; }" & vbCrLf
    html = html & "h2 { color: #34495e; margin-top: 30px; }" & vbCrLf
    html = html & "table { border-collapse: collapse; width: 100%; margin-top: 20px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }" & vbCrLf
    html = html & "th { background-color: #3498db; color: white; padding: 12px; text-align: left; font-weight: bold; }" & vbCrLf
    html = html & "td { padding: 10px; border-bottom: 1px solid #ecf0f1; }" & vbCrLf
    html = html & "tr:nth-child(even) { background-color: #f8f9fa; }" & vbCrLf
    html = html & "tr:hover { background-color: #e8f4f8; }" & vbCrLf
    html = html & ".no-reply { background-color: #ffebee; color: #c62828; font-weight: bold; }" & vbCrLf
    html = html & ".fast-reply { background-color: #e8f5e8; color: #2e7d32; }" & vbCrLf
    html = html & ".slow-reply { background-color: #fff3e0; color: #ef6c00; }" & vbCrLf
    html = html & ".very-slow-reply { background-color: #ffebee; color: #c62828; }" & vbCrLf
    html = html & ".summary-box { background-color: #f8f9fa; padding: 15px; border-left: 4px solid #3498db; margin: 20px 0; }" & vbCrLf
    html = html & "</style>" & vbCrLf
    html = html & "</head>" & vbCrLf
    html = html & "<body>" & vbCrLf
    BuildHTMLHeader = html
End Function

Private Function BuildSummarySection(errorCount As Long) As String
    Dim html As String
    html = "<h1>Email Response Time Analysis</h1>" & vbCrLf
    html = html & "<p><strong>Analysis Period:</strong> Last 7 days (from 2 weeks of data)</p>" & vbCrLf
    html = html & "<p><strong>Generated at:</strong> " & Format(Now, "yyyy-mm-dd hh:nn:ss") & "</p>" & vbCrLf
    html = html & "<div class='summary-box'>" & vbCrLf
    html = html & "<h2>Summary</h2>" & vbCrLf
    html = html & "<p><strong>Total mails in last 2 weeks:</strong> " & filteredMails.Count & "</p>" & vbCrLf
    
    If errorCount > 0 Then
        html = html & "<p><strong>Note:</strong> " & errorCount & " item(s) skipped due to access errors</p>" & vbCrLf
    End If
    html = html & "</div>" & vbCrLf
    BuildSummarySection = html
End Function

Private Function BuildNoDataSection() As String
    Dim html As String
    html = "<div class='summary-box'>" & vbCrLf
    html = html & "<h2>No Data Found</h2>" & vbCrLf
    html = html & "<p>No emails found matching the criteria for the specified date range.</p>" & vbCrLf
    html = html & "</div>" & vbCrLf
    BuildNoDataSection = html
End Function

Private Function BuildDetailedTable(dictDetails As Object) As String
    Dim html As String
    Dim i As Long
    Dim detailKeys As Variant
    Dim emailDetails As Variant
    Dim responseTimeStr As String
    Dim responseClass As String
    
    html = "<h2>Detailed Email Analysis</h2>" & vbCrLf
    html = html & "<table>" & vbCrLf
    html = html & "<thead>" & vbCrLf
    html = html & "<tr>" & vbCrLf
    html = html & "<th>Subject</th>" & vbCrLf
    html = html & "<th>User Email Time</th>" & vbCrLf
    html = html & "<th>1st Team Reply</th>" & vbCrLf
    html = html & "<th>Response Time</th>" & vbCrLf
    html = html & "<th>Team Responder</th>" & vbCrLf
    html = html & "</tr>" & vbCrLf
    html = html & "</thead>" & vbCrLf
    html = html & "<tbody>" & vbCrLf
    
    detailKeys = dictDetails.Keys
    For i = 0 To dictDetails.Count - 1
        emailDetails = Split(dictDetails(detailKeys(i)), "|")
        
        If UBound(emailDetails) >= 4 Then
            ' Determine response time and styling
            If IsNumeric(emailDetails(3)) And CLng(emailDetails(3)) = -1 Then
                responseTimeStr = "No Reply"
                responseClass = "no-reply"
            ElseIf IsNumeric(emailDetails(3)) Then
                Dim minutes As Long
                minutes = CLng(emailDetails(3))
                responseTimeStr = minutes & " minutes"
                
                ' Color code based on response time
                If minutes <= 60 Then
                    responseClass = "fast-reply"
                ElseIf minutes <= 240 Then
                    responseClass = "slow-reply"
                Else
                    responseClass = "very-slow-reply"
                End If
            Else
                responseTimeStr = CStr(emailDetails(3))
                responseClass = ""
            End If
            
            ' Add table row
            html = html & "<tr>" & vbCrLf
            html = html & "<td>" & Left(emailDetails(0), 80) & IIf(Len(emailDetails(0)) > 80, "...", "") & "</td>" & vbCrLf
            html = html & "<td>" & emailDetails(1) & "</td>" & vbCrLf
            html = html & "<td>" & emailDetails(2) & "</td>" & vbCrLf
            html = html & "<td class='" & responseClass & "'>" & responseTimeStr & "</td>" & vbCrLf
            html = html & "<td>" & emailDetails(4) & "</td>" & vbCrLf
            html = html & "</tr>" & vbCrLf
        End If
    Next
    
    html = html & "</tbody>" & vbCrLf
    html = html & "</table>" & vbCrLf
    BuildDetailedTable = html
End Function

Private Function BuildHTMLFooter() As String
    Dim html As String
    html = "<hr style='margin-top: 40px; border: none; border-top: 1px solid #bdc3c7;'>" & vbCrLf
    html = html & "<p style='color: #7f8c8d; font-size: 0.9em;'>Generated by Email Response Time Analyzer | " & Format(Now, "yyyy-mm-dd hh:nn:ss") & "</p>" & vbCrLf
    html = html & "</body>" & vbCrLf
    html = html & "</html>" & vbCrLf
    BuildHTMLFooter = html
End Function

' Helper function to build response time statistics (optimized)
Private Function BuildResponseTimeStats(dictDetails As Object) As String
    Dim stats As String
    Dim i As Long
    Dim detailKeys As Variant
    Dim emailDetails As Variant
    Dim totalReplies As Long
    Dim noReplies As Long
    Dim fastReplies As Long
    Dim slowReplies As Long
    Dim verySlowReplies As Long
    Dim totalMinutes As Long
    Dim avgMinutes As Double
    
    ' Initialize counters
    totalReplies = 0
    noReplies = 0
    fastReplies = 0
    slowReplies = 0
    verySlowReplies = 0
    totalMinutes = 0
    
    detailKeys = dictDetails.Keys
    
    For i = 0 To dictDetails.Count - 1
        emailDetails = Split(dictDetails(detailKeys(i)), "|")
        
        If UBound(emailDetails) >= 4 Then
            If IsNumeric(emailDetails(3)) Then
                Dim minutes As Long
                minutes = CLng(emailDetails(3))
                
                If minutes = -1 Then
                    noReplies = noReplies + 1
                Else
                    totalReplies = totalReplies + 1
                    totalMinutes = totalMinutes + minutes
                    
                    If minutes <= 60 Then
                        fastReplies = fastReplies + 1
                    ElseIf minutes <= 240 Then
                        slowReplies = slowReplies + 1
                    Else
                        verySlowReplies = verySlowReplies + 1
                    End If
                End If
            End If
        End If
    Next
    
    If totalReplies > 0 Then
        avgMinutes = totalMinutes / totalReplies
    End If
    
    stats = "<div class='summary-box'>" & vbCrLf
    stats = stats & "<h2>Response Time Statistics</h2>" & vbCrLf
    stats = stats & "<table style='width: auto; margin: 0;'>" & vbCrLf
    stats = stats & "<tr><td><strong>Total Conversations Analyzed:</strong></td><td>" & dictDetails.Count & "</td></tr>" & vbCrLf
    stats = stats & "<tr><td><strong>Conversations with Team Replies:</strong></td><td>" & totalReplies & "</td></tr>" & vbCrLf
    stats = stats & "<tr><td><strong>No Reply:</strong></td><td class='no-reply'>" & noReplies & "</td></tr>" & vbCrLf
    stats = stats & "<tr><td><strong>Fast Reply (≤1 hour):</strong></td><td class='fast-reply'>" & fastReplies & "</td></tr>" & vbCrLf
    stats = stats & "<tr><td><strong>Slow Reply (1-4 hours):</strong></td><td class='slow-reply'>" & slowReplies & "</td></tr>" & vbCrLf
    stats = stats & "<tr><td><strong>Very Slow Reply (>4 hours):</strong></td><td class='very-slow-reply'>" & verySlowReplies & "</td></tr>" & vbCrLf
    
    If totalReplies > 0 Then
        stats = stats & "<tr><td><strong>Average Response Time:</strong></td><td><strong>" & Format(avgMinutes, "0.0") & " minutes (" & Format(avgMinutes / 60, "0.1") & " hours)</strong></td></tr>" & vbCrLf
    End If
    
    stats = stats & "</table>" & vbCrLf
    stats = stats & "</div>" & vbCrLf
    
    BuildResponseTimeStats = stats
End Function

' Utility subroutines
Sub UpdateTeamSenders()
    Call InitializeTeamSenders
    MsgBox "Team senders list updated successfully!" & vbCrLf & "Total team members: " & UBound(teamSenders) + 1, vbInformation, "Configuration Updated"
End Sub

' FIXED: Enhanced cleanup procedure
Sub ClearFilteredMails()
    On Error Resume Next
    If Not filteredMails Is Nothing Then
        Debug.Print "Clearing " & filteredMails.Count & " filtered mails"
        filteredMails.RemoveAll
        Set filteredMails = Nothing
    End If
    On Error GoTo 0
End Sub


' Test subroutine for debugging
Sub TestTeamSenders()
    Call InitializeTeamSenders
    Dim i As Integer
    Debug.Print "Team Senders List:"
    For i = 0 To UBound(teamSenders)
        Debug.Print i & ": " & teamSenders(i)
    Next
    MsgBox "Check Immediate Window for team senders list", vbInformation
End Sub
