Private Sub SampleXXSendTxtEmail()
    AcquireReferences
    SendTxtEmail "vincentius.budhyanto@generali.co.id", "Test", "Test", False
End Sub

Function SendTxtEmail(ByVal Destination As String, ByVal Title As String, ByVal Content As String, ByVal DirectSend As Boolean, _
    Optional ByVal AttachmentNameAndLocation As String, Optional ByVal CopyTo As String, Optional ByVal BlindCopyTo As String)

    Set OutlookApps = CreateObject("Outlook.Application"): OutlookApps.Session.Logon
    Set NewEmail = OutlookApps.CreateItem(olMailItem)
    
    On Error GoTo CatchErr
    With NewEmail
        .To = Destination: .Subject = Title: .CC = CopyTo: .BCC = BlindCopyTo: .Body = Content
        If AttachmentNameAndLocation = "" Then
        Else: .Attachments.Add AttachmentNameAndLocation
        End If
        If DirectSend = False Then
            .Display
        ElseIf DirectSend = True Then
            .Send
        End If
    End With
    
    Debug.Print "MADE IT_" & Destination & "_" & Title & "."
    Exit Function

CatchErr:
    Debug.Print "ERROR_" & Destination & "_" & Title & "."
    Exit Function
End Function
