Option Explicit On
Option Strict On

''' <summary>
'''     Mail delivery methods
''' </summary>
Public Class SmtpWorker

    ''' <summary>
    ''' The SMTP server name for sending e-mails
    ''' </summary>
    ''' <value></value>
    ''' <remarks>The camm Web-Manager is heavily using e-mails for information and workflow purposes.</remarks>
    Public Property SmtpServerName() As String

    ''' <summary>
    ''' The SMTP server port for sending e-mails
    ''' </summary>
    ''' <value></value>
    ''' <remarks>The camm Web-Manager is heavily using e-mails for information and workflow purposes.</remarks>
    Public Property SmtpServerPort() As Integer = 25

    ''' <summary>
    '''     Sets up an user name for the SMTP server
    ''' </summary>
    ''' <value>The user name</value>
    Public Property SmtpUserName() As String

    ''' <summary>
    '''     Sets up a password for the SMTP server
    ''' </summary>
    ''' <value>The password</value>
    Public Property SmtpPassword() As String

    ''' <summary>
    '''     Configures the authentication methods for the SMTP server
    ''' </summary>
    ''' <value>The authentification type, possible values are "" or "NONE", "LOGIN", "PLAIN" or "NTLM"</value>
    Public Property SmtpAuthType() As String

    ''' <summary>
    '''     Send an e-mail via the default e-mail handler of the .NET Framework
    ''' </summary>
    ''' <param name="RcptName">Receipient's name</param>
    ''' <param name="RcptAddress">Receipient's e-mail address</param>
    ''' <param name="MsgSubject">Message subject</param>
    ''' <param name="MsgTextBody">Message plain text body</param>
    ''' <param name="MsgHTMLBody">Message html text body</param>
    ''' <param name="SenderName">The name of the sender</param>
    ''' <param name="SenderAddress">The e-mail address of the sender</param>
    ''' <param name="MsgCharset">The charset of the message, default is UTF-8</param>
    ''' <param name="Attachments">Array of e-mail attachments</param>
    ''' <param name="Priority">The priority of the e-mail</param>
    ''' <param name="Sensitivity">The sensitivity of the e-mail</param>
    ''' <param name="RequestTransmissionConfirmation">Request a transmission confirmation</param>
    ''' <param name="RequestReadingConfirmation">Request a reading confirmation</param>
    ''' <param name="AdditionalHeaders">Additional headers for the e-mail</param>
    ''' <returns>True if successfull, false for failures (also see bufErrorDetails)</returns>
    ''' <remarks>
    '''     Requires the existance of the CDO.Message object on MS platforms, otherwise this method will fail
    '''     This component only allows to send the e-mail either in plain text or html code format
    ''' </remarks>
    Public Function SendToSingeRecipient(ByVal RcptName As String, ByVal RcptAddress As String, ByVal MsgSubject As String, ByVal MsgTextBody As String, ByVal MsgHTMLBody As String, ByVal SenderName As String, ByVal SenderAddress As String, Optional ByVal MsgCharset As String = Nothing, Optional ByVal Attachments As EMailAttachment() = Nothing, Optional ByVal Priority As EMails.Priority = Nothing, Optional ByVal Sensitivity As EMails.Sensitivity = Nothing, Optional ByVal RequestTransmissionConfirmation As Boolean = False, Optional ByVal RequestReadingConfirmation As Boolean = False, Optional ByVal AdditionalHeaders As Collections.Specialized.NameValueCollection = Nothing) As SmtpSendResult
        Return SendToMultipleRecipients(EMailRecipient.CreateRecipientString(RcptName, RcptAddress), Nothing, Nothing, MsgSubject, MsgTextBody, MsgHTMLBody, SenderName, SenderAddress, MsgCharset, Attachments, CType(Priority, EMails.Priority), CType(Sensitivity, EMails.Sensitivity), RequestTransmissionConfirmation, RequestReadingConfirmation, AdditionalHeaders)
    End Function

    ''' <summary>
    '''     Send an e-mail to multiple receipents via the default e-mail handler of the .NET Framework
    ''' </summary>
    ''' <param name="RcptAddresses_To">TO receipients</param>
    ''' <param name="RcptAddresses_CC">CC receipients</param>
    ''' <param name="RcptAddresses_BCC">BCC receipients</param>
    ''' <param name="MsgSubject">Message subject</param>
    ''' <param name="MsgTextBody">Message plain text body</param>
    ''' <param name="MsgHTMLBody">Message html text body</param>
    ''' <param name="SenderName">The name of the sender</param>
    ''' <param name="SenderAddress">The e-mail address of the sender</param>
    ''' <param name="MsgCharset">The charset of the message, default is UTF-8</param>
    ''' <param name="Attachments">Array of e-mail attachments</param>
    ''' <param name="Priority">The priority of the e-mail</param>
    ''' <param name="Sensitivity">The sensitivity of the e-mail</param>
    ''' <param name="RequestTransmissionConfirmation">Request a transmission confirmation</param>
    ''' <param name="RequestReadingConfirmation">Request a reading confirmation</param>
    ''' <param name="AdditionalHeaders">Additional headers for the e-mail</param>
    ''' <returns>True if successfull, false for failures (also see bufErrorDetails)</returns>
    Public Function SendToMultipleRecipients(ByVal RcptAddresses_To As String, ByVal RcptAddresses_CC As String, ByVal RcptAddresses_BCC As String, ByVal MsgSubject As String, ByVal MsgTextBody As String, ByVal MsgHTMLBody As String, ByVal SenderName As String, ByVal SenderAddress As String, Optional ByVal MsgCharset As String = Nothing, Optional ByVal Attachments As EMailAttachment() = Nothing, Optional ByVal Priority As EMails.Priority = Nothing, Optional ByVal Sensitivity As EMails.Sensitivity = Nothing, Optional ByVal RequestTransmissionConfirmation As Boolean = False, Optional ByVal RequestReadingConfirmation As Boolean = False, Optional ByVal AdditionalHeaders As Collections.Specialized.NameValueCollection = Nothing) As SmtpSendResult
        Dim Result As SmtpSendResult
        Using MyMail As New System.Net.Mail.MailMessage()
            Dim ErrorFound As String = Nothing
            Dim smtp As New System.Net.Mail.SmtpClient()
            EMailAttachment.FixHtmlContentIDs(MsgHTMLBody, Attachments)

            'Prepare main mail properties
            smtp.Host = Me.SmtpServerName
            If Me.SmtpServerPort <> 25 Then
                smtp.Port = Me.SmtpServerPort
            End If

            Select Case UCase(Me.SmtpAuthType)
                Case "LOGIN", "PLAIN"
                    If Me.SmtpUserName = Nothing Then Throw New ArgumentNullException("SmtpUserName")
                    If Me.SmtpPassword = Nothing Then Throw New ArgumentNullException("SmtpPassword")
                    smtp.Credentials = New System.Net.NetworkCredential(Me.SmtpUserName, Me.SmtpPassword)
                    smtp.EnableSsl = False
                    smtp.Timeout = 10 * 1000 'in ms
                Case "LOGIN-SSL"
                    If Me.SmtpUserName = Nothing Then Throw New ArgumentNullException("SmtpUserName")
                    If Me.SmtpPassword = Nothing Then Throw New ArgumentNullException("SmtpPassword")
                    smtp.UseDefaultCredentials = False
                    smtp.Credentials = New System.Net.NetworkCredential(Me.SmtpUserName, Me.SmtpPassword)
                    smtp.EnableSsl = True
                    smtp.Timeout = 10 * 1000 'in ms
                Case "NTLM"
                    Return New SmtpSendResult(New NotSupportedException("NTLM is not supported"))
                Case Else
                    'SmtpAuthMethod = "NONE"
            End Select

            Try
                'Add all recepients
                If RcptAddresses_To <> Nothing Then
                    For Each rcpt As EMailRecipient In EMailRecipient.SplitEMailAddressesIntoEMailRecipientsFromRecipientsList(RcptAddresses_To)
                        MyMail.To.Add(New System.Net.Mail.MailAddress(rcpt.Address, rcpt.Name))
                    Next
                End If
                If RcptAddresses_CC <> Nothing Then
                    For Each rcpt As EMailRecipient In EMailRecipient.SplitEMailAddressesIntoEMailRecipientsFromRecipientsList(RcptAddresses_CC)
                        MyMail.CC.Add(New System.Net.Mail.MailAddress(rcpt.Address, rcpt.Name))
                    Next
                End If
                If RcptAddresses_BCC <> Nothing Then
                    For Each rcpt As EMailRecipient In EMailRecipient.SplitEMailAddressesIntoEMailRecipientsFromRecipientsList(RcptAddresses_BCC)
                        MyMail.Bcc.Add(New System.Net.Mail.MailAddress(rcpt.Address, rcpt.Name))
                    Next
                End If
            Catch ex As Exception
                'Critical error
                Return New SmtpSendResult(ex)
            End Try

            MyMail.Subject = MsgSubject

            Try
                MyMail.SubjectEncoding = System.Text.Encoding.GetEncoding(MsgCharset)
                MyMail.BodyEncoding = System.Text.Encoding.GetEncoding(MsgCharset)
            Catch
            End Try

            Dim plainView As System.Net.Mail.AlternateView = Nothing
            Dim htmlView As System.Net.Mail.AlternateView = Nothing
            If MsgTextBody <> Nothing Then
                plainView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(MsgTextBody, Nothing, "text/plain")
            End If
            If MsgHTMLBody <> Nothing Then
                htmlView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(MsgHTMLBody, Nothing, "text/html")
            End If

            MyMail.From = New System.Net.Mail.MailAddress(SenderAddress, SenderName)

            'Add attachments / related/embedded objects
            Dim TempFiles As New Collections.Generic.List(Of String)
            If Not Attachments Is Nothing Then
                For MyCounter As Integer = 0 To Attachments.Length - 1
                    Dim MyAttachment As System.Net.Mail.Attachment = Nothing
                    Dim MyEmbeddedImg As System.Net.Mail.LinkedResource = Nothing

                    If Not Attachments(MyCounter).RawData Is Nothing OrElse Attachments(MyCounter).RawDataFilename <> Nothing Then
                        Try
                            'Create a temporary file, store the data there and add that file as attachment before removing again
                            Dim tempFile As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid.ToString("n"))
                            TempFiles.Add(tempFile)
                            System.IO.Directory.CreateDirectory(tempFile)
                            tempFile &= System.IO.Path.DirectorySeparatorChar & Attachments(MyCounter).RawDataFilename
                            Dim tempFileStream As System.IO.FileStream = System.IO.File.Create(tempFile)
                            tempFileStream.Write(Attachments(MyCounter).RawData, 0, Attachments(MyCounter).RawData.Length)
                            tempFileStream.Flush()
                            tempFileStream.Close()
                            'TempFiles.Add(tempFile)
                            If Attachments(MyCounter).PlaceholderInMhtmlToBeReplacedByContentID <> Nothing Then
                                MyEmbeddedImg = New System.Net.Mail.LinkedResource(tempFile) With {
                                    .ContentId = Attachments(MyCounter).PlaceholderInMhtmlToBeReplacedByContentID
                                }
                            Else
                                MyAttachment = New System.Net.Mail.Attachment(tempFile)
                            End If
                        Catch ex As Exception
                            ErrorFound = ex.ToString & vbNewLine
                        End Try
                    ElseIf Attachments(MyCounter).FilePath <> Nothing Then
                        Dim fi As New IO.FileInfo(Attachments(MyCounter).FilePath)
                        If fi.Exists Then
                            Try
                                If Attachments(MyCounter).PlaceholderInMhtmlToBeReplacedByContentID <> Nothing Then
                                    MyEmbeddedImg = New System.Net.Mail.LinkedResource(Attachments(MyCounter).FilePath) With {
                                        .ContentId = Attachments(MyCounter).PlaceholderInMhtmlToBeReplacedByContentID
                                    }
                                Else
                                    MyAttachment = New System.Net.Mail.Attachment(Attachments(MyCounter).FilePath)
                                End If
                            Catch ex As Exception
                                Return New SmtpSendResult(ex)
                            End Try
                        Else
                            Return New SmtpSendResult(New Exception("Attachment not found: " & Attachments(MyCounter).FilePath))
                        End If
                    Else
                        Try
                            'System.Environment.StackTrace doesn't work with medium-trust --> work around it using a new exception class
                            Dim WorkaroundEx As New Exception("")
                            Dim WorkaroundStackTrace As String = WorkaroundEx.StackTrace 'contains only last few lines of stacktrace
                            Try
                                WorkaroundStackTrace = System.Environment.StackTrace 'contains full stacktrace
                            Catch
                            End Try
                        Catch
                        End Try
                    End If
                    If Not MyEmbeddedImg Is Nothing Then
                        'add the LinkedResource to the appropriate view
                        htmlView.LinkedResources.Add(MyEmbeddedImg)
                    End If
                    If Not MyAttachment Is Nothing Then
                        MyMail.Attachments.Add(MyAttachment)
                    End If
                Next
            End If

            Try
                If plainView IsNot Nothing Then
                    MyMail.AlternateViews.Add(plainView)
                End If
                If htmlView IsNot Nothing Then
                    MyMail.AlternateViews.Add(htmlView)
                End If

                'Setup mail priority
                If Not Priority = Nothing Then
                    Select Case Priority
                        Case EMails.Priority.Low
                        Case CType(4, EMails.Priority) 'Between Normal and Low
                        Case EMails.Priority.High
                        Case CType(2, EMails.Priority) 'Between Normal and High
                        Case Else 'WMSystem.MailImportance.Normal
                            Priority = EMails.Priority.Normal
                    End Select
                    MyMail.Headers.Add("X-Priority", CByte(Priority).ToString())
                    If Priority = EMails.Priority.High OrElse Priority = EMails.Priority.Normal OrElse Priority = EMails.Priority.Low Then
                        MyMail.Headers.Add("Importance", Priority.ToString())
                        MyMail.Headers.Add("Priority", Priority.ToString())
                        MyMail.Headers.Add("X-MSMail-Priority", Priority.ToString())
                    End If
                    Select Case Priority
                        Case EMails.Priority.High, CType(2, EMails.Priority)
                            MyMail.Priority = System.Net.Mail.MailPriority.High
                        Case EMails.Priority.Low, CType(4, EMails.Priority)
                            MyMail.Priority = System.Net.Mail.MailPriority.Low
                        Case Else
                            MyMail.Priority = System.Net.Mail.MailPriority.Normal
                    End Select
                End If

                'Setup mail sensitivity
                If Not Sensitivity = Nothing Then
                    Select Case Sensitivity
                        Case EMails.Sensitivity.CompanyConfidential
                            MyMail.Headers.Add("Sensitivity", "Company-Confidential")
                        Case EMails.Sensitivity.Personal
                            MyMail.Headers.Add("Sensitivity", "Personal")
                        Case EMails.Sensitivity.Private
                            MyMail.Headers.Add("Sensitivity", "Private")
                    End Select
                End If

                'Setup mail receipt confirmations
                If RequestReadingConfirmation Then
                    'Disposition-Notification-To: "Jochen Wezel" <my-company@my-webmailer-company.com> 'Lesebest?tigung
                    MyMail.Headers.Add("Disposition-Notification-To", SenderAddress)
                End If
                'MIGHT BE FIXED OR IS STILL TODO: The bug is that nothing happens :( _
                If RequestTransmissionConfirmation Then
                    'Return-Receipt-To: "Jochen Wezel" <my-company@my-webmailer-company.com> '?bermittlungsbest?tigung
                    'MyMail.Headers.Add("Return-Receipt-To", SenderAddress)
                    MyMail.DeliveryNotificationOptions = System.Net.Mail.DeliveryNotificationOptions.OnSuccess
                End If

                'Setup additional headers
                If Not AdditionalHeaders Is Nothing Then
                    For Each MyKey As String In AdditionalHeaders
                        If MyKey.ToLower() = "reply-to" AndAlso AdditionalHeaders(MyKey) <> Nothing AndAlso AdditionalHeaders(MyKey) <> "<>" Then 'for some time, there was an issue that additional header reply-to might contain non-empty value but meant as empty (value="<>")
                            Dim rcpt As EMailRecipient = EMailRecipient.SplitEMailAddressesIntoAddressParts(AdditionalHeaders(MyKey))
                            MyMail.ReplyTo = New System.Net.Mail.MailAddress(rcpt.Address, rcpt.Name)
                        ElseIf AdditionalHeaders(MyKey) <> Nothing AndAlso AdditionalHeaders(MyKey) <> "<>" Then 'don't add invalid header value
                            MyMail.Headers.Add(MyKey, AdditionalHeaders(MyKey))
                        End If
                    Next
                End If

                'Send the mail and return the success/error status
                If ErrorFound = Nothing Then
                    smtp.Send(MyMail)
                    Result = New SmtpSendResult()
                Else
                    Result = New SmtpSendResult(New Exception(ErrorFound))
                End If
            Catch ex As Exception
                Result = New SmtpSendResult(ex)
            End Try

            'Clean up all temporary files
            For MyCounter As Integer = 0 To TempFiles.Count - 1
                Try
                    System.IO.Directory.Delete(CType(TempFiles(MyCounter), String), True)
                Catch
                End Try
            Next
        End Using

        Return Result

    End Function

End Class
