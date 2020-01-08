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
    Public Property SmtpAuthType() As SmtpAuthTypes

    Public Enum SmtpAuthTypes
        None = 0
        LoginPlainText = 1
        LoginSsl = 2
    End Enum

    ''' <summary>
    '''     Send an e-mail to multiple receipents via the default e-mail handler of the .NET Framework
    ''' </summary>
    ''' <param name="message">An e-mail message object</param>
    ''' <returns>A result of SmtpSendResult containing success status and if failed the occured exception</returns>
    Public Function SendMessage(message As EMailMessage) As SmtpSendResult
        If message Is Nothing Then Throw New ArgumentNullException(NameOf(message))
        If message.To.Count = 0 AndAlso message.Cc.Count = 0 AndAlso message.Bcc.Count = 0 Then Throw New ArgumentException("At least one recipient must be provided")
        If message.From Is Nothing Then Throw New ArgumentNullException(NameOf(message.From))
        If Me.SmtpServerName = Nothing Then Throw New ArgumentNullException(NameOf(Me.SmtpServerName))

        Dim Result As SmtpSendResult
        Using MyMail As New System.Net.Mail.MailMessage()
            Dim ErrorFound As Exception = Nothing
            Dim Smtp As New System.Net.Mail.SmtpClient()
            message.PrepareForSending()

            'Prepare main mail properties
            Smtp.Host = Me.SmtpServerName
            If Me.SmtpServerPort <> 25 Then
                smtp.Port = Me.SmtpServerPort
            End If

            Select Case Me.SmtpAuthType
                Case SmtpAuthTypes.LoginPlainText
                    If Me.SmtpUserName = Nothing Then Throw New ArgumentNullException("SmtpUserName")
                    If Me.SmtpPassword = Nothing Then Throw New ArgumentNullException("SmtpPassword")
                    Smtp.Credentials = New System.Net.NetworkCredential(Me.SmtpUserName, Me.SmtpPassword)
                    Smtp.EnableSsl = False
                    Smtp.Timeout = 10 * 1000 'in ms
                Case SmtpAuthTypes.LoginSsl
                    If Me.SmtpUserName = Nothing Then Throw New ArgumentNullException("SmtpUserName")
                    If Me.SmtpPassword = Nothing Then Throw New ArgumentNullException("SmtpPassword")
                    Smtp.UseDefaultCredentials = False
                    Smtp.Credentials = New System.Net.NetworkCredential(Me.SmtpUserName, Me.SmtpPassword)
                    Smtp.EnableSsl = True
                    Smtp.Timeout = 10 * 1000 'in ms
                Case Else
                    'SmtpAuthMethod = "NONE"
            End Select

            Try
                'Add all recepients
                For Each rcpt As EMailRecipient In message.To
                    MyMail.To.Add(New System.Net.Mail.MailAddress(rcpt.Address, rcpt.Name))
                Next
                For Each rcpt As EMailRecipient In message.Cc
                    MyMail.CC.Add(New System.Net.Mail.MailAddress(rcpt.Address, rcpt.Name))
                Next
                For Each rcpt As EMailRecipient In message.Bcc
                    MyMail.Bcc.Add(New System.Net.Mail.MailAddress(rcpt.Address, rcpt.Name))
                Next
            Catch ex As Exception
                'Critical error
                Return New SmtpSendResult(ex)
            End Try

            MyMail.Subject = message.Subject

            Try
                MyMail.SubjectEncoding = message.MessageEncoding
                MyMail.BodyEncoding = message.MessageEncoding
            Catch
            End Try

            Dim plainView As System.Net.Mail.AlternateView = Nothing
            Dim htmlView As System.Net.Mail.AlternateView = Nothing
            If message.BodyPlainText <> Nothing Then
                plainView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(message.BodyPlainText, CType(Nothing, Text.Encoding), "text/plain")
            End If
            If message.BodyHtml <> Nothing Then
                htmlView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(message.BodyHtml, CType(Nothing, Text.Encoding), "text/html")
            End If

            MyMail.From = New System.Net.Mail.MailAddress(message.From.Address, message.From.Name)

            'Add attachments / related/embedded objects
            Dim TempFiles As New Collections.Generic.List(Of String)
            For MyCounter As Integer = 0 To message.EMailAttachments.Count - 1
                Dim MyAttachment As System.Net.Mail.Attachment = Nothing
                Dim MyEmbeddedImg As System.Net.Mail.LinkedResource = Nothing

                If Not message.EMailAttachments(MyCounter).RawData Is Nothing OrElse message.EMailAttachments(MyCounter).RawDataFilename <> Nothing Then
                    Try
                        'Create a temporary file, store the data there and add that file as attachment before removing again
                        Dim tempFile As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid.ToString("n"))
                        TempFiles.Add(tempFile)
                        System.IO.Directory.CreateDirectory(tempFile)
                        tempFile &= System.IO.Path.DirectorySeparatorChar & message.EMailAttachments(MyCounter).RawDataFilename
                        Dim tempFileStream As System.IO.FileStream = System.IO.File.Create(tempFile)
                        tempFileStream.Write(message.EMailAttachments(MyCounter).RawData, 0, message.EMailAttachments(MyCounter).RawData.Length)
                        tempFileStream.Flush()
                        tempFileStream.Close()
                        'TempFiles.Add(tempFile)
                        If message.EMailAttachments(MyCounter).PlaceholderInMhtmlToBeReplacedByContentID <> Nothing Then
                            MyEmbeddedImg = New System.Net.Mail.LinkedResource(tempFile) With {
                                .ContentId = message.EMailAttachments(MyCounter).PlaceholderInMhtmlToBeReplacedByContentID
                            }
                        Else
                            MyAttachment = New System.Net.Mail.Attachment(tempFile)
                        End If
                    Catch ex As Exception
                        'Don't return immediately for there might be some temp files to be cleaned up, first
                        ErrorFound = ex
                    End Try
                ElseIf message.emailAttachments(MyCounter).FilePath <> Nothing Then
                    Dim fi As New IO.FileInfo(message.EMailAttachments(MyCounter).FilePath)
                    If fi.Exists Then
                        Try
                            If message.EMailAttachments(MyCounter).PlaceholderInMhtmlToBeReplacedByContentID <> Nothing Then
                                MyEmbeddedImg = New System.Net.Mail.LinkedResource(message.EMailAttachments(MyCounter).FilePath) With {
                                    .ContentId = message.EMailAttachments(MyCounter).PlaceholderInMhtmlToBeReplacedByContentID
                                }
                            Else
                                MyAttachment = New System.Net.Mail.Attachment(message.EMailAttachments(MyCounter).FilePath)
                            End If
                        Catch ex As Exception
                            Return New SmtpSendResult(ex)
                        End Try
                    Else
                        Return New SmtpSendResult(New Exception("Attachment not found: " & message.EMailAttachments(MyCounter).FilePath))
                    End If
                Else
                    'empty attachment - ignore
                End If
                If Not MyEmbeddedImg Is Nothing Then
                    'add the LinkedResource to the appropriate view
                    htmlView.LinkedResources.Add(MyEmbeddedImg)
                End If
                If Not MyAttachment Is Nothing Then
                    MyMail.Attachments.Add(MyAttachment)
                End If
            Next

            Try
                If plainView IsNot Nothing Then
                    MyMail.AlternateViews.Add(plainView)
                End If
                If htmlView IsNot Nothing Then
                    MyMail.AlternateViews.Add(htmlView)
                End If

                'Setup mail priority
                MyMail.Headers.Add("X-Priority", CByte(message.Priority).ToString())
                If message.Priority = EMails.Priority.High OrElse message.Priority = EMails.Priority.Normal OrElse message.Priority = EMails.Priority.Low Then
                    MyMail.Headers.Add("Importance", MyMail.Priority.ToString())
                    MyMail.Headers.Add("Priority", MyMail.Priority.ToString())
                    MyMail.Headers.Add("X-MSMail-Priority", MyMail.Priority.ToString())
                End If
                Select Case message.Priority
                    Case EMails.Priority.High, EMails.Priority.Higher
                        MyMail.Priority = System.Net.Mail.MailPriority.High
                    Case EMails.Priority.Low, EMails.Priority.Lower
                        MyMail.Priority = System.Net.Mail.MailPriority.Low
                    Case Else
                        MyMail.Priority = System.Net.Mail.MailPriority.Normal
                End Select

                'Setup mail sensitivity
                Select Case message.Sensitivity
                    Case EMails.Sensitivity.CompanyConfidential
                        MyMail.Headers.Add("Sensitivity", "Company-Confidential")
                    Case EMails.Sensitivity.Personal
                        MyMail.Headers.Add("Sensitivity", "Personal")
                    Case EMails.Sensitivity.Private
                        MyMail.Headers.Add("Sensitivity", "Private")
                    Case Else
                        'Normal
                End Select

                'Setup mail receipt confirmations
                If message.RequestReadingConfirmation Then
                    'Disposition-Notification-To: "Jon Doe" <jon.doe@my-webmailer-company.com> 'Lesebestätigung
                    MyMail.Headers.Add("Disposition-Notification-To", message.From.ToString)
                End If

                'MIGHT BE FIXED OR IS STILL TODO: The bug is that nothing happens :( 
                If message.RequestTransmissionConfirmation Then
                    'Return-Receipt-To: "Jochen Wezel" <my-company@my-webmailer-company.com> 'Übermittlungsbestätigung
                    'MyMail.Headers.Add("Return-Receipt-To", SenderAddress)
                    MyMail.DeliveryNotificationOptions = System.Net.Mail.DeliveryNotificationOptions.OnSuccess
                End If

                'Setup additional headers
                For Each MyKey As String In message.AdditionalHeaders
                    If message.AdditionalHeaders(MyKey) <> Nothing AndAlso message.AdditionalHeaders(MyKey) <> "<>" Then 'don't add invalid header value
                        MyMail.Headers.Add(MyKey, message.AdditionalHeaders(MyKey))
                    End If
                Next

                'Send the mail and return the success/error status
                If ErrorFound IsNot Nothing Then
                    Smtp.Send(MyMail)
                    Result = New SmtpSendResult()
                Else
                    Result = New SmtpSendResult(ErrorFound)
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