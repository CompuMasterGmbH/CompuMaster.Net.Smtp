Option Explicit On
Option Strict On

Imports CompuMaster.Data

''' <summary>
'''     The content of an e-mail with all required data
''' </summary>
''' <remarks>
'''     ToDo: some fields are not yet use in queueing and queue send method, implement them everywhere
''' </remarks>
Public Class EMailMessage

    Public Sub New()
    End Sub

    Public Sub New(messageXml As String)
        Me.LoadFromXml(messageXml)
    End Sub

    Public Sub New(recipient As EMailRecipient, subject As String, plainTextBody As String, htmlBody As String, sender As EMailRecipient, replyTo As EMailRecipient)
        Me.New(New List(Of EMailRecipient)(New EMailRecipient() {recipient}), New List(Of EMailRecipient), New List(Of EMailRecipient), subject, plainTextBody, htmlBody, sender, replyTo, System.Text.Encoding.GetEncoding("UTF-8"), New List(Of EMailAttachment), EMails.Priority.Normal, EMails.Sensitivity.Normal, False, False, New Specialized.NameValueCollection)
    End Sub

    Public Sub New(recipientsTo As List(Of EMailRecipient), recipientsCc As List(Of EMailRecipient), recipientsBcc As List(Of EMailRecipient), subject As String, plainTextBody As String, htmlBody As String, sender As EMailRecipient, replyTo As EMailRecipient)
        Me.New(recipientsTo, recipientsCc, recipientsBcc, subject, plainTextBody, htmlBody, sender, replyTo, System.Text.Encoding.GetEncoding("UTF-8"), New List(Of EMailAttachment), EMails.Priority.Normal, EMails.Sensitivity.Normal, False, False, New Specialized.NameValueCollection)
    End Sub

    Public Sub New(recipient As EMailRecipient, subject As String, plainTextBody As String, htmlBody As String, sender As EMailRecipient, replyTo As EMailRecipient, messageEncoding As System.Text.Encoding, attachments As List(Of EMailAttachment), priority As EMails.Priority, sensitivity As EMails.Sensitivity, requestTransmissionConfirmation As Boolean, requestReadingConfirmation As Boolean, additionalHeaders As Collections.Specialized.NameValueCollection)
        Me.New(New List(Of EMailRecipient)(New EMailRecipient() {recipient}), New List(Of EMailRecipient), New List(Of EMailRecipient), subject, plainTextBody, htmlBody, sender, replyTo, messageEncoding, attachments, priority, sensitivity, requestTransmissionConfirmation, requestReadingConfirmation, additionalHeaders)
    End Sub

    Public Sub New(recipientsTo As List(Of EMailRecipient), recipientsCc As List(Of EMailRecipient), recipientsBcc As List(Of EMailRecipient), subject As String, plainTextBody As String, htmlBody As String, sender As EMailRecipient, replyTo As EMailRecipient, messageEncoding As System.Text.Encoding, attachments As List(Of EMailAttachment), priority As EMails.Priority, sensitivity As EMails.Sensitivity, requestTransmissionConfirmation As Boolean, requestReadingConfirmation As Boolean, additionalHeaders As Collections.Specialized.NameValueCollection)
        Me.To = recipientsTo
        Me.Cc = recipientsCc
        Me.Bcc = recipientsBcc
        Me.Subject = subject
        Me.MessageEncoding = messageEncoding
        Me.BodyPlainText = plainTextBody
        Me.BodyHtml = htmlBody
        Me.From = sender
        Me.ReplyTo = replyTo
        Me.EMailAttachments = attachments
        Me.Priority = priority
        Me.Sensitivity = sensitivity
        Me.RequestReadingConfirmation = requestReadingConfirmation
        Me.RequestTransmissionConfirmation = requestTransmissionConfirmation
        Me.AdditionalHeaders = additionalHeaders
    End Sub

    Private _MessageEncoding As System.Text.Encoding
    Public Property MessageEncoding() As System.Text.Encoding
        Get
            If _MessageEncoding Is Nothing Then
                Return System.Text.Encoding.GetEncoding("UTF-8")
            Else
                Return Me._MessageEncoding
            End If
        End Get
        Set(Value As System.Text.Encoding)
            Me._MessageEncoding = Value
        End Set
    End Property

    Public Property RequestReadingConfirmation() As Boolean
    Public Property RequestTransmissionConfirmation() As Boolean
    Public Property AdditionalHeaders() As System.Collections.Specialized.NameValueCollection
    Public Property Sensitivity() As EMails.Sensitivity
    Public Property Priority() As EMails.Priority
    Public Property From() As EMailRecipient
    Public Property ReplyTo() As EMailRecipient
    Public Property [To]() As New List(Of EMailRecipient)
    Public Property Cc() As New List(Of EMailRecipient)
    Public Property Bcc() As New List(Of EMailRecipient)
    Public Property Subject() As String
    Public Property BodyPlainText() As String
    Public Property BodyHtml() As String
    Public Property EMailAttachments() As New List(Of EMailAttachment)

    ''' <summary>
    ''' Export the full e-mail message to XML
    ''' </summary>
    ''' <returns></returns>
    Public Function ExportToXml() As String
        Me.PrepareForSending()

        'Setup additional headers
        If AdditionalHeaders Is Nothing Then
            AdditionalHeaders = New System.Collections.Specialized.NameValueCollection
        End If

        'Collect mail data
        Dim MailData As New DataSet("root")

        'Headers
        Dim Headers As DataTable
        Headers = Data.DataTables.ConvertNameValueCollectionToDataTable(AdditionalHeaders)
        Headers.TableName = "headers"
        MailData.Tables.Add(Headers)

        'Attachments
        Dim AttachedFiles As New DataTable("attachments")
        AttachedFiles.Columns.Add("Placeholder", GetType(String))
        AttachedFiles.Columns.Add("FileName", GetType(String))
        AttachedFiles.Columns.Add("FileData", GetType(Byte()))
        AttachedFiles.Columns.Add("OriginFileNameBeforePlaceholderValue", GetType(String)) 'to allow an attachment with only content-id in filename to be saved to disc with correct file type again (e.g. for mail queue monitor e-mail preview dialog)
        If Not Me.EMailAttachments Is Nothing Then
            For MyCounter As Integer = 0 To Me.EMailAttachments.Count - 1
                If Me.EMailAttachments(MyCounter).FilePath <> Nothing AndAlso Not System.IO.File.Exists(Me.EMailAttachments(MyCounter).FilePath) Then
                    Throw New System.IO.FileNotFoundException("Attachment file not found: " & Me.EMailAttachments(MyCounter).FilePath)
                ElseIf (Me.EMailAttachments(MyCounter).RawData Is Nothing OrElse Me.EMailAttachments(MyCounter).RawDataFilename = Nothing) AndAlso Me.EMailAttachments(MyCounter).FilePath <> Nothing AndAlso System.IO.File.Exists(Me.EMailAttachments(MyCounter).FilePath) Then
                    'Load file system data into memory as raw binary data
                    Dim fs As System.IO.FileStream = New System.IO.FileStream(Me.EMailAttachments(MyCounter).FilePath, IO.FileMode.Open)
                    Try
                        Dim fi As IO.FileInfo = New IO.FileInfo(Me.EMailAttachments(MyCounter).FilePath)
                        Dim byteArr(CType(fi.Length, Integer) - 1) As Byte
                        fs.Read(byteArr, 0, CType(fi.Length, Integer))
                        Me.EMailAttachments(MyCounter).RawData = byteArr
                        Me.EMailAttachments(MyCounter).RawDataFilename = IO.Path.GetFileName(Me.EMailAttachments(MyCounter).FilePath)
                        Me.EMailAttachments(MyCounter).FilePath = Nothing
                    Finally
                        fs.Close()
                    End Try
                ElseIf Not Me.EMailAttachments(MyCounter).RawData Is Nothing AndAlso Me.EMailAttachments(MyCounter).RawDataFilename <> Nothing Then
                    'Just add the binary data
                Else
                    Throw New Exception("Empty or invalid email attachment")
                End If

                'Add the binary data
                Dim MyAttachment As EMailAttachment = Me.EMailAttachments(MyCounter)
                Dim MyRow As DataRow = AttachedFiles.NewRow
                MyRow("FileData") = SmtpUtils.ObjectNotNothingOrDBNull(MyAttachment.RawData)
                MyRow("FileName") = SmtpUtils.StringNotNothingOrDBNull(MyAttachment.RawDataFilename)
                MyRow("Placeholder") = SmtpUtils.StringNotNothingOrDBNull(MyAttachment.PlaceholderInMhtmlToBeReplacedByContentID)
                If MyAttachment.PlaceholderInMhtmlToBeReplacedByContentID <> Nothing Then
                    'Required to drop the origin filename if we want to use a Content-ID
                    MyRow("OriginFileNameBeforePlaceholderValue") = MyRow("FileName")
                    MyRow("FileName") = SmtpUtils.StringNotNothingOrDBNull(MyAttachment.PlaceholderInMhtmlToBeReplacedByContentID)
                End If
                AttachedFiles.Rows.Add(MyRow)
            Next
        End If
        MailData.Tables.Add(AttachedFiles)

        'Basic message data
        Dim MessageData As DataTable
        Dim MessageDataCollection As New Collections.Specialized.NameValueCollection
        MessageDataCollection.Add("From", Me.From.ToString)
        MessageDataCollection.Add("ReplyTo", Me.ReplyTo.ToString)
        MessageDataCollection.Add("To", EMailRecipient.JoinToListWithSmtpFormat(Me.To))
        MessageDataCollection.Add("Cc", EMailRecipient.JoinToListWithSmtpFormat(Me.Cc))
        MessageDataCollection.Add("Bcc", EMailRecipient.JoinToListWithSmtpFormat(Me.Bcc))
        MessageDataCollection.Add("Subject", Me.Subject)
        MessageDataCollection.Add("MessageEncoding", Me.MessageEncoding.EncodingName)
        MessageDataCollection.Add("TextBody", Me.BodyPlainText)
        MessageDataCollection.Add("HtmlBody", Me.BodyHtml)
        MessageDataCollection.Add("Priority", CType(Me.Priority, Byte).ToString)
        MessageDataCollection.Add("Sensitivity", CType(Me.Sensitivity, Byte).ToString)
        MessageDataCollection.Add("RequestTransmissionConfirmation", Me.RequestTransmissionConfirmation.ToString)
        MessageDataCollection.Add("RequestReadingConfirmation", Me.RequestReadingConfirmation.ToString)
        MessageData = Data.DataTables.ConvertNameValueCollectionToDataTable(MessageDataCollection)
        MessageData.TableName = "message"
        MailData.Tables.Add(MessageData)

        Return Data.DataTables.ConvertDatasetToXml(MailData)
    End Function

    Public Sub LoadFromXml(messageXml As String)
        Dim MailData As DataSet = Data.DataTables.ConvertXmlToDataset(messageXml)
        Dim MessageData As DataTable = MailData.Tables("message")
        Dim Attachments As DataTable = MailData.Tables("attachments")
        Dim Headers As DataTable = MailData.Tables("headers")

        'Restore header information
        Dim additionalHeaders As New Specialized.NameValueCollection
        For MyCounter As Integer = 0 To Headers.Rows.Count - 1
            additionalHeaders.Add(CType(Headers.Rows(MyCounter)("key"), String), CType(Headers.Rows(MyCounter)("value"), String))
        Next

        'Restore attachments
        Dim mailAttachments As New List(Of EMailAttachment)
        For MyCounter As Integer = 0 To Attachments.Rows.Count - 1
            Dim EMailAttachment As New EMailAttachment With {
                .PlaceholderInMhtmlToBeReplacedByContentID = Data.Utils.NoDBNull(Attachments.Rows(MyCounter)("Placeholder"), CType(Nothing, String)),
                .RawData = CType(Data.Utils.NoDBNull(Attachments.Rows(MyCounter)("FileData")), Byte()),
                .RawDataFilename = Data.Utils.NoDBNull(Attachments.Rows(MyCounter)("FileName"), CType(Nothing, String))
            }
            mailAttachments.Add(EMailAttachment)
        Next

        'Restore message data
        For MyCounter As Integer = 0 To MessageData.Rows.Count - 1
            Select Case LCase(CType(MessageData.Rows(MyCounter)("key"), String))
                Case "from"
                    Me.From = EMailRecipient.CreateFromSmtpFormat(Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String)))
                Case "replyto"
                    Me.ReplyTo = EMailRecipient.CreateFromSmtpFormat(Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String)))
                Case "to"
                    Me.To = EMailRecipient.SplitEMailAddressesIntoEMailRecipientsFromRecipientsList(Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String)))
                Case "cc"
                    Me.Cc = EMailRecipient.SplitEMailAddressesIntoEMailRecipientsFromRecipientsList(Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String)))
                Case "bcc"
                    Me.Bcc = EMailRecipient.SplitEMailAddressesIntoEMailRecipientsFromRecipientsList(Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String)))
                Case "subject"
                    Me.Subject = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
                Case "messageencoding"
                    Dim Charset As String = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
                    If Charset <> Nothing Then Me.MessageEncoding = System.Text.Encoding.GetEncoding(Charset)
                Case "textbody"
                    Me.BodyPlainText = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
                Case "htmlbody"
                    Me.BodyHtml = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
                Case Else
                    Throw New NotSupportedException("Invalid messagedata key: " & CType(MessageData.Rows(MyCounter)("key"), String))
            End Select
        Next

        'Now, fill the class instance
        Me.EMailAttachments = mailAttachments
        Me.AdditionalHeaders = additionalHeaders
    End Sub

    ''' <summary>
    ''' Prepare HTML and attachments for embedded images feature, initialize all list fields if null/Nothing
    ''' </summary>
    Friend Sub PrepareForSending()
        'Fix HTML and attachments
        EMailAttachment.FixHtmlContentIDs(Me.BodyHtml, Me.EMailAttachments)
        'Init missing list instances
        If Me.AdditionalHeaders Is Nothing Then Me.AdditionalHeaders = New Specialized.NameValueCollection
        If Me.To Is Nothing Then Me.To = New List(Of EMailRecipient)
        If Me.Cc Is Nothing Then Me.Cc = New List(Of EMailRecipient)
        If Me.Bcc Is Nothing Then Me.Bcc = New List(Of EMailRecipient)
        If Me.EMailAttachments Is Nothing Then Me.EMailAttachments = New List(Of EMailAttachment)
        'Validate assigned values
        Select Case Me.Priority
            Case EMails.Priority.Low
            Case EMails.Priority.Lower
            Case EMails.Priority.High
            Case EMails.Priority.Higher
            Case Else
                Me.Priority = EMails.Priority.Normal
        End Select
        Select Case Sensitivity
            Case EMails.Sensitivity.CompanyConfidential
            Case EMails.Sensitivity.Personal
            Case EMails.Sensitivity.Private
            Case Else
                Me.Sensitivity = EMails.Sensitivity.Normal
        End Select
    End Sub

End Class
