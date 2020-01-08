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

    Friend Const CompatibilityVersion As String = "2020.01"

    Public Sub New()
    End Sub

    Friend Sub New(ByVal messageXml As String)
        Me.LoadFromXml(messageXml)
    End Sub

    Private _Charset As String
    Public Property Charset() As String
        Get
            Return Me._Charset
        End Get
        Set(ByVal Value As String)
            Me._Charset = Value
        End Set
    End Property

    Private _RequestReadingConfirmation As Boolean
    Public Property RequestReadingConfirmation() As Boolean
        Get
            Return Me._RequestReadingConfirmation
        End Get
        Set(ByVal Value As Boolean)
            Me._RequestReadingConfirmation = Value
        End Set
    End Property

    Private _RequestTransmissionConfirmation As Boolean
    Public Property RequestTransmissionConfirmation() As Boolean
        Get
            Return Me._RequestTransmissionConfirmation
        End Get
        Set(ByVal Value As Boolean)
            Me._RequestTransmissionConfirmation = Value
        End Set
    End Property

    Private _AdditionalHeaders As System.Collections.Specialized.NameValueCollection
    Public Property AdditionalHeaders() As System.Collections.Specialized.NameValueCollection
        Get
            Return Me._AdditionalHeaders
        End Get
        Set(ByVal Value As System.Collections.Specialized.NameValueCollection)
            Me._AdditionalHeaders = Value
        End Set
    End Property

    Private _Sensitivity As EMails.Sensitivity
    Public Property Sensitivity() As EMails.Sensitivity
        Get
            Return Me._Sensitivity
        End Get
        Set(ByVal Value As EMails.Sensitivity)
            Me._Sensitivity = Value
        End Set
    End Property

    Private _Priority As EMails.Priority
    Public Property Priority() As EMails.Priority
        Get
            Return Me._Priority
        End Get
        Set(ByVal Value As EMails.Priority)
            Me._Priority = Value
        End Set
    End Property

    Private _FromAddress As String
    Public Property FromAddress() As String
        Get
            Return Me._FromAddress
        End Get
        Set(ByVal Value As String)
            Me._FromAddress = Value
        End Set
    End Property

    Private _FromName As String
    Public Property FromName() As String
        Get
            Return Me._FromName
        End Get
        Set(ByVal Value As String)
            Me._FromName = Value
        End Set
    End Property

    Private _To As String
    Public Property [To]() As String
        Get
            Return Me._To
        End Get
        Set(ByVal Value As String)
            Me._To = Value
        End Set
    End Property

    Private _Cc As String
    Public Property Cc() As String
        Get
            Return Me._Cc
        End Get
        Set(ByVal Value As String)
            Me._Cc = Value
        End Set
    End Property

    Private _Bcc As String
    Public Property Bcc() As String
        Get
            Return _Bcc
        End Get
        Set(ByVal Value As String)
            _Bcc = Value
        End Set
    End Property

    Private _Subject As String
    Public Property Subject() As String
        Get
            Return _Subject
        End Get
        Set(ByVal Value As String)
            _Subject = Value
        End Set
    End Property

    Private _BodyPlainText As String
    Public Property BodyPlainText() As String
        Get
            Return _BodyPlainText
        End Get
        Set(ByVal Value As String)
            _BodyPlainText = Value
        End Set
    End Property

    Private _BodyHtml As String
    Public Property BodyHtml() As String
        Get
            Return _BodyHtml
        End Get
        Set(ByVal Value As String)
            _BodyHtml = Value
        End Set
    End Property

    Private _EMailAttachments As EMailAttachment()
    Public Property EMailAttachments() As EMailAttachment()
        Get
            If _EMailAttachments Is Nothing Then
                _EMailAttachments = New EMailAttachment() {}
            End If
            Return _EMailAttachments
        End Get
        Set(ByVal Value As EMailAttachment())
            _EMailAttachments = Value
        End Set
    End Property

    Friend Shared Function CreateQueueMessageXml(ByVal rcptAddresses_To As String, ByVal rcptAddresses_CC As String, ByVal rcptAddresses_BCC As String, ByVal msgSubject As String, ByVal msgTextBody As String, ByVal msgHTMLBody As String, ByVal senderName As String, ByVal senderAddress As String, ByVal msgCharset As String, ByVal attachments() As EMailAttachment, ByVal priority As EMails.Priority, ByVal sensitivity As EMails.Sensitivity, ByVal requestTransmissionConfirmation As Boolean, ByVal requestReadingConfirmation As Boolean, ByVal additionalHeaders As Collections.Specialized.NameValueCollection) As System.Data.DataSet
        EMailAttachment.FixHtmlContentIDs(msgHTMLBody, attachments)

        'Setup additional headers
        If additionalHeaders Is Nothing Then
            additionalHeaders = New System.Collections.Specialized.NameValueCollection
        End If

        'Setup mail priority
        If Not priority = Nothing Then
            Select Case priority
                Case EMails.Priority.Low
                Case CType(4, EMails.Priority) 'Between Normal and Low
                Case EMails.Priority.High
                Case CType(2, EMails.Priority) 'Between Normal and High
                Case Else 'WMSystem.MailImportance.Normal
                    priority = EMails.Priority.Normal
            End Select
            additionalHeaders("X-Priority") = CByte(priority).ToString()
            If priority = EMails.Priority.High OrElse priority = EMails.Priority.Normal OrElse priority = EMails.Priority.Low Then
                additionalHeaders("X-MSMail-Priority") = priority.ToString()
                additionalHeaders("Importance") = priority.ToString()
            End If
        End If

        'Setup mail sensitivity
        If Not sensitivity = Nothing Then
            Select Case sensitivity
                Case EMails.Sensitivity.CompanyConfidential
                    additionalHeaders("Sensitivity") = "Company-Confidential"
                Case EMails.Sensitivity.Personal
                    additionalHeaders("Sensitivity") = "Personal"
                Case EMails.Sensitivity.Private
                    additionalHeaders("Sensitivity") = "Private"
            End Select
        End If

        'Setup mail receipt confirmations
        If requestReadingConfirmation Then
            additionalHeaders("Disposition-Notification-To") = senderAddress
        End If
        If requestTransmissionConfirmation Then
            additionalHeaders("Return-Receipt-To") = """" & senderName & """ <" & senderAddress & ">"
        End If

        'Collect mail data
        Dim MailData As New DataSet("root")

        'Version information
        Dim VersionData As DataTable
        Dim VersionDataHashtable As New Dictionary(Of String, String)
        VersionDataHashtable.Add("NetLibrary", CompatibilityVersion)
        VersionData = Data.DataTables.ConvertIDictionaryToDataTable(VersionDataHashtable)
        VersionData.TableName = "version"
        MailData.Tables.Add(VersionData)

        'Headers
        Dim Headers As DataTable
        Headers = Data.DataTables.ConvertNameValueCollectionToDataTable(additionalHeaders)
        Headers.TableName = "headers"
        MailData.Tables.Add(Headers)

        'Attachments
        Dim AttachedFiles As New DataTable("attachments")
        AttachedFiles.Columns.Add("Placeholder", GetType(String))
        AttachedFiles.Columns.Add("FileName", GetType(String))
        AttachedFiles.Columns.Add("FileData", GetType(Byte()))
        AttachedFiles.Columns.Add("OriginFileNameBeforePlaceholderValue", GetType(String)) 'to allow an attachment with only content-id in filename to be saved to disc with correct file type again (e.g. for mail queue monitor e-mail preview dialog)
        If Not attachments Is Nothing Then
            For MyCounter As Integer = 0 To attachments.Length - 1

                If attachments(MyCounter).FilePath <> Nothing AndAlso Not System.IO.File.Exists(attachments(MyCounter).FilePath) Then
                    Throw New System.IO.FileNotFoundException("Attachment file not found: " & attachments(MyCounter).FilePath)
                ElseIf (attachments(MyCounter).RawData Is Nothing OrElse attachments(MyCounter).RawDataFilename = Nothing) AndAlso attachments(MyCounter).FilePath <> Nothing AndAlso System.IO.File.Exists(attachments(MyCounter).FilePath) Then
                    'Load file system data into memory as raw binary data
                    Dim fs As System.IO.FileStream = New System.IO.FileStream(attachments(MyCounter).FilePath, IO.FileMode.Open)
                    Try
                        Dim fi As IO.FileInfo = New IO.FileInfo(attachments(MyCounter).FilePath)
                        Dim byteArr(CType(fi.Length, Integer) - 1) As Byte
                        fs.Read(byteArr, 0, CType(fi.Length, Integer))
                        attachments(MyCounter).RawData = byteArr
                        attachments(MyCounter).RawDataFilename = IO.Path.GetFileName(attachments(MyCounter).FilePath)
                        attachments(MyCounter).FilePath = Nothing
                    Finally
                        fs.Close()
                    End Try
                ElseIf Not attachments(MyCounter).RawData Is Nothing AndAlso attachments(MyCounter).RawDataFilename <> Nothing Then
                    'Just add the binary data
                Else
                    Throw New Exception("Empty or invalid email attachment")
                End If

                'Add the binary data
                Dim MyAttachment As EMailAttachment = attachments(MyCounter)
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
        MessageDataCollection.Add("FromAddress", senderAddress)
        MessageDataCollection.Add("FromName", senderName)
        MessageDataCollection.Add("To", rcptAddresses_To)
        MessageDataCollection.Add("Cc", rcptAddresses_CC)
        MessageDataCollection.Add("Bcc", rcptAddresses_BCC)
        MessageDataCollection.Add("Subject", msgSubject)
        MessageDataCollection.Add("Charset", msgCharset)
        MessageDataCollection.Add("TextBody", msgTextBody)
        MessageDataCollection.Add("HtmlBody", msgHTMLBody)
        MessageData = Data.DataTables.ConvertNameValueCollectionToDataTable(MessageDataCollection)
        MessageData.TableName = "message"
        MailData.Tables.Add(MessageData)

        Return MailData
    End Function

    ''' <summary>
    ''' Export the full e-mail message to XML
    ''' </summary>
    ''' <returns></returns>
    Public Function ExportToXml() As String
        Return Data.DataTables.ConvertDatasetToXml(CreateQueueMessageXml(
            Me.To,
            Me.Cc,
            Me.Bcc,
            Me.Subject,
            Me.BodyPlainText,
            Me.BodyHtml,
            Me.FromName,
            Me.FromAddress,
            Me.Charset,
            Me.EMailAttachments,
            Me.Priority,
            Me.Sensitivity,
            Me.RequestTransmissionConfirmation,
            Me.RequestReadingConfirmation,
            Me.AdditionalHeaders
            ))
    End Function

    Public Sub LoadFromXml(ByVal messageXml As String)
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
        Dim FromName As String = Nothing, FromAddress As String = Nothing, RcptTo As String = Nothing, RcptCc As String = Nothing, RcptBcc As String = Nothing
        Dim Subject As String = Nothing, Charset As String = Nothing, BodyText As String = Nothing, BodyHtml As String = Nothing
        For MyCounter As Integer = 0 To MessageData.Rows.Count - 1
            Select Case LCase(CType(MessageData.Rows(MyCounter)("key"), String))
                Case "fromname"
                    FromName = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
                Case "fromaddress"
                    FromAddress = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
                Case "to"
                    RcptTo = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
                Case "cc"
                    RcptCc = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
                Case "bcc"
                    RcptBcc = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
                Case "subject"
                    Subject = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
                Case "charset"
                    Charset = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
                Case "textbody"
                    BodyText = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
                Case "htmlbody"
                    BodyHtml = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
                Case Else
                    Throw New NotSupportedException("Invalid messagedata key: " & CType(MessageData.Rows(MyCounter)("key"), String))
            End Select
        Next

        'Now, send the e-mail and return
        Dim EMailAttachments As EMailAttachment() = CType(mailAttachments.ToArray, EMailAttachment())

        'Now, fill the class instance
        Me.To = RcptTo
        Me.Cc = RcptCc
        Me.Bcc = RcptBcc
        Me.Subject = Subject
        Me.BodyPlainText = BodyText
        Me.BodyHtml = BodyHtml
        Me.FromName = FromName
        Me.FromAddress = FromAddress
        Me.EMailAttachments = EMailAttachments
        Me.AdditionalHeaders = additionalHeaders
        Me.Charset = Charset

    End Sub

End Class
