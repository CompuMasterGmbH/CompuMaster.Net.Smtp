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

    'Sub New(ByVal xml As String, ByVal webManager As WMSystem)
    '    Dim MailData As DataSet = Data.DataTables.ConvertXmlToDataset(xml)
    '    Dim MessageData As DataTable = MailData.Tables("message")
    '    Dim Attachments As DataTable = MailData.Tables("attachments")
    '    Dim Headers As DataTable = MailData.Tables("headers")

    '    'Restore header information
    '    Dim additionalHeaders As New Specialized.NameValueCollection
    '    For MyCounter As Integer = 0 To Headers.Rows.Count - 1
    '        additionalHeaders.Add(CType(Headers.Rows(MyCounter)("key"), String), CType(Headers.Rows(MyCounter)("value"), String))
    '    Next

    '    'Restore attachments
    '    Dim mailAttachments As New ArrayList
    '    For MyCounter As Integer = 0 To Attachments.Rows.Count - 1
    '        Dim EMailAttachment As New EMailAttachment
    '        EMailAttachment.PlaceholderInMhtmlToBeReplacedByContentID = Data.Utils.NoDBNull(Attachments.Rows(MyCounter)("Placeholder"), CType(Nothing, String))
    '        EMailAttachment.RawData = CType(Data.Utils.NoDBNull(Attachments.Rows(MyCounter)("FileData")), Byte())
    '        EMailAttachment.RawDataFilename = Data.Utils.NoDBNull(Attachments.Rows(MyCounter)("FileName"), CType(Nothing, String))
    '        mailAttachments.Add(EMailAttachment)
    '    Next

    '    'Restore message data
    '    Dim FromName As String = Nothing, FromAddress As String = Nothing, RcptTo As String = Nothing, RcptCc As String = Nothing, RcptBcc As String = Nothing
    '    Dim Subject As String = Nothing, Charset As String = Nothing, BodyText As String = Nothing, BodyHtml As String = Nothing
    '    For MyCounter As Integer = 0 To MessageData.Rows.Count - 1
    '        Select Case LCase(CType(MessageData.Rows(MyCounter)("key"), String))
    '            Case "fromname"
    '                FromName = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
    '            Case "fromaddress"
    '                FromAddress = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
    '            Case "to"
    '                RcptTo = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
    '            Case "cc"
    '                RcptCc = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
    '            Case "bcc"
    '                RcptBcc = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
    '            Case "subject"
    '                Subject = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
    '            Case "charset"
    '                Charset = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
    '            Case "textbody"
    '                BodyText = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
    '            Case "htmlbody"
    '                BodyHtml = Data.Utils.NoDBNull(MessageData.Rows(MyCounter)("value"), CType(Nothing, String))
    '            Case Else
    '                'Invalid
    '                'System.Environment.StackTrace doesn't work with medium-trust --> work around it using a new exception class
    '                Dim WorkaroundEx As New Exception("")
    '                Dim WorkaroundStackTrace As String = WorkaroundEx.StackTrace 'contains only last few lines of stacktrace
    '                Try
    '                    WorkaroundStackTrace = System.Environment.StackTrace 'contains full stacktrace
    '                Catch
    '                End Try
    '        End Select
    '    Next

    '    'Now, send the e-mail and return
    '    Dim EMailAttachments As EMailAttachment() = CType(mailAttachments.ToArray(GetType(EMailAttachment)), EMailAttachment())

    '    'Now, fill the class instance
    '    Me.To = RcptTo
    '    Me.Cc = RcptCc
    '    Me.Bcc = RcptBcc
    '    Me.Subject = Subject
    '    Me.BodyPlainText = BodyText
    '    Me.BodyHtml = BodyHtml
    '    Me.FromName = FromName
    '    Me.FromAddress = FromAddress
    '    Me.EMailAttachments = EMailAttachments
    '    Me.AdditionalHeaders = additionalHeaders
    '    Me.Charset = Charset

    'End Sub

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

End Class
