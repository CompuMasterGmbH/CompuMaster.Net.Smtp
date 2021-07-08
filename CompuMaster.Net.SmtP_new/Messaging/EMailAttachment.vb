Option Explicit On
Option Strict On

''' <summary>
'''     An e-mail attachment
''' </summary>
Public Class EMailAttachment
    Public Sub New()
        'Do nothing
    End Sub
    Public Sub New(filepath As String)
        Me.FilePath = filepath
    End Sub
    Public Sub New(filepath As String, placeholderName As String)
        Me.FilePath = filepath
        Me.PlaceholderInMhtmlToBeReplacedByContentID = placeholderName
    End Sub
    Public Sub New(data As Byte(), filename As String)
        Me.RawData = data
        Me.RawDataFilename = filename
    End Sub
    Public Sub New(data As Byte(), filename As String, placeholderName As String)
        Me.RawData = data
        Me.RawDataFilename = filename
        Me.PlaceholderInMhtmlToBeReplacedByContentID = placeholderName
    End Sub

    Private _AttachmentData As Byte()
    ''' <summary>
    '''     Binary data for this attachment
    ''' </summary>
    Public Property RawData() As Byte()
        Get
            Return _AttachmentData
        End Get
        Set(Value As Byte())
            _AttachmentData = Value
        End Set
    End Property

    Private _AttachmentData_Filename As String
    ''' <summary>
    '''     The filename for the binary data as used in e-mail (might be a content ID)
    ''' </summary>
    Public Property RawDataFilename() As String
        Get
            If _AttachmentData_Filename = Nothing Then
                _AttachmentData_Filename = _RawDataOriginFilename
            End If
            Return _AttachmentData_Filename
        End Get
        Set(Value As String)
            _AttachmentData_Filename = Value
            If _RawDataOriginFilename = Nothing Then
                _RawDataOriginFilename = Value
            End If
        End Set
    End Property

    Private _RawDataOriginFilename As String
    ''' <summary>
    ''' The origin filename for the binary data (as used on disk)
    ''' </summary>
    ''' <returns></returns>
    Public Property RawDataOriginFilename As String
        Get
            If _RawDataOriginFilename = Nothing Then
                _RawDataOriginFilename = _AttachmentData_Filename
            End If
            Return _RawDataOriginFilename
        End Get
        Set(value As String)
            _RawDataOriginFilename = value
            If _AttachmentData_Filename = Nothing Then
                _AttachmentData_Filename = value
            End If
        End Set
    End Property

    Private _AttachmentFile As String
    ''' <summary>
    '''     A path to a file which shall be included
    ''' </summary>
    Public Property FilePath() As String
        Get
            Return _AttachmentFile
        End Get
        Set(Value As String)
            _AttachmentFile = Value
        End Set
    End Property

    Private _PlaceholderInMHTML_ToReplaceWithCID As String ' <Obsolete("Use ContentID instead", False), System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)> _
    ''' <summary>
    '''     A placeholder string (without prefix "cid:") in the HTML code of the message (there it must be with prefix "cid:") which shall be replaced by the CID code of the attachment
    ''' </summary>
    ''' <remarks>
    ''' <para>Define the placeholder which shall be replaced by the Content-ID for the contents of a file to the email. Emails formatted in HTML can include images with this information and internally reference the image through a "cid" hyperlink.</para>
    ''' </remarks>
    Public Property PlaceholderInMhtmlToBeReplacedByContentID() As String
        Get
            Return _PlaceholderInMHTML_ToReplaceWithCID
        End Get
        Set(Value As String)
            _PlaceholderInMHTML_ToReplaceWithCID = Value
        End Set
    End Property

    ''' <summary>
    ''' Replace all Content-IDs by encoding-independent Content-IDs and cleanup all attachments with placeholder but without reference from message HTML
    ''' </summary>
    ''' <param name="htmlBody">The e-mail HTML body</param>
    ''' <param name="attachments">The attachments of the e-mail</param>
    Friend Shared Sub FixHtmlContentIDs(ByRef htmlBody As String, attachments As List(Of EMailAttachment))
        If attachments IsNot Nothing Then
            For MyCounter As Integer = 0 To attachments.Count - 1
                If attachments(MyCounter).PlaceholderInMhtmlToBeReplacedByContentID <> Nothing Then
                    'Ensure working content-ID in all encodings
                    'If suggested Content-ID is not possible for usage without a working encoding (at least it's unknown how to do it, currently), then replace the Content-ID by another value and also do this replacement in the html message part
                    Dim IsGuid As Boolean = attachments(MyCounter).PlaceholderInMhtmlToBeReplacedByContentID <> Nothing AndAlso Guid.TryParse(attachments(MyCounter).PlaceholderInMhtmlToBeReplacedByContentID, Nothing) AndAlso Guid.Parse(attachments(MyCounter).PlaceholderInMhtmlToBeReplacedByContentID) <> Nothing
                    Dim newContentIDName As String
                    If IsGuid = False Then
                        newContentIDName = Guid.NewGuid.ToString
                        'Replace exact matches
                        htmlBody = SmtpUtils.ReplaceString(htmlBody, attachments(MyCounter).PlaceholderInMhtmlToBeReplacedByContentID, newContentIDName, SmtpUtils.ReplaceComparisonTypes.InvariantCultureIgnoreCase)
                        'Replace html- and url-encoded matches
                        htmlBody = SmtpUtils.ReplaceString(htmlBody, System.Net.WebUtility.UrlEncode(attachments(MyCounter).PlaceholderInMhtmlToBeReplacedByContentID), newContentIDName, SmtpUtils.ReplaceComparisonTypes.InvariantCultureIgnoreCase)
                        htmlBody = SmtpUtils.ReplaceString(htmlBody, System.Net.WebUtility.HtmlEncode(attachments(MyCounter).PlaceholderInMhtmlToBeReplacedByContentID), newContentIDName, SmtpUtils.ReplaceComparisonTypes.InvariantCultureIgnoreCase)
                        htmlBody = SmtpUtils.ReplaceString(htmlBody, System.Net.WebUtility.HtmlEncode(System.Net.WebUtility.UrlEncode(attachments(MyCounter).PlaceholderInMhtmlToBeReplacedByContentID)), newContentIDName, SmtpUtils.ReplaceComparisonTypes.InvariantCultureIgnoreCase)
                        attachments(MyCounter).PlaceholderInMhtmlToBeReplacedByContentID = newContentIDName
                    End If
                End If
            Next

            'remove attachments with missing references
            For MyCounter As Integer = attachments.Count - 1 To 0 Step -1
                If attachments(MyCounter).PlaceholderInMhtmlToBeReplacedByContentID <> Nothing AndAlso htmlBody.IndexOf(attachments(MyCounter).PlaceholderInMhtmlToBeReplacedByContentID) = -1 Then
                    'html doesn't contain any reference to the embedded object - removing attachment
                    attachments.RemoveAt(MyCounter)
                End If
            Next
        End If
    End Sub

End Class