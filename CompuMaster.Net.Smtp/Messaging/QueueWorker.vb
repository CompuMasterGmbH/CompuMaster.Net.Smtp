Option Explicit On
Option Strict On

Imports System.Data.SqlClient

Public Class MailQueue

    Private QueueTransactionData As ArrayList
    ''' <summary>
    '''     Start a transaction for queuing mails
    ''' </summary>
    Public Sub BeginQueueTransaction()
        If Not QueueTransactionData Is Nothing Then
            RollbackQueueTransaction()
        End If
        QueueTransactionData = New ArrayList
    End Sub
    ''' <summary>
    '''     Commit all changes in the transaction
    ''' </summary>
    Public Sub CommitQueueTransaction()
        CommitQueueTransaction(Nothing)
    End Sub
    ''' <summary>
    '''     Commit all changes in the transaction
    ''' </summary>
    ''' <param name="connection">An open connection to the camm Web-Manager database</param>
    Public Sub CommitQueueTransaction(ByVal connection As IDbConnection)

        Dim ReUseDBConnection As Boolean = True
        If connection Is Nothing Then
            ReUseDBConnection = False
            connection = New SqlConnection(_WebManager.ConnectionString)
        End If

        Dim MyCmd As New SqlCommand
        MyCmd.Connection = CType(connection, SqlConnection)
        MyCmd.CommandText =
                    "CREATE TABLE #mqueuetransids (ID int)" & vbNewLine &
                    Me.QueueTransactionSqlRelatedIDs & vbNewLine &
                    "UPDATE dbo.Log_eMailMessages SET State = " & CType(QueueMonitoring.QueueStates.Queued, Byte).ToString & " FROM dbo.Log_eMailMessages INNER JOIN #mqueuetransids ON dbo.Log_eMailMessages.ID = #mqueuetransids.ID" & vbNewLine &
                    "DROP TABLE #mqueuetransids"
        MyCmd.CommandType = CommandType.Text

        If ReUseDBConnection = True Then
            Data.DataQuery.AnyIDataProvider.ExecuteNonQuery(MyCmd, Data.DataQuery.AnyIDataProvider.Automations.None)
        Else
            Data.DataQuery.AnyIDataProvider.ExecuteNonQuery(MyCmd, Data.DataQuery.AnyIDataProvider.Automations.AutoOpenAndCloseAndDisposeConnection)
        End If

    End Sub
    ''' <summary>
    '''     Roll back all changes in the transaction
    ''' </summary>
    Public Sub RollbackQueueTransaction()
        RollbackQueueTransaction(Nothing)
    End Sub
    ''' <summary>
    '''     Roll back all changes in the transaction
    ''' </summary>
    ''' <param name="connection">An open connection to the camm Web-Manager database</param>
    Public Sub RollbackQueueTransaction(ByVal connection As IDbConnection)

        Dim ReUseDBConnection As Boolean = True
        If connection Is Nothing Then
            ReUseDBConnection = False
            connection = New SqlConnection(_WebManager.ConnectionString)
        End If

        Dim MyCmd As New SqlCommand
        MyCmd.Connection = CType(connection, SqlConnection)
        MyCmd.CommandText =
                    "CREATE TABLE #mqueuetransids (ID int)" & vbNewLine &
                    Me.QueueTransactionSqlRelatedIDs & vbNewLine &
                    "DELETE dbo.Log_eMailMessages FROM dbo.Log_eMailMessages INNER JOIN #mqueuetransids ON dbo.Log_eMailMessages.ID = #mqueuetransids.ID" & vbNewLine &
                    "DROP TABLE #mqueuetransids"
        MyCmd.CommandType = CommandType.Text
        If ReUseDBConnection = True Then
            Data.DataQuery.AnyIDataProvider.ExecuteNonQuery(MyCmd, Data.DataQuery.AnyIDataProvider.Automations.None)
        Else
            Data.DataQuery.AnyIDataProvider.ExecuteNonQuery(MyCmd, Data.DataQuery.AnyIDataProvider.Automations.AutoOpenAndCloseAndDisposeConnection)
        End If

    End Sub
    ''' <summary>
    '''     Return a piece of sql code for accessing the related IDs to be used in commit or rollback commands
    ''' </summary>
    Private Function QueueTransactionSqlRelatedIDs() As String
        Dim Result As New System.Text.StringBuilder
        For MyCounter As Integer = 0 To Me.QueueTransactionData.Count - 1
            Result.Append("INSERT INTO #mqueuetransids (ID) VALUES (")
            Result.Append(CType(Me.QueueTransactionData(MyCounter), Integer).ToString)
            Result.Append(")")
            Result.Append(vbNewLine)
        Next
        Return Result.ToString
    End Function
    ''' <summary>
    '''     Send an e-mail to the mail queue
    ''' </summary>
    ''' <param name="RcptName">Receipient's name</param>
    ''' <param name="RcptAddress">Receipient's e-mail address</param>
    ''' <param name="MsgSubject">Message subject</param>
    ''' <param name="MsgTextBody">Message plain text body</param>
    ''' <param name="MsgHTMLBody">Message html text body</param>
    ''' <param name="SenderName">The name of the sender</param>
    ''' <param name="SenderAddress">The e-mail address of the sender</param>
    ''' <param name="MsgCharset">The charset of the message, default is UTF-8</param>
    ''' <param name="connection">An open connection to the camm Web-Manager database which shall be used</param>
    ''' <param name="bufErrorDetails">A reference to a string variable which shall contain error information when an exception gets fired</param>
    ''' <param name="Attachments">Array of e-mail attachments</param>
    ''' <param name="Priority">The priority of the e-mail</param>
    ''' <param name="Sensitivity">The sensitivity of the e-mail</param>
    ''' <param name="RequestTransmissionConfirmation">Request a transmission confirmation</param>
    ''' <param name="RequestReadingConfirmation">Request a reading confirmation</param>
    ''' <param name="AdditionalHeaders">Additional headers for the e-mail</param>
    ''' <returns>True if successfull, false for failures (also see bufErrorDetails)</returns>
    Private Function SendViaQueue(ByVal RcptName As String, ByVal RcptAddress As String, ByVal MsgSubject As String, ByVal MsgTextBody As String, ByVal MsgHTMLBody As String, ByVal SenderName As String, ByVal SenderAddress As String, ByVal MsgCharset As String, ByVal connection As IDbConnection, ByRef bufErrorDetails As String, ByVal Attachments() As EMailAttachment, ByVal Priority As EMails.Priority, ByVal Sensitivity As EMails.Sensitivity, ByVal RequestTransmissionConfirmation As Boolean, ByVal RequestReadingConfirmation As Boolean, ByVal AdditionalHeaders As Collections.Specialized.NameValueCollection) As Boolean
        Return SendViaQueue(CreateReceipientString(RcptName, RcptAddress), Nothing, Nothing, MsgSubject, MsgTextBody, MsgHTMLBody, SenderName, SenderAddress, MsgCharset, connection, bufErrorDetails, Attachments, Priority, Sensitivity, RequestTransmissionConfirmation, RequestReadingConfirmation, AdditionalHeaders)
    End Function
    ''' <summary>
    '''     Sends an e-mail to multiple receipients via the mail queue
    ''' </summary>
    ''' <param name="rcptAddresses_To">The receipients comma-separated with format "Complete Name &lt;somebody@yourcompany.com&gt;" or only a simple e-mail address in the format "somebody@yourcompany.com"</param>
    ''' <param name="rcptAddresses_CC">The copy receipients comma-separated with format "Complete Name &lt;somebody@yourcompany.com&gt;" or only a simple e-mail address in the format "somebody@yourcompany.com"</param>
    ''' <param name="rcptAddresses_BCC">The blind copy receipients comma-separated with format "Complete Name &lt;somebody@yourcompany.com&gt;" or only a simple e-mail address in the format "somebody@yourcompany.com"</param>
    ''' <param name="MsgSubject">Message subject</param>
    ''' <param name="msgTextBody">Message plain text body</param>
    ''' <param name="msgHTMLBody">Message html text body</param>
    ''' <param name="senderName">The name of the sender</param>
    ''' <param name="senderAddress">The e-mail address of the sender</param>
    ''' <param name="msgCharset">The charset of the message, default is UTF-8</param>
    ''' <param name="connection">An open connection to the camm Web-Manager database which shall be used</param>
    ''' <param name="bufErrorDetails">A reference to a string variable which shall contain error information when an exception gets fired</param>
    ''' <param name="attachments">Array of e-mail attachments</param>
    ''' <param name="priority">The priority of the e-mail</param>
    ''' <param name="sensitivity">The sensitivity of the e-mail</param>
    ''' <param name="requestTransmissionConfirmation">Request a transmission confirmation</param>
    ''' <param name="requestReadingConfirmation">Request a reading confirmation</param>
    ''' <param name="additionalHeaders">Additional headers for the e-mail</param>
    ''' <returns>True if successfull, false for failures (also see bufErrorDetails)</returns>
    Private Function SendViaQueue(ByVal rcptAddresses_To As String, ByVal rcptAddresses_CC As String, ByVal rcptAddresses_BCC As String, ByVal msgSubject As String, ByVal msgTextBody As String, ByVal msgHTMLBody As String, ByVal senderName As String, ByVal senderAddress As String, ByVal msgCharset As String, ByVal connection As IDbConnection, ByRef bufErrorDetails As String, ByVal attachments() As EMailAttachment, ByVal priority As EMails.Priority, ByVal sensitivity As EMails.Sensitivity, ByVal requestTransmissionConfirmation As Boolean, ByVal requestReadingConfirmation As Boolean, ByVal additionalHeaders As Collections.Specialized.NameValueCollection) As Boolean
        Try
            FixHtmlContentIDs(msgHTMLBody, attachments)

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
                If priority = EMails.Priority.High OrElse priority = WMSystem.MailImportance.Normal OrElse priority = EMails.Priority.Low Then
                    additionalHeaders("X-MSMail-Priority") = priority.ToString()
                    additionalHeaders("Importance") = priority.ToString()
                End If
            End If

            'Setup mail sensitivity
            If Not sensitivity = Nothing Then
                Select Case sensitivity
                    Case EMails.Sensitivity.Status_CompanyConfidential
                        additionalHeaders("Sensitivity") = "Company-Confidential"
                    Case EMails.Sensitivity.Status_Personal
                        additionalHeaders("Sensitivity") = "Personal"
                    Case EMails.Sensitivity.Status_Private
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
            Dim VersionDataHashtable As New Hashtable
            VersionDataHashtable.Add("NetLibrary", _WebManager.System_Version_Ex.ToString)
            VersionData = Tools.Data.DataTables.ConvertIDictionaryToDataTable(VersionDataHashtable)
            VersionData.TableName = "version"
            MailData.Tables.Add(VersionData)

            'Headers
            Dim Headers As DataTable
            Headers = Tools.Data.DataTables.ConvertNameValueCollectionToDataTable(additionalHeaders)
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
                    Dim MyAttachment As EMailAttachment = Nothing

                    If attachments(MyCounter).FilePath <> "" AndAlso Not System.IO.File.Exists(attachments(MyCounter).FilePath) Then
                        bufErrorDetails = "Attachment file not found: " & attachments(MyCounter).FilePath
                        Return False
                    ElseIf (attachments(MyCounter).RawData Is Nothing OrElse attachments(MyCounter).RawDataFilename = Nothing) AndAlso attachments(MyCounter).FilePath <> "" AndAlso System.IO.File.Exists(attachments(MyCounter).FilePath) Then
                        'Load file system data into memory as raw binary data
                        Dim fs As System.IO.FileStream = New System.IO.FileStream(attachments(MyCounter).FilePath, IO.FileMode.Open)
                        Try
                            Dim fi As IO.FileInfo = New IO.FileInfo(attachments(MyCounter).FilePath)
                            Dim byteArr(CType(fi.Length, Integer) - 1) As Byte
                            fs.Read(byteArr, 0, CType(fi.Length, Integer))
                            attachments(MyCounter).RawData = byteArr
                            attachments(MyCounter).RawDataFilename = IO.Path.GetFileName(attachments(MyCounter).FilePath)
                            attachments(MyCounter).FilePath = Nothing
                        Catch ex As Exception
                            bufErrorDetails = ex.ToString
                            Return False
                        Finally
                            fs.Close()
                        End Try
                    ElseIf Not attachments(MyCounter).RawData Is Nothing AndAlso attachments(MyCounter).RawDataFilename <> Nothing Then
                        'Just add the binary data
                    Else
                        'System.Environment.StackTrace doesn't work with medium-trust --> work around it using a new exception class
                        Dim WorkaroundEx As New Exception("")
                        Dim WorkaroundStackTrace As String = WorkaroundEx.StackTrace 'contains only last few lines of stacktrace
                        Try
                            WorkaroundStackTrace = System.Environment.StackTrace 'contains full stacktrace
                        Catch
                        End Try
                        _WebManager.Log.RuntimeWarning("Empty or invalid email attachment", WorkaroundStackTrace, WMSystem.DebugLevels.NoDebug, True, False)
                        bufErrorDetails = "Empty or invalid email attachment, see camm Web-Manager warnings log for full stack trace"
                        Return False
                    End If

                    'Add the binary data
                    MyAttachment = attachments(MyCounter)
                    Dim MyRow As DataRow = AttachedFiles.NewRow
                    MyRow("FileData") = Utils.ObjectNotNothingOrDBNull(MyAttachment.RawData)
                    MyRow("FileName") = Utils.StringNotNothingOrDBNull(MyAttachment.RawDataFilename)
                    MyRow("Placeholder") = Utils.StringNotNothingOrDBNull(MyAttachment.PlaceholderInMhtmlToBeReplacedByContentID)
                    If MyAttachment.PlaceholderInMhtmlToBeReplacedByContentID <> Nothing Then
                        'Required to drop the origin filename if we want to use a Content-ID
                        MyRow("OriginFileNameBeforePlaceholderValue") = MyRow("FileName")
                        MyRow("FileName") = Utils.StringNotNothingOrDBNull(MyAttachment.PlaceholderInMhtmlToBeReplacedByContentID)
                    End If
                    AttachedFiles.Rows.Add(MyRow)
                Next
            End If
            MailData.Tables.Add(AttachedFiles)

            'Basic message data
            If _WebManager.System_DebugLevel >= WMSystem.DebugLevels.Medium_LoggingAndRedirectAllEMailsToDeveloper Then
                rcptAddresses_To = EMails.CreateReceipientString(_WebManager.DevelopmentEMailAccountAddress, _WebManager.DevelopmentEMailAccountAddress)
                rcptAddresses_CC = Nothing
                rcptAddresses_BCC = Nothing
                If msgTextBody <> "" Then 'text body present
                    msgTextBody = "The server currently runs in debug mode. All e-mails will be sent to your address only." & ControlChars.CrLf & ControlChars.CrLf & ControlChars.CrLf & msgTextBody
                End If
                If msgHTMLBody <> "" Then 'HTML body present
                    msgHTMLBody = "<h4><font face=""Arial"" color=""red""><b><em>The server currently runs in debug mode. All e-mails will be sent to your address only.</em></b></font></h4>" & msgHTMLBody
                End If
            End If
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
            MessageData = Tools.Data.DataTables.ConvertNameValueCollectionToDataTable(MessageDataCollection)
            MessageData.TableName = "message"
            MailData.Tables.Add(MessageData)

            'Add the mail to the queue
            Dim MyConn As IDbConnection
            Dim ReUseExistingConnection As Boolean
            If connection Is Nothing Then
                MyConn = New SqlConnection(_WebManager.ConnectionString)
                ReUseExistingConnection = False
            Else
                MyConn = connection
                ReUseExistingConnection = True
            End If
            Dim MyCmd As New SqlCommand("INSERT INTO [dbo].[Log_eMailMessages]([UserID], [data], [State], [DateTime]) VALUES (@UserID, @Data, @State, GETDATE())" & vbNewLine & "SELECT @ID = @@IDENTITY", CType(MyConn, SqlConnection))
            MyCmd.Parameters.Add("@ID", SqlDbType.Int)
            MyCmd.Parameters("@ID").Direction = ParameterDirection.Output
            If System.Web.HttpContext.Current Is Nothing OrElse System.Web.HttpContext.Current.Session Is Nothing Then
                MyCmd.Parameters.Add("@UserID", SqlDbType.Int).Value = _WebManager.CurrentUserID(WMSystem.SpecialUsers.User_Code)
            Else
                MyCmd.Parameters.Add("@UserID", SqlDbType.Int).Value = _WebManager.CurrentUserID(WMSystem.SpecialUsers.User_Anonymous)
            End If
            MyCmd.Parameters.Add("@Data", SqlDbType.NText).Value = Tools.Data.DataTables.ConvertDatasetToXml(MailData)
            If Me.QueueTransactionData Is Nothing Then
                'Send to current queue immediately
                MyCmd.Parameters.Add("@State", SqlDbType.TinyInt).Value = QueueMonitoring.QueueStates.Queued
            Else
                'Add to queue table, but wait for a commit before real queueing starts
                MyCmd.Parameters.Add("@State", SqlDbType.TinyInt).Value = QueueMonitoring.QueueStates.WaitingForReleaseBeforeQueuing
            End If

            'Execute the command
            If ReUseExistingConnection Then
                'Use open connection, never close it
                Tools.Data.DataQuery.AnyIDataProvider.ExecuteNonQuery(MyCmd, Tools.Data.DataQuery.AnyIDataProvider.Automations.None)
            Else
                'Open and close connection as required
                Tools.Data.DataQuery.AnyIDataProvider.ExecuteNonQuery(MyCmd, Tools.Data.DataQuery.AnyIDataProvider.Automations.AutoOpenAndCloseAndDisposeConnection)
            End If
            'Add ID to collection for transaction handling
            If Not Me.QueueTransactionData Is Nothing Then
                Me.QueueTransactionData.Add(Data.Utils.NoDBNull(MyCmd.Parameters("@ID").Value, 0%))
            End If
            Return True
        Catch ex As Exception
            bufErrorDetails = ex.ToString
            Return False
        End Try

    End Function

End Class
