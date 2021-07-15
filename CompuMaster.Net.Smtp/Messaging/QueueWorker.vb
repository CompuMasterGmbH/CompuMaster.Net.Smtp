Option Explicit On
Option Strict On

Imports System.Data.SqlClient
Imports CompuMaster.Data

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
    Public Sub CommitQueueTransaction(connection As IDbConnection)

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
    Public Sub RollbackQueueTransaction(connection As IDbConnection)

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
    Private Function SendViaQueue(RcptName As String, RcptAddress As String, MsgSubject As String, MsgTextBody As String, MsgHTMLBody As String, SenderName As String, SenderAddress As String, MsgCharset As String, connection As IDbConnection, ByRef bufErrorDetails As String, Attachments() As EMailAttachment, Priority As EMails.Priority, Sensitivity As EMails.Sensitivity, RequestTransmissionConfirmation As Boolean, RequestReadingConfirmation As Boolean, AdditionalHeaders As Collections.Specialized.NameValueCollection) As Boolean
        Return SendViaQueue(CreateReceipientString(RcptName, RcptAddress), CType(Nothing, String), CType(Nothing, String), MsgSubject, MsgTextBody, MsgHTMLBody, SenderName, SenderAddress, MsgCharset, connection, bufErrorDetails, Attachments, Priority, Sensitivity, RequestTransmissionConfirmation, RequestReadingConfirmation, AdditionalHeaders)
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
    Private Function SendViaQueue(rcptAddresses_To As String, rcptAddresses_CC As String, rcptAddresses_BCC As String, msgSubject As String, msgTextBody As String, msgHTMLBody As String, senderName As String, senderAddress As String, msgCharset As String, connection As IDbConnection, ByRef bufErrorDetails As String, attachments() As EMailAttachment, priority As EMails.Priority, sensitivity As EMails.Sensitivity, requestTransmissionConfirmation As Boolean, requestReadingConfirmation As Boolean, additionalHeaders As Collections.Specialized.NameValueCollection) As Boolean
        Try
            Dim MailData As String = Data.DataTables.ConvertDatasetToXml(emailmessage.CreateQueueMessageXml(
                rcptAddresses_To,
                rcptAddresses_CC,
                rcptAddresses_BCC,
                msgSubject,
                msgTextBody,
                msgHTMLBody,
                senderName,
                senderAddress,
                msgCharset,
                attachments(),
                priority,
                sensitivity,
                requestTransmissionConfirmation,
                requestReadingConfirmation,
                additionalHeaders
                ))

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
            MyCmd.Parameters.Add("@Data", SqlDbType.NText).Value = MailData
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
                DataQuery.AnyIDataProvider.ExecuteNonQuery(MyCmd, DataQuery.AnyIDataProvider.Automations.None)
            Else
                'Open and close connection as required
                DataQuery.AnyIDataProvider.ExecuteNonQuery(MyCmd, DataQuery.AnyIDataProvider.Automations.AutoOpenAndCloseAndDisposeConnection)
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
