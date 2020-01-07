Option Explicit On
Option Strict On

Namespace CompuMaster.camm.WebManager

    ''' <summary>
    ''' A database layer for MS SQL Server
    ''' </summary>
    Friend Class DataLayerSqlClient
        Implements IDataLayer

        Public Sub SaveLastServiceExecutionDate(webmanager As IWebManager, dbConnection As IDbConnection, triggerServiceVersion As String) Implements IDataLayer.SaveLastServiceExecutionDate
            Dim MyConn As SqlClient.SqlConnection = Nothing
            Try
                'Prepare the connection for the several SQL commands
                If dbConnection Is Nothing Then
                    'only when we created our connection ourself
                    MyConn = New SqlClient.SqlConnection(webmanager.ConnectionString)
                    MyConn.Open()
                Else
                    'use given connection
                    MyConn = CType(dbConnection, SqlClient.SqlConnection)
                End If

                Dim sql As String = "UPDATE [dbo].[System_GlobalProperties] SET ValueDateTime=GetDate(), ValueNText = @VersionString WHERE PropertyName = 'LastWebServiceExecutionDate' AND ValueNVarChar = N'camm WebManager' " & vbNewLine &
                    "IF @@ROWCOUNT = 0 " & vbNewLine &
                    "INSERT INTO [dbo].[System_GlobalProperties] (PropertyName, ValueNVarChar, ValueNText, ValueDateTime) VALUES ('LastWebServiceExecutionDate', 'camm WebManager', @VersionString, GetDate()) "
                Dim cmd As New System.Data.SqlClient.SqlCommand(sql, MyConn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@VersionString", SqlDbType.NText).Value = Utils.StringNotEmptyOrAlternativeValue(triggerServiceVersion, "0.0.0.0")
                Tools.Data.DataQuery.AnyIDataProvider.ExecuteNonQuery(cmd, Tools.Data.DataQuery.AnyIDataProvider.Automations.AutoOpenAndCloseAndDisposeConnection)
            Finally
                'Close connection if created/opened by above code
                If dbConnection Is Nothing Then
                    'executes only when we created our connection by ourself
                    Tools.Data.DataQuery.AnyIDataProvider.CloseAndDisposeConnection(MyConn)
                End If
            End Try
        End Sub

        Public Function QueryLastServiceExecutionDate(webmanager As IWebManager, dbConnection As IDbConnection) As Date Implements IDataLayer.QueryLastServiceExecutionDate
            Dim MyConn As SqlClient.SqlConnection = Nothing
            Try
                'Prepare the connection for the several SQL commands
                If dbConnection Is Nothing Then
                    'only when we created our connection ourself
                    MyConn = New SqlClient.SqlConnection(webmanager.ConnectionString)
                    MyConn.Open()
                Else
                    'use given connection
                    MyConn = CType(dbConnection, SqlClient.SqlConnection)
                End If

                Dim cmd As System.Data.SqlClient.SqlCommand
                cmd = New SqlClient.SqlCommand("SELECT ValueDateTime FROM [dbo].System_GlobalProperties WHERE PropertyName = 'LastWebServiceExecutionDate' AND ValueNVarChar = N'camm WebManager'", MyConn)
                Dim result As Object = Tools.Data.DataQuery.AnyIDataProvider.ExecuteScalar(cmd, Tools.Data.DataQuery.AnyIDataProvider.Automations.AutoOpenAndCloseAndDisposeConnection)
                If result Is Nothing OrElse Data.Utils.NoDBNull(result) Is Nothing Then
                    Return Nothing
                Else
                    Return CType(result, Date)
                End If
            Finally
                'Close connection if created/opened by above code
                If dbConnection Is Nothing Then
                    'executes only when we created our connection by ourself
                    Tools.Data.DataQuery.AnyIDataProvider.CloseAndDisposeConnection(MyConn)
                End If
            End Try
        End Function

    End Class

End Namespace