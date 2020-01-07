''' <summary>
'''     Configuration settings read from web.config/app.config
''' </summary>
Friend Class Configuration

    ''' <summary>
    ''' Debug level of camm Web-Manager
    ''' </summary>
    ''' <value>The new debug level</value>
    ''' <remarks>The logging of warnings, errors can be influenced by this property as well as redirection of e-mails and the functionality level of automatic redirects. This is very helpfull on first setup of camm Web-Manager if you experience any problems or you don't want e-mails to be sent to the origin receipients except you yourself.</remarks>
    Public Shared ReadOnly Property DebugLevel() As WMSystem.DebugLevels
        Get
            Try
                Return CType(WebManagerSettings.Item("DebugLevel"), WMSystem.DebugLevels)
            Catch
                Return WMSystem.DebugLevels.NoDebug
            End Try
        End Get
    End Property

    ''' <summary>
    ''' The SMTP server name for sending e-mails
    ''' </summary>
    ''' <value></value>
    ''' <remarks>The camm Web-Manager is heavily using e-mails for information and workflow purposes.</remarks>
    Public Shared ReadOnly Property SmtpServerName() As String
        Get
            Dim Result As String = Nothing
            Try
                Return WebManagerSettings.Item("SMTPServerName")
            Catch
            End Try
            If Result <> Nothing Then
                Return Result
            Else
                Return "localhost"
            End If
        End Get
    End Property
    ''' <summary>
    ''' The SMTP server port for sending e-mails
    ''' </summary>
    ''' <value></value>
    ''' <remarks>The camm Web-Manager is heavily using e-mails for information and workflow purposes.</remarks>
    Public Shared ReadOnly Property SmtpServerPort() As Integer
        Get
            Dim Result As Integer = 25
            Try
                Dim SettingValue As String = WebManagerSettings.Item("SMTPServerPort")
                If SettingValue Is Nothing OrElse SettingValue = "" Then
                    'Default Port
                Else
                    Result = CType(SettingValue, Integer)
                End If
            Catch
            End Try
            Return Result
        End Get
    End Property

    ''' <summary>
    '''     Configures the authentication methods for the SMTP server
    ''' </summary>
    ''' <value>The authentification type, possible values are "" or "NONE", "LOGIN", "PLAIN" or "NTLM"</value>
    Public Shared ReadOnly Property SmtpAuthentificationMode() As String
        Get
            Try
                Dim Result As String = WebManagerSettings.Item("SMTPServerAuthentificationMode")
                If Result Is Nothing Then
                    Result = ""
                Else
                    Result = Result.ToUpper(System.Globalization.CultureInfo.InvariantCulture)
                End If
                Select Case Result
                    Case "", "NONE", "LOGIN", "PLAIN", "NTLM"
                    Case Else
                        Throw New ArgumentException("The authentification type, possible values are """" or ""NONE"", ""LOGIN"", ""PLAIN"" or ""NTLM""")
                End Select
                Return Result
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    ''' <summary>
    '''     Sets up an user name for the SMTP server
    ''' </summary>
    ''' <value>The user name</value>
    Public Shared ReadOnly Property SmtpUsername() As String
        Get
            Try
                Return WebManagerSettings.Item("SMTPUsername")
            Catch
                Try
                    Return WebManagerSettings.Item("SMTPServerUsername")
                Catch
                    Return Nothing
                End Try
            End Try
        End Get
    End Property

    ''' <summary>
    '''     Sets up a password for the SMTP server
    ''' </summary>
    ''' <value>The password</value>
    Public Shared ReadOnly Property SmtpPassword() As String
        Get
            Try
                Return WebManagerSettings.Item("SMTPPassword")
            Catch
                Try
                    Return WebManagerSettings.Item("SMTPServerPassword")
                Catch
                    Return Nothing
                End Try
            End Try
        End Get
    End Property

    ''' <summary>
    '''     Use the integrated mail queue for sending e-mails
    ''' </summary>
    ''' <value>TripleState.Undefined is for auto-detection, use TripleState.False for direct sending without mail queue or TripleState.True for sending via mail queue</value>
    Public Shared ReadOnly Property ProcessMailQueue() As WMSystem.TripleState
        Get
            Return LoadTripleStateSetting("ProcessMailQueue", True)
        End Get
    End Property

#Region "Load configuration setting helper methods"

    Friend Shared ReadOnly Property WebManagerSettings() As System.Collections.Specialized.NameValueCollection
        Get
            Return System.Configuration.ConfigurationManager.AppSettings
        End Get
    End Property

    ''' <summary>
    '''     Load a boolean value from the configuration file
    ''' </summary>
    ''' <param name="appSettingName">The name of the appSetting item</param>
    ''' <param name="suppressExceptions">True if exceptions shall be suppressed and default value returned or False if exception shall be thrown if there is an error</param>
    Private Shared Function LoadTripleStateSetting(ByVal appSettingName As String, ByVal suppressExceptions As Boolean) As WMSystem.TripleState
        Dim Result As WMSystem.TripleState = WMSystem.TripleState.Undefined
        Try
            Dim value As String = CType(WebManagerSettings.Item(appSettingName), String)
            If Not value Is Nothing Then
                value = value.ToLower(System.Globalization.CultureInfo.InvariantCulture)
            Else
                value = ""
            End If
            Select Case value
                Case "0", "off", "false"
                    Result = WMSystem.TripleState.False
                Case "1", "-1", "on", "true"
                    Result = WMSystem.TripleState.True
                Case "", "auto"
                    'keep default
                Case Else
                    Result = Utils.BooleanToWMTriplestate(Boolean.Parse(value))
            End Select
        Catch ex As Exception
            If suppressExceptions = False Then
                Throw New ArgumentException("Configuration setting for """ & appSettingName & """ is invalid. Choose ""off"" or ""on"" or ""auto"", please.", ex)
            End If
        End Try
        Return Result
    End Function

    ''' <summary>
    '''     Load a boolean value from the configuration file
    ''' </summary>
    ''' <param name="appSettingName">The name of the appSetting item</param>
    ''' <param name="defaultValue">A default value if not configured of configured invalid</param>
    ''' <param name="suppressExceptions">True if exceptions shall be suppressed and default value returned or False if exception shall be thrown if there is an error</param>
    Private Shared Function LoadBooleanSetting(ByVal appSettingName As String, ByVal defaultValue As Boolean, ByVal suppressExceptions As Boolean) As Boolean
        Dim Result As Boolean = defaultValue
        Try
            Dim value As String = CType(WebManagerSettings.Item(appSettingName), String)
            If Not value Is Nothing Then
                value = value.ToLower(System.Globalization.CultureInfo.InvariantCulture)
            Else
                value = ""
            End If
            Select Case value
                Case "0", "off", "false"
                    Result = False
                Case "1", "-1", "on", "true"
                    Result = True
                Case ""
                    'keep default
                Case Else
                    Result = Boolean.Parse(value)
            End Select
        Catch ex As Exception
            If suppressExceptions = False Then
                Throw New ArgumentException("Configuration setting for """ & appSettingName & """ is invalid. Choose ""true"" or ""false"", please.", ex)
            End If
        End Try
        Return Result
    End Function

    ''' <summary>
    '''     Load an integer value from the configuration file
    ''' </summary>
    ''' <param name="appSettingName">The name of the appSetting item</param>
    ''' <param name="defaultValue">A default value if not configured of configured invalid</param>
    ''' <param name="suppressExceptions">True if exceptions shall be suppressed and default value returned or False if exception shall be thrown if there is an error</param>
    Friend Shared Function LoadIntegerSetting(ByVal appSettingName As String, ByVal defaultValue As Integer, ByVal suppressExceptions As Boolean) As Integer
        Dim Result As Integer = defaultValue
        Try
            Dim value As String = CType(WebManagerSettings.Item(appSettingName), String)
            If value = Nothing Then
                Result = defaultValue
            Else
                Result = Integer.Parse(value, System.Globalization.CultureInfo.InvariantCulture)
            End If
        Catch ex As Exception
            If suppressExceptions = False Then
                Throw New ArgumentException("Configuration setting for """ & appSettingName & """ is invalid. Choose a valid integer number, please.", ex)
            End If
        End Try
        Return Result
    End Function

    ''' <summary>
    '''     Load an Int64 value from the configuration file
    ''' </summary>
    ''' <param name="appSettingName">The name of the appSetting item</param>
    ''' <param name="defaultValue">A default value if not configured of configured invalid</param>
    ''' <param name="suppressExceptions">True if exceptions shall be suppressed and default value returned or False if exception shall be thrown if there is an error</param>
    Friend Shared Function LoadLongSetting(ByVal appSettingName As String, ByVal defaultValue As Long, ByVal suppressExceptions As Boolean) As Long
        Dim Result As Long = defaultValue
        Try
            Dim value As String = CType(WebManagerSettings.Item(appSettingName), String)
            If value = Nothing Then
                Result = defaultValue
            Else
                Result = Integer.Parse(value, System.Globalization.CultureInfo.InvariantCulture)
            End If
        Catch ex As Exception
            If suppressExceptions = False Then
                Throw New ArgumentException("Configuration setting for """ & appSettingName & """ is invalid. Choose a valid integer number, please.", ex)
            End If
        End Try
        Return Result
    End Function

    ''' <summary>
    '''     Load a string value from the configuration file
    ''' </summary>
    ''' <param name="appSettingName">The name of the appSetting item</param>
    ''' <param name="defaultValue">A default value if not configured of configured invalid</param>
    Friend Shared Function LoadStringSetting(ByVal appSettingName As String, ByVal defaultValue As String) As String
        Dim value As String = CType(WebManagerSettings.Item(appSettingName), String)
        If value = Nothing Then
            Return defaultValue
        Else
            Return value
        End If
    End Function
#End Region

End Class
