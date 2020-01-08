Option Explicit On
Option Strict On

Imports System.Data.SqlClient

''' <summary>
'''     camm Web-Manager base class
''' </summary>
''' <remarks>
'''     <p>This is the base class of camm Web-Manager. Instantiate this one in your ASP.NET pages by using ~/system/cammWebManager.ascx</p>
''' </remarks>
Public Class WMSystem

    ''' <summary>
    ''' Debug level of camm Web-Manager
    ''' </summary>
    ''' <value>The new debug level</value>
    ''' <remarks>The logging of warnings, errors can be influenced by this property as well as redirection of e-mails and the functionality level of automatic redirects. This is very helpfull on first setup of camm Web-Manager if you experience any problems or you don't want e-mails to be sent to the origin receipients except you yourself.</remarks>
    Public Property DebugLevel() As DebugLevels
        Get
            If _System_DebugLevel = Nothing Then
                _System_DebugLevel = Configuration.DebugLevel
            End If
            Return _System_DebugLevel
        End Get
        Set(Value As DebugLevels)
            _System_DebugLevel = Value
        End Set
    End Property

    Private _System_DebugLevel As DebugLevels = DebugLevels.NoDebug

    ''' <summary>
    ''' Existing debug levels in camm Web-Manager
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum DebugLevels As Byte
        ''' <summary>
        '''     No debugging
        ''' </summary>
        ''' <remarks>
        '''     This feature is set to disabled.
        ''' </remarks>
        NoDebug = 0 'kein Debug
        ''' <summary>
        '''     Access error warnings only
        ''' </summary>
        ''' <remarks>
        '''     Warning messages will be sent to the developer contact configured in your config files.
        ''' </remarks>
        Low_WarningMessagesOnAccessError = 1 'Warn-Messages über Access Errors an Developer Contact
        ''' <summary>
        '''     More access error warnings
        ''' </summary>
        ''' <remarks>
        '''     Additional warning messages will be sent to the developer contact configured in your config files.
        ''' </remarks>
        Low_WarningMessagesOnAccessError_AdditionalDetails = 2 'Protokollierung von zusätzlichen Informationen
        ''' <summary>
        '''     Actively collect data for debugging
        ''' </summary>
        ''' <remarks>
        '''     Even more additional warning messages will be sent to the developer contact configured in your config files.
        ''' </remarks>
        Medium_LoggingOfDebugInformation = 3 'Protokollierung von zusätzlichen Debug-Informationen
        ''' <summary>
        '''     Send all e-mails to developer account - never use in production environments!
        ''' </summary>
        ''' <remarks>
        '''     <para>ATTENTION: Never use this mode in production environments!</para>
        '''     <para>All e-mails will be redirected to the developer; no e-mail will be sent to the origin receipient. This is optimal for project development and testing environments.</para>
        '''		<para>The messages will be sent to the developer contact configured in your config files.</para>
        ''' </remarks>
        <Obsolete("Use Medium_LoggingAndRedirectAllEMailsToDeveloper instead"), System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)> Medium_LoggingOfDebugInformation_AdditionalDetails = 4 'e-mail-Umleitung an Developer Contact
        ''' <summary>
        '''     Send all e-mails to developer account - never use in production environments!
        ''' </summary>
        ''' <remarks>
        '''     <para>ATTENTION: Never use this mode in production environments!</para>
        '''     <para>All e-mails will be redirected to the developer; no e-mail will be sent to the origin receipient. This is optimal for project development and testing environments.</para>
        '''		<para>The messages will be sent to the developer contact configured in your config files.</para>
        ''' </remarks>
        Medium_LoggingAndRedirectAllEMailsToDeveloper = 4 'e-mail-Umleitung an Developer Contact
        ''' <summary>
        '''     Maximum debug level - never use in production environments!
        ''' </summary>
        ''' <remarks>
        '''     <para>ATTENTION: Never use this mode in production environments!</para>
        '''     <para>All e-mails will be redirected to the developer; no e-mail will be sent to the origin receipient. This is optimal for setting up the project.</para>
        '''		<para>The messages will be sent to the developer contact configured in your config files.</para>
        '''     <para>Automatic redirects have to be manually executed. This is ideally for solving redirection bugs or when loosing session paths in cookieless scenarios.</para>
        ''' </remarks>
        Max_RedirectPageRequestsManually = 5 'e-mail-Umleitung an Developer Contact + manual page redirections
        <ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never), Obsolete("Use Max_RedirectPageRequestsManually instead")> Max_RedirectAllEMailsToDeveloper = 5 'e-mail-Umleitung an Developer Contact + manual page redirections
    End Enum

    'Friend ReadOnly Property GlobalConfiguration As GlobalConfiguration
    '    Get
    '        Static _Result As GlobalConfiguration
    '        If _Result Is Nothing Then _Result = New GlobalConfiguration(Me)
    '        Return _Result
    '    End Get
    'End Property

    ''' <summary>
    '''     Represents a state with 3 possible values: True, False and Undefined (default)
    ''' </summary>
    Friend Enum TripleState As Byte
        Undefined = 0
        [True] = 1
        [False] = 2
    End Enum

    '''' <summary>
    ''''     The messaging methods for e-mail distribution
    '''' </summary>
    '''' <value></value>
    'Public ReadOnly Property MessagingEMails() As EMails
    '    Get
    '        Static _MessagingEMails As EMails
    '        If _MessagingEMails Is Nothing Then
    '            _MessagingEMails = New EMails(Me)
    '        End If
    '        Return _MessagingEMails
    '    End Get
    'End Property
    '''' <summary>
    ''''     The queue monitoring methods for message distribution
    '''' </summary>
    '''' <value></value>
    'Public ReadOnly Property MessagingQueueMonitoring() As QueueMonitoring
    '    Get
    '        Static _MessagingQueueMonitoring As QueueMonitoring
    '        If _MessagingQueueMonitoring Is Nothing Then
    '            _MessagingQueueMonitoring = New QueueMonitoring(Me)
    '        End If
    '        Return _MessagingQueueMonitoring
    '    End Get
    'End Property

End Class
