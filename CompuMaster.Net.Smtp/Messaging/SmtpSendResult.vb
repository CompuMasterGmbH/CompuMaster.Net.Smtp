﻿Option Explicit On
Option Strict On

Public Class SmtpSendResult
    Public Sub New()
        Me.Success = True
    End Sub
    Public Sub New(ex As Exception)
        If ex Is Nothing Then Throw New ArgumentNullException(NameOf(ex))
        Me.Success = False
        Me.Exception = ex
    End Sub
    Public Property Success As Boolean
    Public Property Exception As Exception
End Class
