Option Explicit On
Option Strict On

Public Class SmtpSendResult
    Friend Sub New()
        Me.Success = True
    End Sub
    Friend Sub New(ex As Exception)
        Me.Success = False
        Me.Exception = ex
    End Sub
    Public Property Success As Boolean
    Public Property Exception As Exception
End Class
