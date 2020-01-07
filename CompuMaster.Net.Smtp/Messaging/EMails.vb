Option Explicit On
Option Strict On

Public Class EMails

    Public Enum Priority As Integer
        High = 1
        Normal = 3
        Low = 5
    End Enum

    Public Enum Sensitivity As Integer
        Status_Normal = 1
        Status_Personal = 2
        Status_Private = 3
        Status_CompanyConfidential = 4
    End Enum

End Class
