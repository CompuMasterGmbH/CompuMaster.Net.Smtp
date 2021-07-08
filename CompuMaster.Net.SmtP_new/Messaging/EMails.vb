Option Explicit On
Option Strict On

Public NotInheritable Class EMails

    Public Enum Priority As Integer
        High = 1
        ''' <summary>
        ''' Not supported by all e-mail systems
        ''' </summary>
        Higher = 2
        Normal = 3
        ''' <summary>
        ''' Not supported by all e-mail systems
        ''' </summary>
        Lower = 4
        Low = 5
    End Enum

    Public Enum Sensitivity As Integer
        Normal = 1
        Personal = 2
        [Private] = 3
        CompanyConfidential = 4
    End Enum

End Class
