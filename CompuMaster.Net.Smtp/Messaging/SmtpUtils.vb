Option Explicit On
Option Strict On

Friend Class SmtpUtils

    ''' <summary>
    ''' Check the expression and return a strongly typed value
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="expression">The expression for the check</param>
    ''' <param name="trueValue">The return value if the expression is true</param>
    ''' <param name="falseValue">The return value if the expression is false</param>
    ''' <returns>A strongly typed return value</returns>
    Public Shared Function IIf(Of T)(expression As Boolean, trueValue As T, falseValue As T) As T
        If expression Then Return trueValue Else Return falseValue
    End Function

    ''' <summary>
    '''     Split a string by a separator if there is not a special leading character
    ''' </summary>
    ''' <param name="text"></param>
    ''' <param name="separator"></param>
    ''' <param name="exceptLeadingCharacter"></param>
    Public Shared Function SplitString(ByVal text As String, ByVal separator As Char, ByVal exceptLeadingCharacter As Char) As String()
        If text = Nothing Then
            Return New String() {}
        End If
        Dim Result As New List(Of String)
        'Go through every char of the string
        Dim SplitHere As Boolean
        Dim StartPosition As Integer
        For MyCounter As Integer = 0 To text.Length - 1
            'Find split points
            If text.Chars(MyCounter) = separator Then
                If MyCounter = 0 Then
                    SplitHere = True
                ElseIf text.Chars(MyCounter - 1) <> exceptLeadingCharacter Then
                    SplitHere = True
                End If
            End If
            'Add partial string
            If SplitHere OrElse MyCounter = text.Length - 1 Then
                Result.Add(text.Substring(StartPosition, IIf(Of Integer)(SplitHere = False, 1, 0) + MyCounter - StartPosition)) 'If Split=False then this if-block was caused by the end of the text; in this case we have to simulate to be after the last character position to ensure correct extraction of the last text element
                SplitHere = False 'Reset status
                StartPosition = MyCounter + 1 'Next string starts after the current char
            End If
        Next
        Return Result.ToArray
    End Function

    ''' <summary>
    ''' String comparison types for ReplaceString method
    ''' </summary>
    ''' <remarks></remarks>
    Friend Enum ReplaceComparisonTypes As Byte
        ''' <summary>
        ''' Compare 2 strings with case sensitivity
        ''' </summary>
        ''' <remarks></remarks>
        CaseSensitive = 0
        ''' <summary>
        ''' Compare 2 strings by lowering their case based on the current culture
        ''' </summary>
        ''' <remarks></remarks>
        CurrentCultureIgnoreCase = 1
        ''' <summary>
        ''' Compare 2 strings by lowering their case following invariant culture rules
        ''' </summary>
        ''' <remarks></remarks>
        InvariantCultureIgnoreCase = 2
    End Enum

    ''' <summary>
    ''' Replace a string in another string based on a defined StringComparison type
    ''' </summary>
    ''' <param name="original">The original string</param>
    ''' <param name="pattern">The search expression</param>
    ''' <param name="replacement">The string which shall be inserted instead of the pattern</param>
    ''' <param name="comparisonType">The comparison type for searching for the pattern</param>
    ''' <returns>A new string with all replacements</returns>
    ''' <remarks></remarks>
    Friend Shared Function ReplaceString(ByVal original As String, ByVal pattern As String, ByVal replacement As String, ByVal comparisonType As ReplaceComparisonTypes) As String
        If original = Nothing OrElse pattern = Nothing Then
            Return original
        End If
        Dim lenPattern As Integer = pattern.Length
        Dim idxPattern As Integer = -1
        Dim idxLast As Integer = 0
        Dim result As New System.Text.StringBuilder
        Select Case comparisonType
            Case ReplaceComparisonTypes.CaseSensitive
                While True
                    idxPattern = original.IndexOf(pattern, idxPattern + 1, comparisonType)
                    If idxPattern < 0 Then
                        result.Append(original, idxLast, original.Length - idxLast)
                        Exit While
                    End If
                    result.Append(original, idxLast, idxPattern - idxLast)
                    result.Append(replacement)
                    idxLast = idxPattern + lenPattern
                End While
            Case ReplaceComparisonTypes.CurrentCultureIgnoreCase
                While True
                    Dim comparisonStringOriginal As String, comparisonStringPattern As String
                    comparisonStringOriginal = original.ToLower(System.Globalization.CultureInfo.CurrentCulture)
                    comparisonStringPattern = pattern.ToLower(System.Globalization.CultureInfo.CurrentCulture)
                    idxPattern = comparisonStringOriginal.IndexOf(comparisonStringPattern, idxPattern + 1)
                    If idxPattern < 0 Then
                        result.Append(original, idxLast, original.Length - idxLast)
                        Exit While
                    End If
                    result.Append(original, idxLast, idxPattern - idxLast)
                    result.Append(replacement)
                    idxLast = idxPattern + lenPattern
                End While
            Case ReplaceComparisonTypes.InvariantCultureIgnoreCase
                While True
                    Dim comparisonStringOriginal As String, comparisonStringPattern As String
                    comparisonStringOriginal = original.ToLower(System.Globalization.CultureInfo.CurrentCulture)
                    comparisonStringPattern = pattern.ToLower(System.Globalization.CultureInfo.CurrentCulture)
                    idxPattern = comparisonStringOriginal.IndexOf(comparisonStringPattern, idxPattern + 1)
                    If idxPattern < 0 Then
                        result.Append(original, idxLast, original.Length - idxLast)
                        Exit While
                    End If
                    result.Append(original, idxLast, idxPattern - idxLast)
                    result.Append(replacement)
                    idxLast = idxPattern + lenPattern
                End While
            Case Else
                Throw New ArgumentOutOfRangeException(NameOf(comparisonType), "Invalid value")
        End Select
        Return result.ToString()
    End Function

End Class
