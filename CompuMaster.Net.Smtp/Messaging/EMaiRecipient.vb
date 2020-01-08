Option Strict On
Option Explicit On

Public Class EMailRecipient

    ''' <summary>
    '''     Create a valid receipient string for the address lists parameters of method SendEMail
    ''' </summary>
    ''' <param name="name">The name of the receipient</param>
    ''' <param name="address">The e-mail address of the receipient</param>
    ''' <remarks>
    '''     <para>In case of an empty address, an empty string will be returned, too (no ArgumentNullException).</para>
    '''     <para>RFC-821 standard describes as following:</para>
    '''	    <para>&lt;special&gt; ::= "&lt;" | "&gt;" | "(" | ")" | "[" | "]" | "\" | "."
    '''	              | "," | ";" | ":" | "@"  """ | the control
    '''	              characters (ASCII codes 0 through 31 inclusive And _
    '''	              127)</para>
    ''' </remarks>
    Public Shared Function CreateRecipientString(ByVal name As String, ByVal address As String) As String
        If address = Nothing Then Return ""
        If name <> Nothing Then
            name = name.Replace("\", "\\")
            name = name.Replace("<", "\<").Replace(">", "\>").Replace("(", "\(").Replace(")", "\)").Replace("[", "\[").Replace("]", "\]")
            name = name.Replace(".", "\.").Replace(",", "\,").Replace(":", "\:").Replace(";", "\;")
            name = name.Replace("@", "\@")
            name = name.Replace("""", "'") 'quotations marks lead to bugs in Chilkat (or at least there)
            'Replace control characters
            For MyCounter As Integer = 0 To 31
                name = name.Replace(ChrW(MyCounter), " "c)
            Next
            name = name.Replace(ChrW(127), " "c)
        End If
        Return CType(IIf(name <> Nothing, name & " ", ""), String) & "<" & address & ">"
    End Function

    ''' <summary>
    '''     Create a valid receipient string for the address lists parameters of method SendEMail
    ''' </summary>
    ''' <param name="address">The e-mail address of the receipient</param>
    ''' <remarks>
    '''     <para>In case of an empty address, an empty string will be returned, too (no ArgumentNullException).</para>
    ''' </remarks>
    Public Shared Function CreateRecipientString(ByVal address As String) As String
        Return CreateRecipientString(Nothing, address)
    End Function

    ''' <summary>
    '''     Extract e-mail addresses from a receipients list constructed by method CreateReceipientString or manually
    ''' </summary>
    ''' <param name="receipientsCommaSeparated"></param>
    ''' <remarks>
    '''     Valid parameter values are e. g. 
    ''' <list>
    ''' <item>l?nc\,\.\:\;\@\[\]\(\)ma9\'\\\&lt;\&gt;kljkj &lt;compumaster@web.de&gt;, her\,ma9\'\\\&lt;\&gt;kljkj &lt;jwezel@compumaster.de&gt;</item>
    ''' <item>&lt;compumaster@web.de&gt;,&lt;jwezel@compumaster.de&gt;</item>
    ''' <item>compumaster@web.de,jwezel@compumaster.de</item>
    ''' </list>
    ''' </remarks>
    Friend Shared Function SplitEMailAddressesFromRecipientsList(ByVal receipientsCommaSeparated As String) As String()
        Dim Result As String() = SmtpUtils.SplitString(receipientsCommaSeparated, ","c, "\"c)
        For MyCounter As Integer = 0 To Result.Length - 1
            'Now we got strings like
            '"l?nc\,\.\:\;\@\[\]\(\)ma9\'\\\<\>kljkj <compumaster@web.de>"
            '" her\,ma9\'\\\<\>kljkj <jwezel@compumaster.de>"
            '" ""Jochen"" <jwezel@compumaster.de>" <------------TOCHECK!
            '" <compumaster@web.de>"
            '"<jwezel@compumaster.de>" 
            '"compumaster@web.de"
            '"jwezel@compumaster.de"
            Dim AddressStart As Integer = Result(MyCounter).LastIndexOf("<")
            If AddressStart > 0 Then
                Result(MyCounter) = Mid(Result(MyCounter), AddressStart)
            End If
            Result(MyCounter) = Trim(Result(MyCounter))
            If InStr(Result(MyCounter), "@") < 2 OrElse InStr(Result(MyCounter), "@") > Result(MyCounter).Length - 1 Then
                Result(MyCounter) = ""
            End If
            'Now we got strings like
            '"<compumaster@web.de>"
            '"<jwezel@compumaster.de>"
            '"<compumaster@web.de>"
            '"<jwezel@compumaster.de>" 
            '"compumaster@web.de"
            '"jwezel@compumaster.de"
            If Left(Result(MyCounter), 1) = "<" AndAlso Right(Result(MyCounter), 1) = ">" Then
                Result(MyCounter) = Mid(Result(MyCounter), 2, Len(Result(MyCounter)) - 2)
            End If
        Next
        Return Result
    End Function

    ''' <summary>
    ''' Split an e-mail address in typical encoding for SMTP protocol into the two parts e-mail address and receipient name
    ''' </summary>
    ''' <param name="emailAddressInSmtpFormat"></param>
    ''' <remarks></remarks>
    Friend Shared Function SplitEMailAddressesIntoAddressParts(ByVal emailAddressInSmtpFormat As String) As EMailRecipient
        Dim Receipient As New EMailRecipient

        If Not emailAddressInSmtpFormat.LastIndexOf("<") = -1 Then
            Receipient.Address = emailAddressInSmtpFormat.Substring(emailAddressInSmtpFormat.LastIndexOf("<"), emailAddressInSmtpFormat.Length - emailAddressInSmtpFormat.LastIndexOf("<"))
            Receipient.Address = Receipient.Address.Replace("<", "").Replace(">", "")
        Else
            Receipient.Address = emailAddressInSmtpFormat
        End If

        If emailAddressInSmtpFormat.LastIndexOf(" ") > 0 AndAlso emailAddressInSmtpFormat.LastIndexOf(" ") = emailAddressInSmtpFormat.LastIndexOf("<") - 1 Then
            Receipient.Name = emailAddressInSmtpFormat.Substring(0, emailAddressInSmtpFormat.LastIndexOf(" "))
            Receipient.Name = Receipient.Name.Replace("<", "").Replace(">", "")
        Else
            Receipient.Name = String.Empty
        End If

        Return Receipient
    End Function

    ''' <summary>
    ''' Split an e-mail address list in typical encoding for SMTP protocol into the two parts e-mail address and receipient name
    ''' </summary>
    ''' <param name="emailAddressesInSmtpFormat"></param>
    ''' <remarks></remarks>
    Friend Shared Function SplitEMailAddressesIntoEMailRecipientsFromRecipientsList(ByVal emailAddressesInSmtpFormat As String) As EMailRecipient()
        Dim Result As String() = SmtpUtils.SplitString(emailAddressesInSmtpFormat, ","c, "\"c)
        Dim EmailReceipient(Result.Length - 1) As EMailRecipient
        For MyCounter As Integer = 0 To Result.Length - 1
            EmailReceipient(MyCounter) = SplitEMailAddressesIntoAddressParts(Result(MyCounter))
        Next
        Return EmailReceipient
    End Function

    Public Sub New()
    End Sub

    Public Sub New(ByVal name As String, ByVal address As String)
        If address Is Nothing Then
            Throw New ArgumentNullException(NameOf(address))
        End If
        Me.Name = name
        Me.Address = address
    End Sub

    Private _Name As String
    Private _Address As String

    Public Property Address() As String
        Get
            Return _Address
        End Get
        Set(ByVal value As String)
            If value Is Nothing Then
                Throw New ArgumentNullException("Address")
            End If
            _Address = value
        End Set
    End Property
    Public Property Name() As String
        Get
            Return _Name
        End Get
        Set(ByVal value As String)
            _Name = value
        End Set
    End Property
End Class