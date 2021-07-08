Option Strict On
Option Explicit On

Public Class EMailAddress

    Public Sub New(address As System.Net.Mail.MailAddress)
        Me.New(address.DisplayName, address.Address)
    End Sub

    Public Sub New(address As String)
        Me.New(CType(Nothing, String), address)
    End Sub

    Public Sub New(name As String, address As String)
        If address Is Nothing Then
            Throw New ArgumentNullException(NameOf(address))
        End If
        Me.Name = name
        Me.Address = address
    End Sub

    Private _Address As String
    Public Property Address() As String
        Get
            Return _Address
        End Get
        Set(value As String)
            If value Is Nothing Then
                Throw New ArgumentNullException("Address")
            End If
            _Address = value
        End Set
    End Property

    Public Property Name() As String

    ''' <summary>
    ''' Convert EMailRecipient into address format of SMTP standard
    ''' </summary>
    ''' <returns></returns>
    Public Overrides Function ToString() As String
        Return CreateRecipientString(Me.Name, Me.Address)
    End Function

    Friend Shared Function JoinToListWithSmtpFormat(recipients As List(Of EMailAddress)) As String
        If recipients Is Nothing Then Return Nothing
        Dim Result As New System.Text.StringBuilder
        For Each recipient As EMailAddress In recipients
            If Result.Length <> 0 Then
                Result.Append(",")
            End If
            Result.Append(recipient.ToString)
        Next
        Return Result.ToString
    End Function

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
    Friend Shared Function CreateRecipientString(name As String, address As String) As String
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
        Return SmtpUtils.IIf(Of String)(name <> Nothing, """" & name & """ ", "") & "<" & address & ">"
    End Function

    ''' <summary>
    '''     Create a valid receipient string for the address lists parameters of method SendEMail
    ''' </summary>
    ''' <param name="address">The e-mail address of the receipient</param>
    ''' <remarks>
    '''     <para>In case of an empty address, an empty string will be returned, too (no ArgumentNullException).</para>
    ''' </remarks>
    Friend Shared Function CreateRecipientString(address As String) As String
        Return CreateRecipientString(Nothing, address)
    End Function

    ''' <summary>
    ''' Split an e-mail address in typical encoding for SMTP protocol into the two parts e-mail address and receipient name
    ''' </summary>
    ''' <param name="emailAddressInSmtpFormat"></param>
    ''' <remarks></remarks>
    Public Shared Function CreateFromSmtpFormat(emailAddressInSmtpFormat As String()) As List(Of EMailAddress)
        If emailAddressInSmtpFormat Is Nothing Then
            Return Nothing
        Else
            Dim Result As New List(Of EMailAddress)
            For Each Address As String In emailAddressInSmtpFormat
                Result.Add(CreateFromSmtpFormat(Address))
            Next
            Return Result
        End If
    End Function

    ''' <summary>
    ''' Split an e-mail address in typical encoding for SMTP protocol into the two parts e-mail address and receipient name
    ''' </summary>
    ''' <param name="emailAddressInSmtpFormat"></param>
    ''' <remarks></remarks>
    Public Shared Function CreateFromSmtpFormat(emailAddressInSmtpFormat As String) As EMailAddress
        Dim Address As String
        If Not emailAddressInSmtpFormat.LastIndexOf("<") = -1 Then
            Address = emailAddressInSmtpFormat.Substring(emailAddressInSmtpFormat.LastIndexOf("<"), emailAddressInSmtpFormat.Length - emailAddressInSmtpFormat.LastIndexOf("<"))
            Address = Address.Replace("<", "").Replace(">", "")
        Else
            Address = emailAddressInSmtpFormat
        End If

        Dim Name As String
        If emailAddressInSmtpFormat.LastIndexOf(" ") > 0 AndAlso emailAddressInSmtpFormat.LastIndexOf(" ") = emailAddressInSmtpFormat.LastIndexOf("<") - 1 Then
            Name = emailAddressInSmtpFormat.Substring(0, emailAddressInSmtpFormat.LastIndexOf(" "))
            Name = Name.Replace("<", "").Replace(">", "")
            If Name.StartsWith("""") AndAlso Name.EndsWith("""") Then
                Name = Name.Substring(1, Name.Length - 2)
            End If
            Name = Name.Replace("\\", ChrW(0)).Replace("\", "").Replace(ChrW(0), "\") 'Unencode all "\*" but preserve "\\" as "\"
        Else
            Name = String.Empty
        End If

        Return New EMailAddress(Name, Address)
    End Function

    ''' <summary>
    '''     Extract e-mail addresses from a receipients list constructed by method CreateReceipientString or manually
    ''' </summary>
    ''' <param name="recipientsCommaSeparated"></param>
    ''' <remarks>
    '''     Valid parameter values are e. g. 
    ''' <list>
    ''' <item>l?nc\,\.\:\;\@\[\]\(\)ma9\'\\\&lt;\&gt;kljkj &lt;my-company@my-webmailer-company.com&gt;, her\,ma9\'\\\&lt;\&gt;kljkj &lt;jon.doe@my-company.com&gt;</item>
    ''' <item>&lt;my-company@my-webmailer-company.com&gt;,&lt;jon.doe@my-company.com&gt;</item>
    ''' <item>my-company@my-webmailer-company.com,jon.doe@my-company.com</item>
    ''' </list>
    ''' </remarks>
    Friend Shared Function SplitEMailAddressesFromRecipientsList(recipientsCommaSeparated As String) As String()
        Dim Result As String() = SmtpUtils.SplitString(recipientsCommaSeparated, ","c, "\"c)
        For MyCounter As Integer = 0 To Result.Length - 1
            'Now we got strings like
            '"l?nc\,\.\:\;\@\[\]\(\)ma9\'\\\<\>kljkj <my-company@my-webmailer-company.com>"
            '" her\,ma9\'\\\<\>kljkj <jon.doe@my-company.com>"
            '" ""Jochen"" <jon.doe@my-company.com>" <------------TOCHECK!
            '" <my-company@my-webmailer-company.com>"
            '"<jon.doe@my-company.com>" 
            '"my-company@my-webmailer-company.com"
            '"jon.doe@my-company.com"
            Dim AddressStart As Integer = Result(MyCounter).LastIndexOf("<")
            If AddressStart > 0 Then
                Result(MyCounter) = Mid(Result(MyCounter), AddressStart)
            End If
            Result(MyCounter) = Trim(Result(MyCounter))
            If InStr(Result(MyCounter), "@") < 2 OrElse InStr(Result(MyCounter), "@") > Result(MyCounter).Length - 1 Then
                Result(MyCounter) = ""
            End If
            'Now we got strings like
            '"<my-company@my-webmailer-company.com>"
            '"<jon.doe@my-company.com>"
            '"<my-company@my-webmailer-company.com>"
            '"<jon.doe@my-company.com>" 
            '"my-company@my-webmailer-company.com"
            '"jon.doe@my-company.com"
            If Left(Result(MyCounter), 1) = "<" AndAlso Right(Result(MyCounter), 1) = ">" Then
                Result(MyCounter) = Mid(Result(MyCounter), 2, Len(Result(MyCounter)) - 2)
            End If
        Next
        Return Result
    End Function

    ''' <summary>
    ''' Split an e-mail address list in typical encoding for SMTP protocol into the two parts e-mail address and receipient name
    ''' </summary>
    ''' <param name="emailAddressesInSmtpFormat"></param>
    ''' <remarks></remarks>
    Friend Shared Function SplitEMailAddressesIntoEMailRecipientsFromRecipientsList(emailAddressesInSmtpFormat As String) As List(Of EMailAddress)
        Dim Addresses As String() = SmtpUtils.SplitString(emailAddressesInSmtpFormat, ","c, "\"c)
        Dim EmailReceipients As New List(Of EMailAddress)
        For MyCounter As Integer = 0 To Addresses.Length - 1
            EmailReceipients.Add(CreateFromSmtpFormat(Addresses(MyCounter)))
        Next
        Return EmailReceipients
    End Function

    Public Overrides Function GetHashCode() As Integer
        Return Me.ToString.GetHashCode()
    End Function

    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing OrElse GetType(EMailAddress).IsInstanceOfType(obj) = False Then
            Return False
        Else
            Return Me.ToString = CType(obj, EMailAddress).ToString
        End If
    End Function

End Class