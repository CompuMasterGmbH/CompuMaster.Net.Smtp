Imports NUnit.Framework
Imports CompuMaster.Net.Smtp

<TestFixture> Public Class EMailRecipientTest

    <Test> Public Sub AddressTest()
        Dim Name As String
        Dim Address As String
        Dim Recipient As EMailAddress
        Name = "#äaflaeÄ'LP""ELÄ PLóËךכףגךש\דחכaöjfsdöl@afei"
        Address = "test@domain.invalid"
        Recipient = New EMailAddress(Name, Address)
        Assert.AreEqual(Name, Recipient.Name)
        Assert.AreEqual(Address, Recipient.Address)
        Dim SmtpAddress As String = Recipient.ToString
        Dim ExpectedAddress As String = New System.Net.Mail.MailAddress(Address, """" & Name.Replace("""", "'").Replace("\", "\\").Replace("@", "\@") & """").ToString
        Assert.AreEqual(ExpectedAddress, SmtpAddress)
        Dim RecipientReloaded As EMailAddress = EMailAddress.CreateFromSmtpFormat(SmtpAddress)
        Assert.AreEqual(Name.Replace("""", "'"), RecipientReloaded.Name)
        Assert.AreEqual(Address, RecipientReloaded.Address)
    End Sub

    <Test> Public Sub CreateFromSmtpFormat()
        Dim Name As String
        Dim Address As String
        Name = "#äaflaeÄ'LP""ELÄ PLóËךכףגךש\דחכaöjfsdöl@afei"
        Address = "test@domain.invalid"
        Dim TestAddressStyle1 As String = New System.Net.Mail.MailAddress(Address, """" & Name.Replace("""", "'").Replace("\", "\\").Replace("@", "\@") & """").ToString
        Dim TestAddressStyle2 As String = New System.Net.Mail.MailAddress(Address, Name.Replace("""", "'").Replace("\", "\\").Replace("@", "\@")).ToString
        Dim TestAddressStyle3 As String = New System.Net.Mail.MailAddress(Address, Name).ToString
        Assert.AreEqual(Name.Replace("""", "'"), EMailAddress.CreateFromSmtpFormat(TestAddressStyle1).Name)
        Assert.AreEqual(Address, EMailAddress.CreateFromSmtpFormat(TestAddressStyle1).Address)
        Assert.AreEqual(Name.Replace("""", "'"), EMailAddress.CreateFromSmtpFormat(TestAddressStyle2).Name)
        Assert.AreEqual(Address, EMailAddress.CreateFromSmtpFormat(TestAddressStyle2).Address)
        Assert.AreEqual(Name.Replace("\", ""), EMailAddress.CreateFromSmtpFormat(TestAddressStyle3).Name)
        Assert.AreEqual(Address, EMailAddress.CreateFromSmtpFormat(TestAddressStyle3).Address)
        Assert.AreEqual(Nothing, EMailAddress.CreateFromSmtpFormat(CType(Nothing, String())))
#Disable Warning IDE0079 ' Unnötige Unterdrückung entfernen
#Disable Warning CA1825 ' Avoid zero-length array allocations.
        Assert.AreEqual(0, EMailAddress.CreateFromSmtpFormat(New String() {}).Count)
        Dim ConvertedAddressList As List(Of EMailAddress) = EMailAddress.CreateFromSmtpFormat(New String() {TestAddressStyle1, TestAddressStyle2, TestAddressStyle3})
#Enable Warning CA1825 ' Avoid zero-length array allocations.
#Enable Warning IDE0079 ' Unnötige Unterdrückung entfernen
        Assert.AreEqual(3, ConvertedAddressList.Count)
        Assert.AreEqual(Address, ConvertedAddressList(0).Address)
        Assert.AreEqual(Address, ConvertedAddressList(1).Address)
        Assert.AreEqual(Address, ConvertedAddressList(2).Address)
        Assert.AreEqual(Name.Replace("""", "'"), ConvertedAddressList(0).Name)
        Assert.AreEqual(Name.Replace("""", "'"), ConvertedAddressList(1).Name)
        Assert.AreEqual(Name.Replace("\", ""), ConvertedAddressList(2).Name)
    End Sub

End Class
