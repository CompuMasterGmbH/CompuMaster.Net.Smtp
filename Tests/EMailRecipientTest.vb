Imports NUnit.Framework
Imports CompuMaster.Net.Smtp

<TestFixture> Public Class EMailRecipientTest

    <Test> Public Sub AddressTest()
        Dim Name As String
        Dim Address As String
        Dim Recipient As EMailRecipient
        Name = "#äaflaeÄ'LP""ELÄPLóËךכףגךשדחכaöjfsdöl@afei"
        Address = "test@domain.invalid"
        Recipient = New EMailRecipient(Name, Address)
        Assert.AreEqual(Name, Recipient.Name)
        Assert.AreEqual(Address, Recipient.Address)
        Dim SmtpAddress As String = Recipient.ToString
        Assert.AreEqual("", SmtpAddress)
        Dim RecipientReloaded As EMailRecipient = EMailRecipient.CreateFromSmtpFormat(SmtpAddress)
        Assert.AreEqual(Name, RecipientReloaded.Name)
        Assert.AreEqual(Address, RecipientReloaded.Address)
    End Sub

End Class
