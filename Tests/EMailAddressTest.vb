Imports NUnit.Framework
Imports CompuMaster.Net.Smtp

<TestFixture> Public Class EMailAddressTest

    <Test> Public Sub IsEqual()
        Dim addr0 As EMailAddress = Nothing
        Dim addr1 As New EMailAddress(String.Empty)
        Dim addr2 As New EMailAddress("abc@dom123.com")
        Dim addr3 As New EMailAddress("ABC", "abc@dom123.com")
        Dim addr3a As New EMailAddress("ABC", "abc@dom123.com")
        Dim addr4 As New EMailAddress("abc", "abc@dom123.com")
        Dim addr5 As New EMailAddress("abc", "abc@DOM123.com")

        Assert.AreEqual(Nothing, addr0)
        Assert.AreNotEqual(addr0, addr1)
        Assert.AreNotEqual(addr1, addr2)
        Assert.AreNotEqual(addr2, addr3)
        Assert.AreNotEqual(addr3, addr4)
        Assert.AreNotEqual(addr3, addr5)
        Assert.AreEqual(addr3a, addr3)

#Disable Warning IDE0028 ' Initialisierung der Sammlung vereinfachen
        Dim l As New List(Of EMailAddress)
#Enable Warning IDE0028 ' Initialisierung der Sammlung vereinfachen
        l.Add(addr3)
        Assert.AreEqual(True, l.Contains(addr3))
        Assert.AreEqual(True, l.Contains(addr3a))
        Assert.AreEqual(False, l.Contains(addr2))
    End Sub

End Class
