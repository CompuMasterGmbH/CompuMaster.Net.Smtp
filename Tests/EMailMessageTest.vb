Imports NUnit.Framework
Imports CompuMaster.Net.Smtp

<TestFixture> Public Class EMailMessageTest

    <Test> Public Sub Initialization()
        Dim MyMessage1 As New EMailMessage()
        Dim MyMessage3 As New EMailMessage(New EMailAddress("Jon Doe", "jon.doe@domain.invalid"), "Test without embedded image attachment", "plain body", "<h1>html body</h1>", New EMailAddress("Sender Doe", "sender.doe@domain.invalid"), New EMailAddress("Reply Doe", "reply.doe@domain.invalid"))
        Dim MyAttachments As New List(Of EMailAttachment)(New EMailAttachment() {New EMailAttachment("logo_64x64.png", "logo")})
        Dim AdditionalHeaders As New Specialized.NameValueCollection
        AdditionalHeaders("X-CustomTestHeader") = "Test"
        Dim MyMessage4 As New EMailMessage(New EMailAddress("Jon Doe", "jon.doe@domain.invalid"), "Test with embedded image attachment", "plain body", "<h1>html body</h1><img src=""cid:" & MyAttachments(0).PlaceholderInMhtmlToBeReplacedByContentID & """>", New EMailAddress("Sender Doe", "sender.doe@domain.invalid"), New EMailAddress("Reply Doe", "reply.doe@domain.invalid"), System.Text.Encoding.UTF32, MyAttachments, EMails.Priority.High, EMails.Sensitivity.CompanyConfidential, True, True, AdditionalHeaders)
        Dim MyMessage5 As New EMailMessage(New List(Of EMailAddress)(New EMailAddress() {New EMailAddress("Jon Doe", "jon.doe@domain.invalid")}), CType(Nothing, List(Of EMailAddress)), CType(Nothing, List(Of EMailAddress)), "Test with embedded image attachment", "plain body", "<h1>html body</h1><img src=""cid:" & MyAttachments(0).PlaceholderInMhtmlToBeReplacedByContentID & """>", New EMailAddress("Sender Doe", "sender.doe@domain.invalid"), New EMailAddress("Reply Doe", "reply.doe@domain.invalid"), System.Text.Encoding.UTF32, MyAttachments, EMails.Priority.High, EMails.Sensitivity.CompanyConfidential, True, True, AdditionalHeaders)
        Dim MyMessage5Serialized As String = MyMessage5.ExportToXml
        Dim MyMessage2 As New EMailMessage(MyMessage5Serialized)
        System.Console.WriteLine("## MyMessage5 (serialized)")
        Dim Msg5Export1 As String = MyMessage5.ExportToXml
        System.Console.WriteLine(Msg5Export1)
        System.Console.WriteLine("## MyMessage5 (re-serialized)")
        Dim Msg5Export2 As String = MyMessage5.ExportToXml
        System.Console.WriteLine(Msg5Export2)
        Assert.AreEqual(Msg5Export1, Msg5Export2)
        System.Console.WriteLine("## MyMessage2 (re-serialized)")

        Assert.AreEqual(MyMessage5.Subject, MyMessage2.Subject)
        Assert.AreEqual(MyMessage5.BodyHtml, MyMessage2.BodyHtml)
        Assert.AreEqual(MyMessage5.BodyPlainText, MyMessage2.BodyPlainText)
        Assert.AreEqual(MyMessage5.From, MyMessage2.From)
        Assert.AreEqual(MyMessage5.ReplyTo, MyMessage2.ReplyTo)
        Assert.AreEqual(MyMessage5.AdditionalHeaders, MyMessage2.AdditionalHeaders)
        Assert.AreEqual(MyMessage5.AdditionalHeaders.ToString, MyMessage2.AdditionalHeaders.ToString)
        Assert.AreEqual(MyMessage5.AdditionalHeaders.Count, MyMessage2.AdditionalHeaders.Count)
        Assert.AreEqual(1, MyMessage2.AdditionalHeaders.Count)
        Assert.AreEqual("X-CustomTestHeader", MyMessage2.AdditionalHeaders.AllKeys(0))
        Assert.AreEqual("Test", MyMessage2.AdditionalHeaders(MyMessage2.AdditionalHeaders.AllKeys(0)))
        Assert.AreEqual(MyMessage5.Priority, MyMessage2.Priority)
        Assert.AreEqual(MyMessage5.Sensitivity, MyMessage2.Sensitivity)
        Assert.AreEqual(MyMessage5.RequestReadingConfirmation, MyMessage2.RequestReadingConfirmation)
        Assert.AreEqual(MyMessage5.RequestTransmissionConfirmation, MyMessage2.RequestTransmissionConfirmation)
        Assert.AreEqual(MyMessage5.EMailAttachments.Count, MyMessage2.EMailAttachments.Count)
        Assert.AreEqual(1, MyMessage2.EMailAttachments.Count)
        Assert.AreEqual(MyMessage5.EMailAttachments(0).RawDataOriginFilename, MyMessage2.EMailAttachments(0).RawDataOriginFilename)
        'Assert.AreEqual(MyMessage5.EMailAttachments(0).RawDataFilename, MyMessage2.EMailAttachments(0).RawDataFilename)
        Assert.AreEqual(MyMessage5.EMailAttachments(0).RawData, MyMessage2.EMailAttachments(0).RawData)
        Assert.AreEqual(MyMessage5.MessageEncoding.WebName, MyMessage2.MessageEncoding.WebName)
        Assert.AreEqual(MyMessage5.ToString, MyMessage2.ToString)

        Dim Msg2Export1 As String = MyMessage5.ExportToXml
        System.Console.WriteLine(Msg2Export1)
        Assert.AreEqual(Msg5Export1, Msg2Export1)
    End Sub

End Class
