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

    <Ignore("Manual test call only - requires opened Adobe Reader with opened test file")>
    <Test> Public Sub ExportXmlWithOpenedAttachment()
        Dim TestPdfPath As String = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location), "AttachmentTest.pdf")
        System.Diagnostics.Process.Start(TestPdfPath)
        System.Threading.Thread.Sleep(5000)
        Dim MyAttachments As New List(Of EMailAttachment)(New EMailAttachment() {New EMailAttachment(TestPdfPath)})
        Dim MyMessage As New EMailMessage(New EMailAddress("Jon Doe", "jon.doe@domain.invalid"), "Test with embedded image attachment", "plain body", "<h1>html body</h1>", New EMailAddress("Sender Doe", "sender.doe@domain.invalid"), New EMailAddress("Reply Doe", "reply.doe@domain.invalid"), System.Text.Encoding.UTF32, MyAttachments, EMails.Priority.High, EMails.Sensitivity.CompanyConfidential, True, True, Nothing)
        Assert.AreEqual(1, MyMessage.EMailAttachments.Count)
        Dim MyMessageSerialized As String = MyMessage.ExportToXml
        Assert.AreEqual(1, MyMessage.EMailAttachments.Count)
        Assert.Pass()
    End Sub

    <Test> Public Sub FixHtmlContentIDs()
        Dim Attachments As List(Of CompuMaster.Net.Smtp.EMailAttachment)

        EMailAttachment.FixHtmlContentIDs("", Nothing)

        Attachments = New List(Of CompuMaster.Net.Smtp.EMailAttachment)
        EMailAttachment.FixHtmlContentIDs("", Attachments)
        Assert.That(Attachments.Count, [Is].EqualTo(0))

        Attachments = New List(Of CompuMaster.Net.Smtp.EMailAttachment)
        Attachments.Add(New CompuMaster.Net.Smtp.EMailAttachment())
        EMailAttachment.FixHtmlContentIDs(Nothing, Attachments)
        Assert.That(Attachments.Count, [Is].EqualTo(1))

        Attachments = CreateEMailAttachmentsSample1()
        EMailAttachment.FixHtmlContentIDs(Nothing, Attachments)
        Assert.That(Attachments.Count, [Is].EqualTo(0))

        Attachments = CreateEMailAttachmentsSample1()
        EMailAttachment.FixHtmlContentIDs("<html>" &
                                          "<img src=""cid:" & Attachments(0).PlaceholderInMhtmlToBeReplacedByContentID & """/>" &
                                          "<img src=""cid:" & Attachments(1).PlaceholderInMhtmlToBeReplacedByContentID & """/>" &
                                          "<img src=""cid:" & Attachments(2).PlaceholderInMhtmlToBeReplacedByContentID & """/>" &
                                          "<img src=""cid:" & Attachments(3).PlaceholderInMhtmlToBeReplacedByContentID & """/>" &
                                          "<img src=""cid:" & Attachments(4).PlaceholderInMhtmlToBeReplacedByContentID & """/>" &
                                          "</html>", Attachments)
        Assert.That(Attachments.Count, [Is].EqualTo(5))

        Attachments = CreateEMailAttachmentsSample1()
        EMailAttachment.FixHtmlContentIDs("<html>" &
                                          "<img src=""cid:" & Attachments(0).PlaceholderInMhtmlToBeReplacedByContentID & """/>" &
                                          "</html>", Attachments)
        Assert.That(Attachments.Count, [Is].EqualTo(1))

        Attachments = CreateEMailAttachmentsSample2()
        EMailAttachment.FixHtmlContentIDs("<html>" &
                                          "</html>", Attachments)
        Assert.That(Attachments.Count, [Is].EqualTo(2))

        Attachments = CreateEMailAttachmentsSample2()
        EMailAttachment.FixHtmlContentIDs("<html>" &
                                          "<img src=""cid:" & Attachments(0).PlaceholderInMhtmlToBeReplacedByContentID & """/>" &
                                          "<img src=""cid:" & Attachments(3).PlaceholderInMhtmlToBeReplacedByContentID & """/>" &
                                          "</html>", Attachments)
        Assert.That(Attachments.Count, [Is].EqualTo(4))
    End Sub

    Private Function CreateEMailAttachmentsSample1() As List(Of CompuMaster.Net.Smtp.EMailAttachment)
        Dim Attachments As New List(Of CompuMaster.Net.Smtp.EMailAttachment)
        Attachments.Add(New CompuMaster.Net.Smtp.EMailAttachment() With {
            .FilePath = Nothing,
            .RawDataFilename = "ProjectDetails_CM-4825_Fachhandwerk360 Dreier+Herber GbR (DEB300336) _ Helpdesk _ IT Management + Support.pdf.pdf",
            .RawDataOriginFilename = "ProjectDetails_CM-4825_Fachhandwerk360 Dreier+Herber GbR (DEB300336) _ Helpdesk _ IT Management + Support.pdf.pdf",
            .RawData = New Byte() {1, 2, 3},
            .PlaceholderInMhtmlToBeReplacedByContentID = "ed642552-3e6c-41ac-8984-85b24ff2ba0c"
        })
        Attachments.Add(New CompuMaster.Net.Smtp.EMailAttachment() With {
            .FilePath = Nothing,
            .RawDataFilename = "ProjectDetails_CM-4825_Overview_Fachhandwerk360 Dreier+Herber GbR (DEB300336) _ Helpdesk _ IT Management + Support.pdf.pdf",
            .RawDataOriginFilename = "ProjectDetails_CM-4825_Overview_Fachhandwerk360 Dreier+Herber GbR (DEB300336) _ Helpdesk _ IT Management + Support.pdf.pdf",
            .RawData = New Byte() {1, 2, 3},
            .PlaceholderInMhtmlToBeReplacedByContentID = "62ac3106-82c2-44dd-bc03-cc29f674fc0b"
        })
        Attachments.Add(New CompuMaster.Net.Smtp.EMailAttachment() With {
            .FilePath = Nothing,
            .RawDataFilename = "ProjectDetails_CM-4825_MainCategories_Fachhandwerk360 Dreier+Herber GbR (DEB300336) _ Helpdesk _ IT Management + Support.pdf.pdf",
            .RawDataOriginFilename = "ProjectDetails_CM-4825_MainCategories_Fachhandwerk360 Dreier+Herber GbR (DEB300336) _ Helpdesk _ IT Management + Support.pdf.pdf",
            .RawData = New Byte() {1, 2, 3},
            .PlaceholderInMhtmlToBeReplacedByContentID = "4ff43e11-fa0a-49f7-984c-7bd2869d87f2"
        })
        Attachments.Add(New CompuMaster.Net.Smtp.EMailAttachment() With {
            .FilePath = "RE559902886.pdf",
            .RawDataFilename = Nothing,
            .RawDataOriginFilename = Nothing,
            .RawData = Nothing,
            .PlaceholderInMhtmlToBeReplacedByContentID = "4b38e763-79b0-470f-80bd-d493c6b2eb9f"
        })
        Attachments.Add(New CompuMaster.Net.Smtp.EMailAttachment() With {
            .FilePath = "RE559902886_printer.pdf",
            .RawDataFilename = Nothing,
            .RawDataOriginFilename = Nothing,
            .RawData = Nothing,
            .PlaceholderInMhtmlToBeReplacedByContentID = "676fdb55-af32-4c13-b70a-933e4d20dedf"
        })
        Return Attachments
    End Function

    ''' <summary>
    ''' Sample attachments: 2 with placeholder + 2 without placeholder
    ''' </summary>
    ''' <returns></returns>
    Private Function CreateEMailAttachmentsSample2() As List(Of CompuMaster.Net.Smtp.EMailAttachment)
        Dim Attachments As New List(Of CompuMaster.Net.Smtp.EMailAttachment)
        Attachments.Add(New CompuMaster.Net.Smtp.EMailAttachment(New Byte() {1, 2, 3}, "ProjectDetails_CM-4825_Fachhandwerk360 Dreier+Herber GbR (DEB300336) _ Helpdesk _ IT Management + Support.pdf.pdf", "ed642552-3e6c-41ac-8984-85b24ff2ba0c"))
        Attachments.Add(New CompuMaster.Net.Smtp.EMailAttachment(New Byte() {1, 2, 3}, "ProjectDetails_CM-4825_Overview_Fachhandwerk360 Dreier+Herber GbR (DEB300336) _ Helpdesk _ IT Management + Support.pdf.pdf"))
        Attachments.Add(New CompuMaster.Net.Smtp.EMailAttachment("RE559902886.pdf"))
        Attachments.Add(New CompuMaster.Net.Smtp.EMailAttachment("RE559902886_printer.pdf", "676fdb55-af32-4c13-b70a-933e4d20dedf"))
        Return Attachments
    End Function

End Class
