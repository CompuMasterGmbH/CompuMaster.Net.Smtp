Imports CompuMaster.Net.Smtp

Public Module DemoApp

    Public Sub Main()

        'Ask user of DemoApp for most important settings (optionally add more if required for your demo environment)
        My.Settings.SmtpServerName = InputBox("SmtpServerName",, My.Settings.SmtpServerName)
        My.Settings.TestRecipientAddress = InputBox("TestRecipientAddress",, My.Settings.TestRecipientAddress)
        My.Settings.Save()

        'Setup SmtpWorker settings
        Dim MySmtpWorker As New SmtpWorker With {
            .SmtpServerName = My.Settings.SmtpServerName,
            .SmtpServerPort = My.Settings.SmtpServerPort,
            .SmtpAuthType = My.Settings.SmtpAuthType,
            .SmtpUserName = My.Settings.SmtpUsername,
            .SmtpPassword = My.Settings.SmtpPassword
        }

        'Use a demo logo from file
        Dim MyAttachment As New EMailAttachment("logo_64x64.png", "logo")


        'Send demo e-mail using the method for a single recipient
        Dim SendResult As CompuMaster.Net.Smtp.SmtpSendResult
        SendResult = MySmtpWorker.SendToSingleRecipient(My.Settings.TestRecipientName, My.Settings.TestRecipientAddress, "TestSingle", "plain body", "<h1>html body</h1>", My.Settings.TestSenderName, My.Settings.TestSenderAddress)
        Console.WriteLine("Single #1 Success=" & SendResult.Success)
        SendResult = MySmtpWorker.SendToSingleRecipient(My.Settings.TestRecipientName, My.Settings.TestRecipientAddress, "TestSingle", "plain body", "<h1>html body</h1>", My.Settings.TestSenderName, My.Settings.TestSenderAddress, My.Settings.TestSenderName, My.Settings.TestSenderAddress)
        Console.WriteLine("Single #2 Success=" & SendResult.Success)
        SendResult = MySmtpWorker.SendToSingleRecipient(My.Settings.TestRecipientName, My.Settings.TestRecipientAddress, "TestSingle", "plain body", "<h1>html body</h1><img src=""cid:" & MyAttachment.PlaceholderInMhtmlToBeReplacedByContentID & """>", My.Settings.TestSenderName, My.Settings.TestSenderAddress, "utf-8", New EMailAttachment() {MyAttachment}, EMails.Priority.High, EMails.Sensitivity.Normal, False, False, CType(Nothing, Specialized.NameValueCollection))
        Console.WriteLine("Single #3 Success=" & SendResult.Success)
        SendResult = MySmtpWorker.SendToSingleRecipient(My.Settings.TestRecipientName, My.Settings.TestRecipientAddress, "TestSingle", "plain body", "<h1>html body</h1><img src=""cid:" & MyAttachment.PlaceholderInMhtmlToBeReplacedByContentID & """>", My.Settings.TestSenderName, My.Settings.TestSenderAddress, My.Settings.TestSenderName, My.Settings.TestSenderAddress, "utf-8", New EMailAttachment() {MyAttachment}, EMails.Priority.High, EMails.Sensitivity.Normal, False, False, CType(Nothing, Specialized.NameValueCollection))
        Console.WriteLine("Single #4 Success=" & SendResult.Success)

        'Send demo e-mail using the method for multiple recipients 
        Dim MyRecipient As String = EMailRecipient.CreateRecipientString(My.Settings.TestRecipientName, My.Settings.TestRecipientAddress)
        SendResult = MySmtpWorker.SendToMultipleRecipients(MyRecipient, "", "", "TestMultiple", "plain body", "<h1>html body</h1><img src=""cid:" & MyAttachment.PlaceholderInMhtmlToBeReplacedByContentID & """>", My.Settings.TestSenderName, My.Settings.TestSenderAddress, "utf-8", New EMailAttachment() {MyAttachment}, EMails.Priority.High, EMails.Sensitivity.Normal, False, False, Nothing)
        Console.WriteLine("Success=" & SendResult.Success)

    End Sub

End Module
