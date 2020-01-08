Imports CompuMaster.Net.Smtp

Public Module DemoApp

    Public Sub Main()

        My.Settings.SmtpServerName = InputBox("SmtpServerName",, My.Settings.SmtpServerName)
        My.Settings.TestRecipientAddress = InputBox("TestRecipientAddress",, My.Settings.TestRecipientAddress)
        My.Settings.Save()

        Dim MySmtpWorker As New SmtpWorker With {
            .SmtpServerName = My.Settings.SmtpServerName,
            .SmtpServerPort = My.Settings.SmtpServerPort,
            .SmtpAuthType = My.Settings.SmtpAuthType,
            .SmtpUserName = My.Settings.SmtpUsername,
            .SmtpPassword = My.Settings.SmtpPassword
        }

        Dim MyAttachment As New EMailAttachment("logo_64x64.png", "logo")

        Dim SendResult As CompuMaster.Net.Smtp.SmtpSendResult = MySmtpWorker.SendToSingeRecipient(My.Settings.TestRecipientName, My.Settings.TestRecipientAddress, "TestSingle", "plain body", "<h1>html body</h1><img src=""cid:" & MyAttachment.PlaceholderInMhtmlToBeReplacedByContentID & """>", My.Settings.TestSenderName, My.Settings.TestSenderAddress, "utf-8", New EMailAttachment() {MyAttachment}, EMails.Priority.High, EMails.Sensitivity.Status_Normal, False, False, Nothing)
        Console.WriteLine("Success=" & SendResult.Success)

        Dim MyRecipient As String = EMailRecipient.CreateRecipientString(My.Settings.TestRecipientName, My.Settings.TestRecipientAddress)
        SendResult = MySmtpWorker.SendToMultipleRecipients(MyRecipient, "", "", "TestMultiple", "plain body", "<h1>html body</h1><img src=""cid:" & MyAttachment.PlaceholderInMhtmlToBeReplacedByContentID & """>", My.Settings.TestSenderName, My.Settings.TestSenderAddress, "utf-8", New EMailAttachment() {MyAttachment}, EMails.Priority.High, EMails.Sensitivity.Status_Normal, False, False, Nothing)
        Console.WriteLine("Success=" & SendResult.Success)

    End Sub

End Module
