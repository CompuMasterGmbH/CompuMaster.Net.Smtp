Imports CompuMaster.Net.Smtp

Public Module DemoApp

    Public Sub Main()

        Dim MySmtpWorker As New SmtpWorker

        'Ask user of DemoApp for most important settings (optionally add more if required for your demo environment)
        My.Settings.SmtpServerName = InputBox("SmtpServerName",, My.Settings.SmtpServerName)
        If My.Settings.SmtpServerName = Nothing Then Return
        My.Settings.TestRecipientAddress = InputBox("TestRecipientAddress",, My.Settings.TestRecipientAddress)
        If My.Settings.TestRecipientAddress = Nothing Then Return
        My.Settings.Save()

        'Setup SmtpWorker settings
        MySmtpWorker.SmtpServerName = My.Settings.SmtpServerName
        MySmtpWorker.SmtpServerPort = My.Settings.SmtpServerPort
        MySmtpWorker.SmtpAuthType = My.Settings.SmtpAuthType
        MySmtpWorker.SmtpUserName = My.Settings.SmtpUsername
        MySmtpWorker.SmtpPassword = My.Settings.SmtpPassword

        'Prepare email message
        Dim MyAttachment As New EMailAttachment("logo_64x64.png", "logo")
        Dim Attachments As New List(Of EMailAttachment)(New EMailAttachment() {MyAttachment})
        Dim Recipient As New EMailAddress(My.Settings.TestRecipientName, My.Settings.TestRecipientAddress)
        Dim Sender As New EMailAddress(My.Settings.TestSenderName, My.Settings.TestSenderAddress)
        Dim ReplyToAddress As EMailAddress = Sender
        Dim MyMessage As New EMailMessage(Recipient,
                                          "TestSubject",
                                          "Plain body",
                                          "<h1>HTML body with embedded image</h1><img src=""cid:" & MyAttachment.PlaceholderInMhtmlToBeReplacedByContentID & """>",
                                          Sender,
                                          ReplyToAddress,
                                          System.Text.Encoding.UTF8,
                                          Attachments,
                                          EMails.Priority.High,
                                          EMails.Sensitivity.Normal,
                                          False,
                                          False,
                                          CType(Nothing, Specialized.NameValueCollection))

        'Send demo e-mail and check for exceptions
        Dim SendResult As CompuMaster.Net.Smtp.SmtpSendResult
        SendResult = MySmtpWorker.SendMessage(MyMessage)
        If SendResult.Success Then
            System.Console.WriteLine("EMAIL SENT SUCCESSFULLY :-)")
        Else
            System.Console.WriteLine("FAILURE: " & SendResult.Exception.ToString)
        End If

    End Sub

End Module
