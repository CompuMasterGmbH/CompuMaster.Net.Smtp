# CompuMaster.Network.Smtp
SMTP messaging with support for alternative body and embedded images

## Sample: Quick initialization of SMTP component
```vb.net
Dim MySmtpWorker As New SmtpWorker()
MySmtpWorker.SmtpServerName = "localhost"
MySmtpWorker.SmtpServerPort = 465
MySmtpWorker.SmtpAuthType = "LOGIN-SSL"
MySmtpWorker.SmtpUserName = "user"
MySmtpWorker.SmtpPassword = "password"
```

## Sample: Send with alternative message parts and attachment for embedded image
```vb.net
Dim MyAttachment As New EMailAttachment("logo_64x64.png", "logo")

Dim SendResult As CompuMaster.Net.Smtp.SmtpSendResult = MySmtpWorker.SendToSingeRecipient(My.Settings.TestRecipientName, My.Settings.TestRecipientAddress, "TestSingle", "plain body", "<h1>html body</h1><img src=""cid:" & MyAttachment.PlaceholderInMhtmlToBeReplacedByContentID & """>", My.Settings.TestSenderName, My.Settings.TestSenderAddress, "utf-8", New EMailAttachment() {MyAttachment}, EMails.Priority.High, EMails.Sensitivity.Status_Normal, False, False, Nothing)
System.Console.WriteLine("Success=" & SendResult.Success)
```

## Sample: Send to multiple recipients
```vb.net
Dim MyRecipient1TO As String = EMailRecipient.CreateRecipientString(My.Settings.TestRecipientName, My.Settings.TestRecipientAddress)
Dim MyRecipient2CC As String = EMailRecipient.CreateRecipientString(My.Settings.TestRecipientName, My.Settings.TestRecipientAddress)
SendResult = MySmtpWorker.SendToMultipleRecipients(MyRecipient1TO, MyRecipient2CC, "", "TestMultiple", "plain body", "<h1>html body</h1><img src=""cid:" & MyAttachment.PlaceholderInMhtmlToBeReplacedByContentID & """>", My.Settings.TestSenderName, My.Settings.TestSenderAddress, "utf-8", New EMailAttachment() {MyAttachment}, EMails.Priority.High, EMails.Sensitivity.Status_Normal, False, False, Nothing)
System.Console.WriteLine("Success=" & SendResult.Success)
```
