# CompuMaster.Network.Smtp
SMTP messaging with support for alternative body and embedded images

## Simple download/installation using NuGet
```nuget
Install-Package CompuMaster.Network.Smtp
```

## Sample: Quick initialization of SMTP component
```vb.net
Dim MySmtpWorker As New SmtpWorker()
MySmtpWorker.SmtpServerName = "localhost"
MySmtpWorker.SmtpServerPort = 465
MySmtpWorker.SmtpAuthType = SmtpWorker.SmtpAuthTypes.LoginSsl,
MySmtpWorker.SmtpUserName = "user"
MySmtpWorker.SmtpPassword = "password"
```

## Sample: Prepare a message with alternative message parts and attachment for embedded image
```vb.net
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
```

## Sample: Send demo e-mail and check for exceptions
```vb.net
Dim SendResult As CompuMaster.Net.Smtp.SmtpSendResult
SendResult = MySmtpWorker.SendMessage(MyMessage)
If SendResult.Success Then
    System.Console.WriteLine("EMAIL SENT SUCCESSFULLY :-)")
Else
    System.Console.WriteLine("FAILURE: " & SendResult.Exception.ToString)
End If
```
