function Send-MailMessage-Net
{
    param (
        [string[]]$To,
        [parameter(mandatory)][string]$Subject,
        [parameter(mandatory)][string]$Body,
        [string]$AttachmentFilePath
    )

    # SMTP アカウント
    #$gmailAccount = "v.remote.desktop@gmail.com"
    #$gmailPassword = "V0tir0@sds"
    $smtpAccount = "yuhashimoto@asgent.co.jp"

    # メッセージ生成
    $smtpMessage = New-Object Net.Mail.MailMessage($null)
    # 送信者
    $smtpMessage.From = New-Object Net.Mail.MailAddress($smtpAccount)
    # 受信者
    for ($i = 0; $i -lt $To.Length; $i++)
    {
        $smtpMessage.To.Add($To[$i])
    }
    # 件名
    $smtpMessage.Subject = $Subject
    # 本文
    $smtpMessage.Body = $Body
    # 添付ファイル
    if ($AttachmentFilePath)
    {
        $smtpAttachmentFile = New-Object Net.Mail.Attachment($AttachmentFilePath)
        $smtpMessage.Attachments.Add($smtpAttachmentFile)
    }

    # メール送信
    #$smtpServer = "smtp.gmail.com"
    #$smtpPort = 587
    $smtpServer = "asgent-co-jp.mail.protection.outlook.com"
    $smtpPort = 25
    $smtpClient = New-Object Net.Mail.SmtpClient($smtpServer, $smtpPort)
    $smtpClient.EnableSsl = $True
    $smtpClient.Send($smtpMessage)

}