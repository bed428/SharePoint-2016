function Send-Mail($subject, $body){
        #SMTP server name
        $smtpServer = "smtp.domain.com"
 
        #Creating a Mail object
        $msg = new-object Net.Mail.MailMessage
 
        #Creating SMTP server object
        $smtp = new-object Net.Mail.SmtpClient($smtpServer)
 
        #Email structure
        $msg.From = "OutageAlert@portal.domain.com"
        $msg.ReplyTo = "OutageAlert@portal.domain.com"
        $msg.To.Add(“first.last@domain.com”)
        $msg.To.Add(“first.last@domain.com”)
        #$msg.To.Add(“first.last@domain.com")
        #$msg.To.Add(“first.last@domain.com”)
        #$msg.CC.Add()
        $msg.Priority = [System.Net.Mail.MailPriority]::High
        $msg.subject = $subject
        $msg.body = $body
 
        #Sending email
        $smtp.Send($msg)
    }