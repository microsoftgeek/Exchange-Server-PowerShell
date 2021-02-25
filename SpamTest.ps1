#O365 transport relay

$cred = Get-Credential ; 
1..11000 | ForEach-Object {Send-MailMessage `
    -To UserTo@Tenant.onmicrosoft.com `
    -From UserFrom@tenant.com `
    -SmtpServer smtp.office365.com `
    -usessl -Credential $cred -Port 587 `
    -Subject "Test Message $_" `
    -Body "This is the body of Message $_" `
    ; write-host “Sending Message $_”; Start-Sleep -Milliseconds 500}


#Using Outlook native application for relay

Start-Process outlook.exe

$o = New-Object -ComObject Outlook.Application
1..5000 | ForEach-Object { `
  $mail = $o.CreateItem(0)
  $mail.Subject = "Test message $_"
  $mail.Body = "This is the body of test message $_"
  $mail.To = "User@tenant.onmicrosoft.com"
  $mail.Send()
}
$o.Quit()


