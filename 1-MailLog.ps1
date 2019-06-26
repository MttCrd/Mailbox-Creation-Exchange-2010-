$Mail = '' #Smtp or Alias (User to send the Log mail)

$smtpServer = ""

$user = get-mailbox $Mail

$DisplayName = Get-Mailbox $Alias | select DisplayName

$Dn = $DisplayName.DisplayName

$C = $CompanyUser.Company
$E = $employee.EmployeeType

$reciever = $user.primarysmtpaddress
$user
$reciever

$msg = new-object Net.Mail.MailMessage

$smtp = new-object Net.Mail.SmtpClient($smtpServer)

$msg.From = "" #Smtp (User who sends the email)
$msg.To.Add("$reciever")
$msg.subject = "New User Creation $Dn"
$msg.IsBodyHtml = $True

$body = @"
<html>
<body>

DisplayName             : $Dn <br><br>
Smtp                    : $SmtpTxt <br><br>
LinkedMasterAccount     : $LinkedMasterAccountTxt <br><br>
UserPrincipalName       : $UserNameTxt <br><br>
Database                : $Database <br><br>
EmployeeType            : $E <br><br>
Company                 : $C <br><br>

</body>
</html>
"@

$msg.body = $body
$smtp.Send($msg)
$msg.dispose();
