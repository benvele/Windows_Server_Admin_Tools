<#
  .SYNOPSIS
    Supplements default Windows password expiration notifications with email notifications in Active Directory environments.
  .DESCRIPTION
    Queries AD to get a list of all users passwords that will expire within the amount of days set in the $daysbeforeexpiretonotify variable.
    This is intended to be run as a scheduled task.
  .NOTES
    A text file with a Hex version of a SecureString encoded password is needed for this to work. To get the Hex version of the password needed, please use the following
    code on the machine the script will be ran:
      1) $Credential = Get-Credential
      2) $Credentail.Password | ConvertFrom-SecureString
      3) Save the output of the last command in a text file
      4) Enter path of text file into $Email_Password Variable
      
    To use the EventLog function at the bottom you will need to run the following one-time command first:
      -> Write-EventLog -LogName Application -Source "Password Expiration Script"
      
    The Active Directory Module is needed to run the command. See the following link for instructions to install:
    https://docs.microsoft.com/en-us/powershell/module/addsadministration/?view=win10-ps
    
    TIP: You can send emails using HTML. Create an email in outlook and send it to yourself. From the received email, view source and copy the HTML. 
    Put the HTML in the $Email_Body variable and add the -BodyAsHtml flag to the Send-MailMessage command.
    
    NOTE: In some cases I have to specify the server in the Get-ADUser command. To do this just add the -Server flag with the IP or Hostname as a String at the end 
    of the command.
#>

#These variables are the ones that are most likely to be updated. Email Subject and Body are set in foreach statement
$daysbeforeexpiretonotify = 14
$Email = "example@example.com"
$SMTPServer = "your.smtp.server.here"


$users = ''
$Credential = New-Object -TypeName PSCredential -ArgumentList $Email, (Get-Content -Path "C:\PathTo\Scripts\Password_Expiration\hash.txt" | ConvertTo-SecureString)
$sentToEmails = ''
$startTime = get-date 
$users = Get-ADUser -filter { Enabled -eq $True -and PasswordNeverExpires -eq $False } -Properties "msDS-UserPasswordExpiryTimeComputed",mail ` -searchbase "OU=Company,DC=company,DC=local" |   
   where { $_."msDS-UserPasswordExpiryTimeComputed" -lt ((get-date).AddDays($daysbeforeexpiretonotify).ToFileTime()) -and $_."msDS-UserPasswordExpiryTimeComputed" -gt ((get-date).ToFileTime()) }
   Select-Object "Name",  
                 "Mail" |  
    sort-object name  
  
foreach ($user in $users) {
    $Email_Subject = "Your password will expire in $([int](($user.'msDS-UserPasswordExpiryTimeComputed' - $now) / 864000000000)) days"
    $Email_Body = "$($user.Name), your password will expire at $([datetime]::FromFileTime($user.'msDS-UserPasswordExpiryTimeComputed')) (EST)."
    Send-MailMessage -To $user.mail -From $Email  -Subject $Email_Subject -Body $Email_Body -Credential ($Credential) -SmtpServer $SMTPServer -Port 587
    $sentToEmails += $user.Mail + ", "
}

$endTime = get-date
Write-EventLog -LogName Application -Source "Password Expiration Script" -EntryType Information -EventId 100 -Message "Password Expiration script started at $($startTime). Password expiration notifications were sent to $($sentToEmails) and the script ended at $($endTime)."
