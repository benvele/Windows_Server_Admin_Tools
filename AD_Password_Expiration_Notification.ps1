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
      
    The Active Directory Module is needed to run the command. See the following link for instructions to install:
    https://docs.microsoft.com/en-us/powershell/module/addsadministration/?view=win10-ps
#>

#These variables are the ones that are most likely to be updated
$daysbeforeexpiretonotify = 7
$Email = #Add email here
$Email_Password = Get-Content #Path to password SecureString Hex File here
$SMTPServer = #Your SMTP server address here
$Email_Subject = "Your Active Directory password will expire in $($user.DaysToExpiry) days"
$Email_Body = “Your password will expire at $($user.ExpiryDate) (EST).”

#No Changes Needed Below Here
$Credential = New-Object -TypeName PSCredential -ArgumentList $Email, ($Email_Password | ConvertTo-SecureString)  
$now = (get-date).ToFileTime()  
$threshold = (get-date).adddays($daysbeforeexpiretonotify).ToFileTime()  
$users = Get-ADUser -filter { Enabled -eq $True -and PasswordNeverExpires -eq $False } –Properties "msDS-UserPasswordExpiryTimeComputed",mail -searchbase "OU=Company,DC=company,DC=local" |   
   where { $_."msDS-UserPasswordExpiryTimeComputed" -lt $threshold -and $_."msDS-UserPasswordExpiryTimeComputed" -gt $now }
   Select-Object "Name",  
                 "Mail",  
                 @{Name="ExpireDate";Expression={  
                       [datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")  
                       }  
                 },  
                 @{Name="DaysToExpire";Expression={  
                        [int](($_."msDS-UserPasswordExpiryTimeComputed" - $now) / 864000000000)  
                        }  
                 } |  
    sort-object name  
  
foreach ($user in $users) {
    Send-MailMessage -To $user.mail -From $Email  -Subject $Email_Subject -Body $Email_Body -Credential ($Credential) -SmtpServer $SMTPServer -Port 587
}
