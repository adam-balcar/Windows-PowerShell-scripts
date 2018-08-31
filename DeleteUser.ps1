#DELETE USER

# User info
$login="novak"
# Email redirection to
$presmerovani="info@company.cz"
# AD container for removed users
$kontejner = '2018-07'


# Create credential $cred
$user = "$Env:UserName"
$passwdFile = "$env:USERPROFILE\MyScript-$user"
if ((Test-Path $passwdFile) -eq $false) {
  $cred = new-object system.management.automation.pscredential $user,
        (read-host -assecurestring -prompt "Enter $user password")
    $cred.Password | ConvertFrom-SecureString | Set-Content $passwdFile
}
else {
  $cred = new-object system.management.automation.pscredential $user,
        (Get-Content $passwdFile | ConvertTo-SecureString)
}

# Import AD module
Try 
{ 
  Import-Module ActiveDirectory -ErrorAction Stop 
} 
Catch 
{ 
  Write-Host "[ERROR] ActiveDirectory Module couldn't be loaded. Script will stop!" 
  Exit 1 
} 

# Import Exchange module
Try 
{  
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://Email/PowerShell/ -Authentication Kerberos -Credential $cred
    Import-PSSession $session -AllowClobber -ErrorAction Stop
} 
Catch 
{ 
  Write-Host "[ERROR] Exchange Module couldn't be loaded. Script will stop!" 
  Exit 1 
}

# Values formatting
$vyraz = "OU=$kontejner,OU=Vyřazené účty,OU=Firma,DC=company,DC=local"
$email = $login + "@company.cz"
$name = Get-ADUser -Identity $login | Select -Expand Name


# Create admin credential $creda
$user = "Administrator@company.local"
$passwdFile = "$env:USERPROFILE\MyScript-$user"
if ((Test-Path $passwdFile) -eq $false) {
  $creda = new-object system.management.automation.pscredential $user,
        (read-host -assecurestring -prompt "Enter Administrator password")
    $creda.Password | ConvertFrom-SecureString | Set-Content $passwdFile
}
else {
  $creda = new-object system.management.automation.pscredential $user,
        (Get-Content $passwdFile | ConvertTo-SecureString)
}

# import Azure AD module
# Create credential user@company.cz $credo
$usermail = "$Env:UserName@company.cz"
$passwdFile = "$env:USERPROFILE\MyScript-$usermail"
if ((Test-Path $passwdFile) -eq $false) {
  $credo = new-object system.management.automation.pscredential $usermail,
        (read-host -assecurestring -prompt "Enter $usermail password")
    $credo.Password | ConvertFrom-SecureString | Set-Content $passwdFile
}
else {
  $credo = new-object system.management.automation.pscredential $usermail,
        (Get-Content $passwdFile | ConvertTo-SecureString)
} 

# connect Azure AD module
Connect-MsolService -credential $credo
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $credo -Authentication Basic -AllowRedirection
Import-PSSession $session -AllowClobber


#Setting O365 - redirection, shared mailobox, removie licence, move in AD
Set-Mailbox -Identity $email -ForwardingAddress $presmerovani
Set-Mailbox -Identity $email -Type shared 

Write "300s pauza for sync"
Start-Sleep -s 300

Set-MsolUserLicense -UserPrincipalName $email -RemoveLicenses "dascz:STANDARDPACK" 

$cestausr = (get-aduser -f {mailNickname -eq $login}).DistinguishedName
Move-ADObject -Identity $cestausr -TargetPath $vyraz


$groups = (Get-ADPrincipalGroupMembership $login).name
Foreach ($group in $groups) {
if($group -ne "Domain Users"){remove-adgroupmember -Identity $group -Member $login -Confirm:$false }
}

Disable-ADAccount -Identity $login

$zobrjm = 'Ω '+$name
$ajdent = "company.local/Firma/Vyřazené účty/$kontejner/$name"


Set-RemoteMailbox -DisplayName $zobrjm -HiddenFromAddressListsEnabled $true -Identity $ajdent

write-host "Finished."
