
# Import-Module AD + SQL 
Import-Module ActiveDirectory, SqlServer

# Enter user login here
$login="jan.novak"

# Random 8 characters password generator, with user friendly vowels and consonants.
$h1 = Get-Random -InputObject "A","E","I","O","U"
$h2 = Get-Random -InputObject "b","c","d","f","g","h","j","k","l","m","n","p","q","r","s","t","v","w","x"
$h3 = Get-Random -InputObject "a","e","i","o","u"
$h4 = Get-Random -InputObject "b","c","d","f","g","h","j","k","l","m","n","p","q","r","s","t","v","w","x"
$rnd=Get-Random -Minimum 1000 -Maximum 10000
$heslo=$h1+$h2+$h3+$h4+$rnd


$telefon = (get-aduser $login -Properties mobile).mobile -replace (' ') # Cell number from AD
if (!$telefon) {$telefon = (get-aduser $login -Properties telephonenumber).telephonenumber -replace (' ')} # No cell = hard line number
if (!$telefon) {$telefon = $(Read-Host "Zadej mobilní číslo pro zaslání hesla")} # No number = enter manually
#Write $telefon

# AD > Set new password, unlock account.
Set-ADAccountPassword $login -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $heslo -Force)
Unlock-ADAccount -Identity $login

# Text message body
$sms = "Dobrý den. Vaše nové heslo pro přístup je: $heslo"

# Query for sms gateway
$query = @”

INSERT INTO [XXX].[dbo].[Messages]
      (Sender, Recipient, Body)
VALUES
      ('+420777666555', '$telefon', '$sms')
“@

# Run query.
Invoke-Sqlcmd -ServerInstance SQL -Query $query

# Info msg
write-host $login " / " $heslo " - SMS odeslána na telefon " $telefon
