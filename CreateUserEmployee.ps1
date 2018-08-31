#USER creation LDAP AD 


#nastaveni zakladnich promennych
$jmeno = "Marcela"
$prijmeni = "Nováková"
$dlouhylogin = 0 #0= login je prijmeni. 1=login je prijmeni.jmeno

#Vzorový user pro kopii a nastavení práv na extranetu.
$vzorovyuser = "novotna" #POZOR MALA PISMENA
$sef="novak"
#pokud ma podrizene je treba zmenit MANAGER u nich!

$titul = "" #muze byt prazdne
$titulzj = "" #titul za jmenem - muze byt prazdne
<#BLBNE OTHERATRIBUTES, TAKZE JSOU OBA TITULY ZAREMOVANY V AD CREATE#>

$cesta = 'OU=Administrativa,OU=Úsek likvidace škod,OU=Centrála,OU=Zaměstnanci,OU=Firma,DC=company,DC=local' #oddeleni format AD

$pozice = "Asistentka administrativy" #pozice
$oddeleni = "Klientský servis" #oddělení

$jepravnik = "0" #0 ne,  1 ano - Je pravnik?

$skupina1 = 'Web' #muze byt prazdna
$skupina2 = 'OneDrive' #muze byt prazdna
$skupina3 = '' #muze byt prazdna
$skupina4 = '' #muze byt prazdna
$skupina5 = 'Asistentka administrativy' #muze byt prazdna
$skupina6 = '' #muze byt prazdna
$heslo = "Xcvd4568" #musi splnovat domenove pozadavky
#nastaveni promennych pro sql dotaz
$cisloagenta = "000000" #cisloagenta cislo ala string
$pohlavi = 2 #1=muž, 2=žena

$mesto = 'Praha 4' #mesto
$kraj = 'Praha' #okres
$smerovacka = '140 00' 
$ulice = 'BB Centrum, budova βeta, Vyskočilova 1481/4' #adresa
$telefon = '+420 267 990 844' #telefon
$share='\\file\Soukrome\%UserName%'




#vytvoreni credentialu user $cred
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

#Import AD modulu 
Try 
{ 
  Import-Module ActiveDirectory -ErrorAction Stop 
} 
Catch 
{ 
  Write-Host "[ERROR] ActiveDirectory Module couldn't be loaded. Script will stop!" 
  Exit 1 
} 

#Import Exchange modulu
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

#fce na odstraneni diakritiky
function Remove-StringLatinCharacters
{
    PARAM ([string]$String)
    [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
}




#formatovani promennych

If($titul){
$personaltitle = $titul
$titul = $titul.Insert(0,", ")}

If($titulzj){
$generationqualifier = $titulzj
$titulzj = $titulzj.Insert(0,", ")}

if($dlouhylogin -like "1"){
$login = ((Remove-StringLatinCharacters $prijmeni).ToLower() +"."+ (Remove-StringLatinCharacters $jmeno).ToLower())}
else{$login = ((Remove-StringLatinCharacters $prijmeni).ToLower())}
$zobrazovane = $prijmeni + " " + $jmeno + $titul + $titulzj
$email = $login + "@company.cz"

$share = '\\File\Soukrome\'+$login

$eheslo = ConvertTo-SecureString $heslo -AsPlainText -Force

$manager = Get-ADUser -Identity $sef | Select -Expand DistinguishedName

<#
Write $login
Write $zobrazovane
Write $email
Write $cesta
Write $share
#>




#vytvoreni uzivatele v AD
New-ADUser  -SamAccountName $login -Name $zobrazovane -UserPrincipalName $email -GivenName $jmeno -Surname $prijmeni -DisplayName $zobrazovane -Path $cesta -AccountPassword $eheslo -Enabled 1 -OtherAttributes @{<#'personalTitle'=$personaltitle;<#'generationQualifier'=$generationqualifier;#>'manager'=$manager;'msExchUserAccountControl'=0;'showInAddressBook'="CN=Default Global Address List,CN=All Global Address Lists,CN=Address Lists Container,CN=company,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=company,DC=local","CN=Všichni uživatele,CN=All Address Lists,CN=Address Lists Container,CN=company,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=company,DC=local";'c'="CZ";'co'="Česká republika";'company'="D.A.S. Rechtsschutz AG";'department'=$oddeleni;'l'=$mesto;'st'=$kraj;'streetAddress'=$ulice;'title'=$pozice;'telephoneNumber'=$telefon;'scriptPath'="Logon.bat";'homeDrive'="U:";'homeDirectory'=$share}
write-host "User created in AD"

#pridani uzivatele do skupin
If($skupina1){Add-ADGroupMember $skupina1 -Members $login}
If($skupina2){Add-ADGroupMember $skupina2 -Members $login}
If($skupina3){Add-ADGroupMember $skupina3 -Members $login}
If($skupina4){Add-ADGroupMember $skupina4 -Members $login}
If($skupina5){Add-ADGroupMember $skupina5 -Members $login}
If($skupina6){Add-ADGroupMember $skupina6 -Members $login}
write-host "User added to groups"

write-host "Pausing 10 Seconds for AD Changes"
Start-Sleep -s 10




#povoleni mailove schranky
Enable-Mailbox  -Alias $login -Identity $login
write-host "User mailbox created on Exchange"

write-host "Pausing 180 Seconds for AD Changes"
Start-Sleep -s 180




#vytvoreni admin $creda
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

#spusteni AD sync s admin $creda
$session = new-pssession -computername DC2 -credential $creda
Invoke-Command -session $session -script { Import-Module ADSync }
Invoke-Command -session $session -script { Start-ADSyncSyncCycle -PolicyType Delta }
write-host "ADSync running.."

write-host "Pausing 180 Seconds for AD Changes"
Start-Sleep -s 180



#import Azure modulu
#vytvoreni credentialu user@company.cz $credo
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
#pripojeni Azure modulu
Connect-MsolService -credential $credo
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $credo -Authentication Basic -AllowRedirection
Import-PSSession $session -AllowClobber




#nastavi location a prideli licenci E1
Set-MsolUser -UserPrincipalName $email -UsageLocation "CZ"
Set-MsolUserLicense -UserPrincipalName $email -AddLicenses "companycz:STANDARDPACK" 
write-host "365 License E1 added"

Start-Sleep -s 60 

#vytvori novou davku migrace do 365
New-MoveRequest -Identity $email -Remote -RemoteHostName mail.company.cz -TargetDeliveryDomain companycz.mail.onmicrosoft.com -RemoteCredential $cred -BatchName 'Skript'



write-host "Migration running .. batch named Skript"
Start-Sleep -s 60 
#ceka na dokonceni
Do{ Start-Sleep -s 60 
Write-host "Waiting for migration to finish."  }Until((Get-MoveRequest -BatchName "Skript" -MoveStatus 'Completed'))
write-host "Migration finished."

#nastaveni licenci viz. mail od Standy
$LO = New-MsolLicenseOptions -AccountSkuId "companycz:STANDARDPACK" -DisabledPlans "FLOW_O365_P1","POWERAPPS_O365_P1","PROJECTWORKMANAGEMENT","INTUNE_O365"
Set-MsolUserLicense -UserPrincipalName $email -LicenseOptions $LO

#import pro praci s sql
Get-Module –ListAvailable -name SQLPS
Add-PSSnapin SqlServerCmdletSnapin100
Add-PSSnapin SqlServerProviderSnapin100

#formatovani a naplneni promennych
$tituldb = ""
If($titul){$tituldb = $titul -replace ".*, "}

$CL_Pers_Nr = $cisloagenta
$CL_Tit_1 = $tituldb
$CL_Vorname = $jmeno
$CL_FamName = $prijmeni
$FRM_Id = 627
$CA_Id = 1
$FA_Id = 2
$SEX_Id = $pohlavi
$NA_Id = 2
$STAN_Id = 2
$STT_Id = 11
$Exp_Id = 39
$KD_Login = $login
$STT_Id = 11
$R_Id = 55
$KD_PWD = "1234"

#vyplneni sql dotazu
$query = @”
--spolecne promene
/*
DECLARE	@CL_Id as int
DECLARE	@STT_Id	as	smallint

--promenne clovicek
DECLARE	@CL_Pers_Nr	as	varchar(50)
DECLARE	@CL_Tit_1	as	nvarchar(30)
DECLARE	@CL_Vorname	as	nvarchar(100)
DECLARE	@CL_FamName	as	nvarchar(100)
DECLARE	@CL_Erstkontakt	as	datetime
DECLARE	@FRM_Id	as	int
DECLARE	@CA_Id	as	smallint
DECLARE	@FA_Id	as	smallint
DECLARE	@SEX_Id	as	smallint
DECLARE	@NA_Id	as	smallint
DECLARE	@STAN_Id	as	smallint
DECLARE	@STT_Id	as smallint
DECLARE	@Exp_Id	as	int

--promenne kunden
DECLARE	@KD_Id	as	int
DECLARE	@KD_Login	as	varchar(50)
DECLARE	@R_Id	as	int
*/

DECLARE	@CL_Erstkontakt	as	datetime
SET @CL_Erstkontakt = GETDATE()

--insert clovicek
INSERT INTO [Personal_DB].[dbo].[Clovicek]
	([CL_Pers_Nr]
      ,[CL_Tit_1]
      ,[CL_Vorname]
      ,[CL_FamName]
      ,[CL_Erstkontakt]
      ,[FRM_Id]
      ,[CA_Id]
      ,[FA_Id]
      ,[SEX_Id]
      ,[NA_Id]
      ,[STAN_Id]
      ,[STT_Id]
      ,[Exp_Id])
VALUES
     	('$CL_Pers_Nr'
      ,'$CL_Tit_1'
      ,'$CL_Vorname'
      ,'$CL_FamName'
      ,@CL_Erstkontakt
      ,$FRM_Id
      ,$CA_Id
      ,$FA_Id
      ,$SEX_Id
      ,$NA_Id
      ,$STAN_Id
      ,$STT_Id
      ,$Exp_Id) 

--propojeni tabulek pres CL_Id
DECLARE @zjistene_CL_Id as int
SET @zjistene_CL_Id = (SELECT MAX(CL_Id) FROM [Personal_DB].[dbo].[Clovicek])
      
--insert kunden
INSERT INTO [Personal_DB].[dbo].[Kunden]
      ([KD_Login]
      ,[STT_Id]
      ,[CL_Id]
      ,[R_Id]
      ,[KD_PWD])
VALUES
      ('$KD_Login'
      ,$STT_Id
      ,@zjistene_CL_Id
      ,$R_Id
      ,'$KD_PWD') 
“@


#spusteni sql dotazu/insertu
Invoke-Sqlcmd -ServerInstance SQL -Query $query
write-host "Zapsáno do PDB."

#vyplneni sql dotazu extranet
$queryextranet = @”
INSERT INTO [company].[dbo].[PRIST_PRAVA]
      ([Userid]
      ,[Jmeno]
      ,[Expoz]
      ,[m1]
      ,[m2]
      ,[m3]
      ,[m4]
      ,[m5]
      ,[m6]
      ,[m7]
      ,[m8]
      ,[m9]
      ,[m10])

SELECT 
      subsel1.KD_Id,
      subsel1.KD_Login,
      subsel1.Expoz,

      subsel2.m1, 
      subsel2.m2,
      subsel2.m3,
      subsel2.m4,
      subsel2.m5,
      subsel2.m6,
      subsel2.m7,
      subsel2.m8,
      subsel2.m9,
      subsel2.m10

FROM
      (
            SELECT TOP 1
                  K.[KD_Id],
                  K.KD_Login,
                  0 As Expoz
            FROM [Personal_DB].[dbo].[Kunden] K 
          
            INNER JOIN [Personal_DB].[dbo].[Clovicek] C 
                  ON K.[CL_Id]=C.[CL_Id] 
          
            LEFT OUTER JOIN company..Cis_Expozitury E 
                  ON E.ID=C.Exp_Id
          
            WHERE K.STT_Id=11 And C.STT_Id=11
              and k.KD_Login = '$login'

      ) subsel1
      
INNER JOIN
      (
            SELECT TOP 1
                  H.[m1], H.[m2], H.[m3], H.[m4], H.[m5], H.[m6], H.[m7], H.[m8], H.[m9], H.[m10]
            FROM [company].[dbo].[PRIST_PRAVA] H
            WHERE Jmeno like '$vzorovyuser'

      ) subsel2
      on 1=1

“@


#spusteni sql dotazu/insertu
Invoke-Sqlcmd -ServerInstance SQL -Query $queryextranet
write-host "Práva Extranet nastavena."

if($jepravnik){
$companyadquery = @”
--insert companyAD
INSERT INTO [companyAD].[dbo].[Pravnici]
	([JmenoPrijmeni]
      ,[PravnikLogin])
VALUES
     	('$zobrazovane'
      ,'$login') 
“@

#spusteni sql dotazu/insertu
Invoke-Sqlcmd -ServerInstance SQL -Query $companyadquery
write-host "companyAD SQL nastaveno."
}

write-host "Pausing 250 Seconds for AD Changes"
Start-Sleep -s 250



#doplnek 
Set-MailboxRegionalConfiguration -Identity $email -DateFormat "d. M. yyyy" -TimeFormat "H:mm" -Language "cs-CZ" -TimeZone "Central Europe Standard Time"
Set-MailboxRegionalConfiguration  -Identity $email -LocalizeDefaultFolderName
$hodnotka = $login + ":\Kalendář"
Add-MailBoxFolderPermission $hodnotka –User Firma –AccessRights Reviewer 


write-host "Kaněc filma."