#import pro praci s sql
Get-Module –ListAvailable -name SQLPS
Add-PSSnapin SqlServerCmdletSnapin100
Add-PSSnapin SqlServerProviderSnapin100

#$base = 'OU=Centrála,OU=Počítače,dc=das,dc=local'
#$base = 'OU=Expozitury,OU=Počítače,dc=das,dc=local'
$base = 'OU=Virtualy,OU=Počítače,dc=das,dc=local'
#$base = 'OU=Sklad,OU=Počítače,dc=das,dc=local'
#$base = 'OU=Vyřazené,OU=Počítače,dc=das,dc=local'


$filter = 'ObjectClass -eq "Computer"'
$items = 'POSPISILOVA-NB2' #Get-ADComputer -SearchBase $base -Filter $filter | Select -Expand Name


Foreach ($item in $items) {
   	    
   Write 'Doing my best regarding '$item'.'
    if (test-connection -computername $item -count 1 -quiet){ 
 
#formatovani a naplneni promennych
$servicetag = (Get-WmiObject -ComputerName $item win32_SystemEnclosure  | select serialnumber | Format-List ) 

$servicetag = (Out-String -InputObject $servicetag -Width 100)

$servicetag = $servicetag.Replace('serialnumber : ','')
if ($servicetag.length -gt 11) {$servicetag =  $servicetag.substring(0, 11)} 
$servicetag = $servicetag.Replace('None','')
$servicetag = $servicetag.Trim()
if($servicetag -eq ""){$servicetag = "virtual"}

$computername = $item

$username = (Get-WmiObject -ComputerName $item Win32_ComputerSystem | Select-Object -Property UserName | Format-List )
$username = (Out-String -InputObject $username -Width 100)
$username = $username.Replace('UserName : DAS\','')
$username = $username.Replace('UserName :','')
$username = $username.Trim()
$username = $username.ToLower()
if($username -eq ""){$username = "unavailable"}

$officever = Get-WmiObject -ComputerName $item Win32_SoftwareElement | ? { $_.name -like 'Microsoft.Office.Tools.Outlook*' } | Select Version | Format-List
$officever = (Out-String -InputObject $officever -Width 100)
$officever = $officever.Replace('Version :','')
$officever = $officever.Trim()
if($officever){$officever =  $officever.substring(0, 2)}

if($officever -eq '14'){$officever = 'Office OEM 2010'}
elseif($officever -eq '15'){$officever = 'Office 2013 ProPlus'}
elseif($officever -eq '16'){$officever = 'Office 365 Enterprise E3'}
else {$officever = 'unknown'}

#vyplneni sql dotazu
$query = @”

--insert 
INSERT INTO [Audit].[dbo].[Inventory]
	([servicetag]
      ,[computername]
      ,[primaryuser]
      ,[officelicense]
      ,[updated])
VALUES
     	('$servicetag'
      ,'$computername'
      ,'$username'
      ,'$officever'
      ,getdate()) 
“@
#spusteni sql dotazu/insertu
Invoke-Sqlcmd -ServerInstance SQL -Query $query

    }
    Else {

     $computername = $item
     if($servicetag -eq ""){$servicetag = "virtual"}
     #$servicetag = "unavailable"
     $username = "unavailable"
     $officever = "unavailable"

     #vyplneni sql dotazu
$query = @”

--insert 
INSERT INTO [Audit].[dbo].[Inventory]
	([servicetag]
      ,[computername]
      ,[primaryuser]
      ,[officelicense]
      ,[updated])
VALUES
     	('$servicetag'
      ,'$computername'
      ,'$username'
      ,'$officever'
      ,getdate()) 
“@
#spusteni sql dotazu/insertu
Invoke-Sqlcmd -ServerInstance SQL -Query $query

     } #>
    } 
