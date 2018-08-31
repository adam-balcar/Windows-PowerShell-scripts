#import pro praci s sql
Get-Module –ListAvailable -name SQLPS
Add-PSSnapin SqlServerCmdletSnapin100
Add-PSSnapin SqlServerProviderSnapin100


$query = @”
 SELECT COUNT(*)
  FROM [Audit].[dbo].[Inventory] 
WHERE [servicetag] like 'unavailable'
“@
$count = Invoke-Sqlcmd -ServerInstance SQL -Query $query | Format-List
$count = (Out-String -InputObject $count -Width 100)
$count = $count.Replace('Column1 : ','')
$count = $count.Trim()
$count = [convert]::ToInt32($count, 10)

Write 'Count is'$count

For ($i=1; $i -le $count; $i++) {
Write 'Round'$i

$query = @”
 SELECT TOP 1 [computername]
  FROM [Audit].[dbo].[Inventory] 
WHERE [servicetag] like 'unavailable'
order by [updated] asc 
“@
#spusteni sql dotazu/insertu
$item = Invoke-Sqlcmd -ServerInstance SQL -Query $query | Format-List
$item = (Out-String -InputObject $item -Width 100)
$item = $item.Replace('computername : ','')
$item = $item.Trim()

Write 'Item'$item
    if (test-connection -computername $item -count 1 -quiet){ 
 
#formatovani a naplneni promennych
$servicetag = (Get-WmiObject -ComputerName $item win32_SystemEnclosure  | select serialnumber | Format-List ) 
$servicetag = (Out-String -InputObject $servicetag -Width 100)
$servicetag = $servicetag.Replace('serialnumber : ','')
if ($servicetag.length -gt 11) {$servicetag =  $servicetag.substring(0, 11)} 
$servicetag = $servicetag.Replace('None','')
$servicetag = $servicetag.Trim()
if($servicetag -eq ""){$servicetag = "virtual"}

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
--update 
UPDATE [Audit].[dbo].[Inventory]
SET	   [servicetag] = '$servicetag'
      ,[primaryuser] = '$username'
      ,[officelicense] = '$officever'
      ,[updated] = getdate()

WHERE [computername] like '$item' 
“@

#spusteni sql dotazu/insertu
Invoke-Sqlcmd -ServerInstance SQL -Query $query

    }
    Else {
     #vyplneni sql dotazu
$query = @”
UPDATE [Audit].[dbo].[Inventory]
SET	   [updated] = getdate()

WHERE [computername] like '$item' 
“@
#spusteni sql dotazu/insertu
Invoke-Sqlcmd -ServerInstance SQL -Query $query

     } 
     
}