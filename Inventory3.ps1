#import pro praci s sql
Get-Module –ListAvailable -name SQLPS
Add-PSSnapin SqlServerCmdletSnapin100
Add-PSSnapin SqlServerProviderSnapin100


$query = @”
 SELECT COUNT(*)
  FROM [Temp2].[dbo].[Inventory] 

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
  FROM [Temp2].[dbo].[Inventory] 
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

$CPUInfo = Get-WmiObject Win32_Processor -ComputerName $item #Get CPU Information 
$OSInfo = Get-WmiObject Win32_OperatingSystem -ComputerName $item #Get OS Information 
$driveinfo = Get-WmiObject -Class MSFT_PhysicalDisk -computername $item -Namespace root\Microsoft\Windows\Storage | 
     Select       @{
         name="Size";
         expression={
                
         [string]::Concat([math]::round(($_.Size) / 1073741824),'GB')} # to cislo znamena 1024^3 aby odpovidala velikost
         },
     @{
         name="MediaType"; 
         expression={
             switch ($_.MediaType) { 
                 3 {"HDD"}
                 4 {"SSD"} 
             } 
         }
     } | Format-List

#Get Memory Information. The data will be shown in a table as MB, rounded to the nearest second decimal. 
$OSTotalVirtualMemory = [math]::round($OSInfo.TotalVirtualMemorySize / 1MB, 2) 
$OSTotalVisibleMemory = [math]::round(($OSInfo.TotalVisibleMemorySize  / 1MB), 2) 
$PhysicalMemory = Get-WmiObject CIM_PhysicalMemory -ComputerName $item | Measure-Object -Property capacity -sum | % {[math]::round(($_.sum / 1GB),2)}
$PhysicalMemory = $PhysicalMemory.ToString()+'GB'
$driveinfo = (Out-String -InputObject $driveinfo -Width 100)
$driveinfo = $driveinfo.Replace('Size      : ',', ')
$driveinfo = $driveinfo.Replace('MediaType : ','-')
$driveinfo = $driveinfo -replace "`r|`n",""
$driveinfo = $driveinfo.Trim(", ")
$cpu = $CPUInfo.Name + ", "+$CPUInfo.NumberOfCores + " cores"

$infoObject = New-Object PSObject 
#The following add data to the infoObjects. 
#Add-Member -inputObject $infoObject -memberType NoteProperty -name "Name" -value $CPUInfo.SystemName 
Add-Member -inputObject $infoObject -memberType NoteProperty -name "CPU" -value $cpu

Add-Member -inputObject $infoObject -memberType NoteProperty -name "RAM" -value $PhysicalMemory 
Add-Member -inputObject $infoObject -memberType NoteProperty -name "DRIVE" -value $driveinfo 

#vyplneni sql dotazu
$query = @”
--update 
INSERT INTO [Temp2].[dbo].[Inventory3] ([servicetag],[cpu],[ram],[drive],[updated])
VALUES('$servicetag','$cpu','$PhysicalMemory','$driveinfo',getdate())
“@

#spusteni sql dotazu/insertu
Invoke-Sqlcmd -ServerInstance SQL -Query $query
     #vyplneni sql dotazu
$query = @”
UPDATE [Temp2].[dbo].[Inventory]
SET	   [updated] = getdate()

WHERE [computername] like '$item' 
“@
#spusteni sql dotazu/insertu
Invoke-Sqlcmd -ServerInstance SQL -Query $query
    }
    Else {


$query = @”
 SELECT TOP 1 [servicetag]
  FROM [Temp2].[dbo].[Inventory] 
order by [updated] asc 
“@
#spusteni sql dotazu/insertu
$stg = Invoke-Sqlcmd -ServerInstance SQL -Query $query | Format-List
$stg = (Out-String -InputObject $stg -Width 100)
$stg = $stg.Replace('servicetag : ','')
$stg = $stg.Trim()

#vyplneni sql dotazu
$query = @”
--update 
INSERT INTO [Temp2].[dbo].[Inventory3] ([servicetag],[updated])
VALUES('$stg',getdate())
“@

#spusteni sql dotazu/insertu
Invoke-Sqlcmd -ServerInstance SQL -Query $query

     #vyplneni sql dotazu
$query = @”
UPDATE [Temp2].[dbo].[Inventory]
SET	   [updated] = getdate()

WHERE [computername] like '$item' 
“@
#spusteni sql dotazu/insertu
Invoke-Sqlcmd -ServerInstance SQL -Query $query

     } 
     
}