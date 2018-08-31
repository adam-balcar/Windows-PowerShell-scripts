# Create Phonebook from LDAP AD and upload to PBX

#import pro praci s sql
Get-Module –ListAvailable -name SQLPS
Add-PSSnapin SqlServerCmdletSnapin100
Add-PSSnapin SqlServerProviderSnapin100

#sql dotaz
$query = @”
(
SELECT

 (CASE 
			WHEN tblADSI.department like '%Budějovice%' THEN REPLACE((cast(SUBSTRING(tblADSI.department,12,20) as varchar(max))collate SQL_Latin1_General_Cp1251_CS_AS),'Ceske ','') + ' - '
			WHEN tblADSI.department like '%Králové%' THEN REPLACE((cast(SUBSTRING(tblADSI.department,12,20) as varchar(max))collate SQL_Latin1_General_Cp1251_CS_AS),' Kralove','') + ' - '
			WHEN tblADSI.department like '%Labem%' THEN REPLACE((cast(SUBSTRING(tblADSI.department,12,20) as varchar(max))collate SQL_Latin1_General_Cp1251_CS_AS),' nad Labem','') + ' - '
			WHEN tblADSI.department like 'Expozitura%' THEN (cast(SUBSTRING(tblADSI.department,12,20) as varchar(max))collate SQL_Latin1_General_Cp1251_CS_AS) + ' - '
			ELSE ''
			 END) +
UPPER(
	LEFT(tblADSI.sAMAccountName,1))
		 + SUBSTRING(
			(CASE WHEN CHARINDEX('.',tblADSI.sAMAccountName)>0 THEN LEFT(tblADSI.sAMAccountName,CHARINDEX('.',tblADSI.sAMAccountName)-1) + ' '
				+ UPPER(
						LEFT(RIGHT(tblADSI.sAMAccountName, LEN(tblADSI.sAMAccountName) - CHARINDEX('.',tblADSI.sAMAccountName)),1))
							+ SUBSTRING(RIGHT(tblADSI.sAMAccountName, LEN(tblADSI.sAMAccountName) - CHARINDEX('.',tblADSI.sAMAccountName)),2,LEN(tblADSI.sAMAccountName))
					ELSE tblADSI.sAMAccountName END )
			,2,LEN(tblADSI.sAMAccountName)
			) as Name,
			
		


 (CASE WHEN tblADSI.telephoneNumber = '+420 666 999 888' OR
                      tblADSI.telephoneNumber NOT LIKE '+420 666 999 ___' THEN tblADSI.homePhone ELSE RIGHT(tblADSI.telephoneNumber, 3) END) AS Telephone
FROM         OPENQUERY(ADSI, 
                      'SELECT  sAMAccountName, homePhone, telephoneNumber, displayName, department
  FROM  ''LDAP://company.local/OU=Zaměstnanci,OU=Firma,DC=company,DC=local'' 
  WHERE objectClass =  ''User'' 
  ') 
                      AS tblADSI
WHERE tblADSI.telephoneNumber IS NOT NULL                     


UNION

SELECT
UPPER(
	LEFT(tblADSI.sAMAccountName,1))
		 + SUBSTRING(
			(CASE WHEN CHARINDEX('.',tblADSI.sAMAccountName)>0 THEN LEFT(tblADSI.sAMAccountName,CHARINDEX('.',tblADSI.sAMAccountName)-1) + ' '
				+ UPPER(
						LEFT(RIGHT(tblADSI.sAMAccountName, LEN(tblADSI.sAMAccountName) - CHARINDEX('.',tblADSI.sAMAccountName)),1))
							+ SUBSTRING(RIGHT(tblADSI.sAMAccountName, LEN(tblADSI.sAMAccountName) - CHARINDEX('.',tblADSI.sAMAccountName)),2,LEN(tblADSI.sAMAccountName))
					ELSE tblADSI.sAMAccountName END )
			,2,LEN(tblADSI.sAMAccountName)
			) as Name,

REPLACE(('0'+(RIGHT(tblADSI.mobile, 11))),' ','') AS Telephone
FROM         OPENQUERY(ADSI, 
                      'SELECT  sAMAccountName, mobile, displayName 
  FROM  ''LDAP://company.local/OU=Zaměstnanci,OU=Firma,DC=company,DC=local'' 
  WHERE objectClass =  ''User'' 
  ') 
                      AS tblADSI
WHERE tblADSI.mobile IS NOT NULL                      
)ORDER BY name

FOR XML RAW ('DirectoryEntry'), ROOT ('YealinkIPPhoneDirectory'), ELEMENTS XSINIL; 
“@

#sql export do xml
$DBState = (Invoke-Sqlcmd -ServerInstance SQL -Query $query  ) 
$DBState = $DBState.ItemArray[0..$DBState.Length] 
$DBState | Out-File ("$env:APPDATA\contact.xml") -NoNewline

#ftp upload
$server = "ustredna.company.local"
$filelist = "$env:APPDATA\contact.xml"   

$user = 'company'
$password = 'secret'
$dir = '/var/www/html'

"open $server
user $user $password
binary  
cd $dir     
" +
($filelist.split(' ') | %{ "put ""$_""`n" }) | ftp -i -in
