(
SELECT

 (CASE 
			WHEN tblADSI.department like '%Bud�jovice%' THEN REPLACE((cast(SUBSTRING(tblADSI.department,12,20) as varchar(max))collate SQL_Latin1_General_Cp1251_CS_AS),'Ceske ','') + ' - '
			WHEN tblADSI.department like '%Kr�lov�%' THEN REPLACE((cast(SUBSTRING(tblADSI.department,12,20) as varchar(max))collate SQL_Latin1_General_Cp1251_CS_AS),' Kralove','') + ' - '
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
			
		


 (CASE WHEN tblADSI.telephoneNumber = '+420 267 990 711' OR
                      tblADSI.telephoneNumber NOT LIKE '+420 267 990 ___' THEN tblADSI.homePhone ELSE RIGHT(tblADSI.telephoneNumber, 3) END) AS Telephone
FROM         OPENQUERY(ADSI, 
                      'SELECT  sAMAccountName, homePhone, telephoneNumber, displayName, department
  FROM  ''LDAP://das.local/OU=Zam�stnanci,OU=Firma,DC=das,DC=local'' 
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
  FROM  ''LDAP://das.local/OU=Zam�stnanci,OU=Firma,DC=das,DC=local'' 
  WHERE objectClass =  ''User'' 
  ') 
                      AS tblADSI
WHERE tblADSI.mobile IS NOT NULL                      
)ORDER BY name

--FOR XML RAW ('DirectoryEntry'), ROOT ('YealinkIPPhoneDirectory'), ELEMENTS XSINIL; 