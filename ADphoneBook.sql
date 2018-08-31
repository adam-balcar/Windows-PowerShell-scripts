

SELECT displayName as Jméno,
 department as Oddìlení,
 mail as Email,
(CASE WHEN tblADSI.telephoneNumber = '+420 999 990 611' OR telephoneNumber not like '+420 999 990 ___' THEN tblADSI.homePhone ELSE RIGHT(tblADSI.telephoneNumber,3) END) as Telefon,
 RIGHT(mobile,11) as Mobil
FROM         OPENQUERY(ADSI, 'SELECT  homePhone, mobile, telephoneNumber, mail, department, displayName 
  FROM  ''LDAP://company.local/OU=Zamìstnanci,OU=Firma,DC=das,DC=local'' 
  WHERE objectClass =  ''User'' 
  ') 
                      AS tblADSI
ORDER BY department, displayName