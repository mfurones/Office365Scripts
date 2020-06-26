# Bloqueo y desbloqueo usuarios Office365 powershell

## Introduccion

El script permite el bloqueo o desbloqueo de uno o varios usuarios de Office 365 por powershell.

## Ejecucion

### Bloqueo o Desbloqueo de un usuario

Para el bloqueo o desbloqueo del usuario hay que pasar el UPN y su nuevo estado como parametros.

###### Powershell

`.\PS_O365_BlockUsers.ps1 -upn usuario@dominio.com -block True`

`.\PS_O365_BlockUsers.ps1 -upn usuario@dominio.com -block False`

### Bloqueo o Desbloqueo de una lista de usuario

Para el bloqueo o desbloqueo de una lista de usuarios hay que pasar el archivo CSV y su nuevo estado como parametros.

###### Powershell

`.\PS_O365_BlockUsers.ps1 -CSV NombreEjemplo.csv -block True`

`.\PS_O365_BlockUsers.ps1 -CSV NombreEjemplo.csv -block False`

## Help

Para tener mas informacion sobre el contenido del script se puede ver su informcion mediante:

###### Powershell

`Get-Help .\PS_O365_BlockUsers.ps1`