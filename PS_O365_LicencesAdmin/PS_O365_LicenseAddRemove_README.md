# Add & Remove licencias de Office 365

## Introduccion

Este script permite agregar y remover licencias en Office 365 por powershell.

## Ejecucion

### Parametros

El script puede ser ejecutado sin parametros (a excepcion de los Mandatorios) y la consola solicitara la informacion necesaria para su ejecucion.

#### Parametro: -tipoUsuario (Mandatorio)

Admite 2 tipos de valores:
 - O365 --> Este valor busca todos los usuarios disponibles en Office 365
 - [NombreEjemplo].csv --> Agregando el nombre de un archivo en CSV que contenga el listado de UPNs de los usuarios, permite efectuar acciones sobre los usuarios especificados.

#### Parametro: -accion

Admite 2 tipos de valores:
 - Add --> Permite agregar una licencia a cada usuario.
 - Remove --> Permite remover una licencia a cada usuario.

### Ejemplos

###### Powershell

`.\PS_O365_LicenseAddRemove.ps1 -tipoUsuario O365 -accion Add -location AR -AccountSkuId enterprise:ENTERPRISEPACK -servicios "EXCHANGE_S_ENTERPRISE","YAMMER_ENTERPRISE"`

`.\PS_O365_LicenseAddRemove.ps1 -tipoUsuario UPN.csv -accion Add`

`.\PS_O365_LicenseAddRemove.ps1 -tipoUsuario O365 -accion Remove`

`.\PS_O365_LicenseAddRemove.ps1 -tipoUsuario UPN.csv -accion Remove -AccountSkuId enterprise:ENTERPRISEPACK`

## Help

Para tener mas informacion sobre el contenido del script se puede ver su informcion mediante:

###### Powershell

`Get-Help .\PS_O365_LicenseAddRemove.ps1`