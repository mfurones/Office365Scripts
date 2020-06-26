# Reporte de Licencias y Servicios de usuarios de Office365

## Introduccion

Este script permite generar un reporte de licencias y servicios en Office 365 por powershell.

## Ejecucion

### Parametros

Segun el tipo de parametros se obtiene un tipo de reporte.

#### Parametro: -tipoUsuario

Admite 2 tipos de valores:
 - O365 --> Este valor busca todos los usuarios disponibles en Office 365
 - [NombreEjemplo].csv --> Agregando el nombre de un archivo en CSV que contenga el listado de UPNs de los usuarios, permite exportar un reporte solo para dichos usuarios.

#### Parametro: -tipoReporte

Admite 2 tipos de valores:
 - License --> Exporta un reporte con todas las licencias activas o no de cada usuario.
 - Service --> Exporta un reporte con todos los servicios (y sus estados) de cada licencia por cada usuario.

### Ejemplos

###### Powershell

`.\PS_O365_License_Report.ps1 -tipoUsuario O365 -tipoReporte Licencias`

`.\PS_O365_License_Report.ps1 -tipoUsuario UPN.csv -tipoReporte Licencias`

`.\PS_O365_License_Report.ps1 -tipoUsuario O365 -tipoReporte Servicios`

`.\PS_O365_License_Report.ps1 -tipoUsuario UPN.csv -tipoReporte Servicios`

## Help

Para tener mas informacion sobre el contenido del script se puede ver su informcion mediante:

###### Powershell

`Get-Help .\PS_O365_License_Report.ps1`