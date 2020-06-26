# Office 365 powershell check user

## Introduccion

El script permite el chequeo de una lista de usuarios en formato CSV contra Office 365.

## Ejecucion

### Parametros

#### Parametro: -CSV

Se puede incluir el nombre del archivo CSV en formato [nombredelarchivo].csv. En caso de omision, el nombre por defecto es UPN.csv


#### Pre-condicion

Debe de existir el archivo CSV en la misma ruta en la que se encuentra el script.
* Debe utilizar el nombre por defecto UPN.csv o
* Pasarse como parametro en el script.

###### Powershell

`.\PS_O365_CheckUser.ps1`

`.\PS_O365_CheckUser.ps1 -CSV NombreEjemplo.csv`

## Output

El script emitira 2 archivos CSV en la misma ruta del script, uno con las lista de los usuarios que se encontraron en __Office 365__ (USRsFound.csv) y otra con los usuarios que NO se encontraron en __Office 365__ (USRsNotFound.csv)

## Help

Para tener mas informacion sobre el contenido del script se puede ver su informcion mediante:

###### Powershell

`Get-Help .\PS_O365_CheckUser.ps1`