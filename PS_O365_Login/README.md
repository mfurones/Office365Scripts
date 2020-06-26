# Office 365 powershell login

## Introduccion

Este repositorio contiene un script para el login a Office 365.
la forma de trabaja es mediante el pasaje de parametros o por medio de la consola interactiva.

## Ejecucion

Se puede ejecutar mediante el uso de parametros.

###### Powershell

`.\PS_O365_Login.ps1 -user 'unUsuario' -password 'C0ntr4s3n4'`

Tambien se puede ejecutar sin el uso de parametros, de esta forma la consola nos pedira de forma interactiva todos los datos necesarios para completar la tarea.

###### Powershell

`.\PS_O365_Login.ps1`

## Help

Para tener mas informacion sobre el contenido del script se puede ver su informcion mediante:

###### Powershell

`Get-Help .\PS_O365_Login.ps1`