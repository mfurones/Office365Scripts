# Cambio de Contraseñas para usuarios Office365 powershell

## Introduccion

El script permite el cambio de contraseña de uno o varios usuarios de Office 365 por powershell.

## Ejecucion

### Parametros

#### Parametro: -tipoUsuario

Soporta 3 opciones:

* O365: Afecta a todo el directorio de usuarios de Office 365
* [nombredelarchivo].csv: Afecta a un listado especifico de usuarios de Office 365
* usuario@dominio.com: Afecta a un usuario en particular de Office 365

##### Pre-condicion

[No Mandatorio] Debe de existir el archivo CSV en la misma ruta en la que se encuentra el script.

#### Parametro: -NewPassword

Nueva contraseña a aplicar.

#### Parametro: -ForceChangePassword

Indica si fuerza o no al usuario a modificar su contraseña luego de hacer login.

###### Powershell

`.\PS_O365_UsersPassword.ps1`

`.\PS_O365_UsersPassword.ps1 -tipoUsuario O365 -NewPassword P4ssw0rd@1 -ForceChangePassword False`

`.\PS_O365_UsersPassword.ps1 -tipoUsuario NombreEjemplo.csv -NewPassword P4ssw0rd@1 -ForceChangePassword True`

`.\PS_O365_UsersPassword.ps1 -tipoUsuario user@domain.com -NewPassword P4ssw0rd@1 -ForceChangePassword False`

## Help

Para tener mas informacion sobre el contenido del script se puede ver su informcion mediante:

###### Powershell

`Get-Help .\PS_O365_UsersPassword.ps1`