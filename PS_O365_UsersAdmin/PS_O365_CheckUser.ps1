<#
 .SYNOPSIS
    Allow to check if a user in CSV is a valid user in Office 365.

 .DESCRIPTION
    This script allows to validate a list of users in a CSV file against Office 365.

 .PRE-REQUISITES
    * Have the CSV file with the UPN complete in the same path than the script. (file default: UPN.csv)

 .OUTPUT
    * 2 Files with the right and wrongs UPNs
    * In console, the result of the check. (in verbose mode)

 .PARAMETER
 .-CSV
    Defines the CSV file name.
    Variable type [string].
    Example: NombreEjemplo.csv

 .-v or -Verbose
    Sets Verbose mode ON.

.EXAMPLE

    .\PS_O365_CheckUser.ps1

    .\PS_O365_CheckUser.ps1 -CSV NombreEjemplo.csv

#>

[CmdLetBinding()]
param(

 [string] $CSV = "UPN.csv"

)


 <#
******************************************************************************
Funciones
******************************************************************************
#>

<# Formato Fecha #>

function fFormatDate
{
    param()
    process
    {
        return "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T)"
    }
}


 <#
******************************************************************************
Script
******************************************************************************
#>


$currentTime = (Get-Date)

#Importacion del CSV
try
{
    $ErrorActionPreference = "Stop"
    $UPNs = Import-Csv $CSV
    Write-Verbose "$(fFormatDate) | Importando CSV..."
}
catch
{
    Write-host "$(fFormatDate) | Error en la importación | Verificar la existencia y nombre del CSV" -ForegroundColor Red
    Write-host "$(fFormatDate) | Debe de existir el archivo: UPN.csv"
    Write-host "$(fFormatDate) | O pasar como parametro el nuevo nombre del archivo."
    Write-host "$(fFormatDate) | Ejemplo: .\PS_O365_CheckUser.ps1 -CSV NombreEjemplo.csv"
    exit
}
finally
{
   $ErrorActionPreference = "Continue"
}

$UsuarioNotFound = @()
$UsuarioFound = @()

$porCount = 0
Write-Verbose "$(fFormatDate) | Inicio del chequeo..."
# Recorro la lista de UPN
$UPNs | ForEach-Object {

    # Busco en O365 el usuario por medio del UPN
    $user = Get-MsolUser -UserPrincipalName $_.UserPrincipalName -ErrorAction SilentlyContinue
    if($user)
    {
        Write-Verbose "$(fFormatDate) | Usuario OK: $($_.UserPrincipalName)"
        $UsuarioFound += @($_.UserPrincipalName)
    }
    else
    {
        Write-Verbose "$(fFormatDate) | Usuario NOK: $($_.UserPrincipalName)"
        $UsuarioNotFound += @($_.UserPrincipalName)
    }

}

Write-Verbose "$(fFormatDate) | Fin de chequeo. (Duración: $(("{0:hh\:mm\:ss}" -f ((Get-Date) - $currentTime))))"
Write-Verbose ""
Write-Verbose "$($UsuarioFound.Count) Usuarios encontrados en O365"
Write-Verbose "$($UsuarioNotFound.Count) Usuarios NO encontrados en O365"
Write-Verbose ""
#Exportacion de resultados a CSV
try
{
    $ErrorActionPreference = "Stop"
    $UsuarioFound | Export-Csv "USRsFound.csv"
    $UsuarioNotFound | Export-Csv "USRsNotFound.csv"
    Write-Verbose "$(fFormatDate) | Se exportaron los resultados"
}
catch
{
    Write-Verbose "$(fFormatDate) | Error en la exportación"
}
finally
{
   $ErrorActionPreference = "Continue"
}

Write-Verbose "$(fFormatDate) | Fin del script. (Duración: $(("{0:hh\:mm\:ss}" -f ((Get-Date) - $currentTime))))"
