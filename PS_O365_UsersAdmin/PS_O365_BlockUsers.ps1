<#
 .SYNOPSIS
    Allow to lock or unlock users in Office 365 by powershell.

 .DESCRIPTION
    This script allows to lock or unlock users in Office 365 by powershell.

 .PRE-REQUISITES
    * Have the CSV file with the UPN complete in the same path than the script.
    * OR have the UPN from the user.

 .PARAMETER
 .-CSV
    Defines the CSV file name.
    Variable type [string].
    Example: NombreEjemplo.csv

 .-UPN
    Defines the User Principal Name.
    Variable type [string].
    Example: usuario@dominio.com

 .-block
    Defines if lock or unlock the users.
    Variable type [string].
    Example: True or False

 .-v or -Verbose
    Sets Verbose mode ON.

.EXAMPLE

    .\PS_O365_BlockUsers.ps1 -upn usuario@dominio.com -block True

    .\PS_O365_BlockUsers.ps1 -CSV NombreEjemplo.csv -block False

#>

[CmdLetBinding()]
param(

 [string] $CSV,
 [string] $upn,
 [string] $block

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


<# Error de parametros #>

function fErrorParameter
{
    param(
        [string] $texto
    )
    process
    {
        Write-host "$(fFormatDate) | $($texto)" -ForegroundColor Red
        Write-host "$(fFormatDate) | fin del script"
        exit
    }
}


<#
******************************************************************************
Script
******************************************************************************
#>

$currentTime = (Get-Date)

if($block)
{
    if($block.ToLower() -eq "true")
    {
        $blocked = $true
    }
    elseif($block.ToLower() -eq "false")
    {
        $blocked = $false
    }
    else
    {
        fErrorParameter -texto "Definición incorrecta. Las opciones de -block son: True o False"
    }
}
else
{
    fErrorParameter -texto "Falta definir el parametro -block"
}

if ($CSV -and !$upn)
{
    #Importacion del CSV
    try
    {
        $ErrorActionPreference = "Stop"
        $UPNs = Import-Csv $CSV
        Write-Verbose "$(fFormatDate) | Importando CSV..."
    }
    catch
    {
        fErrorParameter -texto "Error con la importacion del archivo CSV. Verificar su formato o nombre."
    }
    finally
    {
       $ErrorActionPreference = "Continue"
    }

    $UPNs | ForEach-Object {
        Write-Verbose "$(fFormatDate) | Usuario: $($_.UserPrincipalName) | Bloqueado: $($blocked)"
        Set-MsolUser -UserPrincipalName $_.UserPrincipalName –blockcredential $blocked
    }
}
elseif (!$CSV -and $upn)
{
    Write-Verbose "$(fFormatDate) | Usuario: $($_.UserPrincipalName) | Bloqueado: $($blocked)"
    Set-MsolUser -UserPrincipalName $upn –blockcredential $blocked
}
elseif (!$CSV -and !$upn)
{
    fErrorParameter -texto "Falta definir uno de los 2 parametros: -CSV o -upn"
}
elseif ($CSV -and $upn)
{
    fErrorParameter -texto "Esta sobre definido, debe elegir uno de los 2 parametros: -CSV o -upn"
}
else
{
    fErrorParameter -texto "Se produjo un error inesperado"
}

Write-Verbose "$(fFormatDate) | Fin del script. (Duración: $(("{0:hh\:mm\:ss}" -f ((Get-Date) - $currentTime))))"

