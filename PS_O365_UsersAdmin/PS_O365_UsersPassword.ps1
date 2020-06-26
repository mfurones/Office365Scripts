<#
 .SYNOPSIS
    Allow to change a user password in Office 365 by powershell.

 .DESCRIPTION
    This script allows to change a user password in Office 365. Can be executed for a custom list, single user o for all users in Office 365.

 .PRE-REQUISITES
    Not Mandatory
    * Have the CSV file with the UPN complete in the same path than the script.

 .PARAMETERS
 .-tipoUsuario [Mandatory]
    Defines the CSV file name, single user or all O365 users.
    Variable type [string].
    Example: O365 --> for all users.
    Example: NombreEjemplo.csv --> for custom users.
    Example: user@domain.com --> for single user.

 .-NewPassword
    Defines the new password fot user/s.
    Variable type [string].
    Example: enterprise: P4ssw0rd@1

 .-ForceChangePassword
    Defines if the user is forced to change the password in next login.
    Variable type [object].
    Example: True | False

 .-v or -Verbose
    Sets Verbose mode ON.

.EXAMPLES

    .\PS_O365_UsersPassword.ps1

    .\PS_O365_UsersPassword.ps1 -tipoUsuario O365 -NewPassword P4ssw0rd@1 -ForceChangePassword False

    .\PS_O365_UsersPassword.ps1 -tipoUsuario NombreEjemplo.csv -NewPassword P4ssw0rd@1 -ForceChangePassword True

    .\PS_O365_UsersPassword.ps1 -tipoUsuario user@domain.com -NewPassword P4ssw0rd@1 -ForceChangePassword False

#>

[CmdLetBinding()]

param(
 [string] $tipoUsuario,
 [string] $NewPassword,
 [string] $ForceChangePassword

 )


 <#
******************************************************************************
Funciones
******************************************************************************
#>

function fEnumerateList
{
    param(
        [array] $array
    )
    process
    {
        $ar = $array | ConvertTo-Csv -NoTypeInformation
        for($i=0; $i -lt $ar.Count; $i++) {
            if($i) {$ar[$i] = "`"$($i)`","+ $ar[$i]}
            else {$ar[$i] = "`"Item`","+ $ar[$i]}
        }
        return $ar | ConvertFrom-Csv
    }
}


<# Formato Fecha #>

function fFormatDate
{
    param()
    process
    {
        return "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T)"
    }
}


<# Error Output #>

function fErrorParameter
{
    param(
        [string] $texto
    )
    process
    {
        Write-host "$(fFormatDate) | $($texto)" -ForegroundColor Red
        Write-host "$(fFormatDate) | fin del script"
        exit 1
    }
}


<# Buscar usuarios en Office 365 #>

function fBuscarUsuarios {
    param()
    
    $CT = (Get-Date)
    Write-Verbose "$(fFormatDate) | Buscando todos los usuarios de la cuenta de O365..."
    #Busca todos los usuarios de O365
    $users = Get-MsolUser -all
    #$users = Get-MsolUser -MaxResults 15
    Write-Verbose "$(fFormatDate) | Final de la busqueda. (Duración: $(("{0:hh\:mm\:ss}" -f ((Get-Date) - $CT))))"
    return $users
}

<# Busca usuarios en Office 365 mediante una lista previa #>

function fBuscarUsuariosLista {
    param(
        [Parameter(Mandatory=$true)]
        [object] $listado
    )

    $CT = (Get-Date)
    # Busca todos los usuarios a partir de un listado
    Write-Verbose "$(fFormatDate) | Buscando lista de usuarios en la cuenta de O365..."
    $users = $listado | ForEach-Object {
        Write-Verbose "$(fFormatDate) | Buscando usuario:  $($_.UserPrincipalName)..."
        Get-MsolUser -UserPrincipalName $_.UserPrincipalName
    }
    Write-Verbose "$(fFormatDate) | Final de la busqueda. (Duración: $(("{0:hh\:mm\:ss}" -f ((Get-Date) - $CT))))"
    return $users
}



<#
******************************************************************************
Script
******************************************************************************
#>

$currentTime = (Get-Date)
Write-Verbose "$(fFormatDate) | Start Script..."

<# Verificacion parametro -tipoUsuario #>

if (!$tipoUsuario)
{
    echo "------------------------------"
    Write-Host "Select an Users type intro."
    echo "------------------------------"
    Write-Host "1 - O365"
    Write-Host "2 - CSV File"
    Write-Host "3 - Single user"
    $pos = Read-Host "Enter N°"
    switch ($pos) 
    { 
        "1" {
            $tipoUsuario = "O365"
        } # 1
        "2" {
            echo "------------------------------"
            $tipoUsuario = Read-Host "Enter CSV file name"
        } # 2
        "3" {
            echo "------------------------------"
            $tipoUsuario = Read-Host "Enter UPN"
        } # 3
        default {fErrorParameter -texto "Opcion no determinada"}
    }
}

<# Busqueda de Usuario/s #>

if($tipoUsuario -eq "O365") #Filtra O365
{
    $usuarios = fBuscarUsuarios # Busca todos los usuarios de O365
}
elseif($tipoUsuario -match "^\w+(.csv)$") #filtra archivos .csv
{
    
    #Importacion del CSV
    try
    {
        $ErrorActionPreference = "Stop"
        $UPNs = Import-Csv $tipoUsuario
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
    
    # busqueda del usuario
    $usuarios = fBuscarUsuariosLista -listado $UPNs # Busca los usuarios mediante una lista
}
elseif($tipoUsuario -match "^\w+([-.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$") #filtra direcciones de email
{
    # Genero el objeto del usuario para ser leido como una lista importada de CSV
    $UPNs = [PSCustomObject]@{UserPrincipalName = "$($tipoUsuario)"}
    # busqueda del usuario
    $usuarios = fBuscarUsuariosLista -listado $UPNs # Busca el usuario 
}
else
{
    fErrorParameter -texto "Definición incorrecta. Las opciones de -tipoUsuario son: O365 | [nombrearchivo].csv | usuario@dominio.com"
}


<# Verificacion de parametro ForceChangePassword #>

if($ForceChangePassword)
{
    if($ForceChangePassword.ToLower() -eq "true") {$ForceChangePassword = $true}
    elseif($ForceChangePassword.ToLower() -eq "false") {$ForceChangePassword = $false}
    else {fErrorParameter -texto "Definición incorrecta. Las opciones de -ForceChangePassword son: True | False"}
}
{
    $TMP = $false, $true
    echo "------------------------------"
    Write-Host "In next Login."
    echo "------------------------------"
    Write-Host "1 - No force to chenge password"
    Write-Host "2 - Force to chenge password"
    $pos = Read-Host "Enter N°"
    $ForceChangePassword = $TMP[$pos-1]
}


<# Verificacion de parametro NewPassword #>

if(!$NewPassword)
{
    echo "------------------------------"
    $NewPassword = Read-Host "Enter the new password"
}


<# Ejecucion #>

$usuarios | ForEach-Object {
    Write-Verbose "$(fFormatDate) | Changing password: $($_.UserPrincipalName)"
    Set-MsolUserPassword -UserPrincipalName $_.UserPrincipalName -NewPassword $NewPassword -ForceChangePassword $ForceChangePassword
}


Write-Verbose "$(fFormatDate) | End of Script. (Duración: $(("{0:hh\:mm\:ss}" -f ((Get-Date) - $currentTime))))"
