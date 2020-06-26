<#
 .SYNOPSIS
    Allow to login to Office 365.

 .DESCRIPTION
    This script allows to login to Office 365. It can use the parameters to fill the command or use the interactive console to select or input data.

 .PARAMETER
 .-user
    Defines the user name.
    Variable type [string].
    Example: usuario@outlook.com

 .-password
    Defines the password of the user.
    Variable type [string].
    Example: C0ntr4s3ñ4

.EXAMPLE

    .\PS_O365.ps1 -user 'unUsuario' -password 'C0ntr4s3ñ4'

#>

param(
 [string] $user,
 [string] $password
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


<# Error Output #>

function fErrorParameter
{
    param(
        [string] $texto
    )
    process
    {
        Write-host "$(fFormatDate) | $($texto)" -ForegroundColor Red
        Write-host "$(fFormatDate) | End of the script."
        exit 1
    }
}


<#
******************************************************************************
Script
******************************************************************************
#>

$currentTime = (Get-Date)
Write-Host "$(fFormatDate) | Starting Login..."

if(!$user)
{
    $user = Read-Host "Enter user"
}

if($password)
{
    $usrPassword = ConvertTo-SecureString -String $password -AsPlainText -Force
}
else
{
    $usrPassword = Read-Host -assecurestring "Enter password"
}

# sign in

$cred = New-Object System.Management.Automation.PSCredential ($user, $usrPassword)

$conection = Connect-MsolService -Credential $cred -ErrorAction SilentlyContinue;


try
    {
        $ErrorActionPreference = "Stop"
        Connect-MsolService -Credential $cred
        Write-Host "$(fFormatDate) | Welcome. (Duration: $(("{0:hh\:mm\:ss}" -f ((Get-Date) - $currentTime))))"
    }
    catch
    {
        fErrorParameter -texto "Wrong user name or password."
    }
    finally
    {
        $ErrorActionPreference = "Continue"
    }



