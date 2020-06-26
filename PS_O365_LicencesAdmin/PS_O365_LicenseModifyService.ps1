<#
 .SYNOPSIS
    Allow to add or remove a License for users in Office 365 by powershell.

 .DESCRIPTION
    This script allows to add or remove a license to a list of users in Office 365. Can be executed for a custom list o for all users in Office 365.

 .PRE-REQUISITES
    Not Mandatory
    * Have the CSV file with the UPN complete in the same path than the script.

 .PARAMETERS
 .-tipoUsuario [Mandatory]
    Defines the CSV file name or if use all O365 users.
    Variable type [string].
    Example: O365 --> for all users.
    Example: NombreEjemplo.csv --> for custom users.

 .-accion
    Defines what the script must to do. Add or remove a license.
    Variable type [string].
    Example: Enable | Disable

 .-AccountSkuId
    Defines the license ID.
    Variable type [string].
    Example: enterprise:ENTERPRISEPACK

 .-servicios
    Defines which services  can be disable.
    Variable type [object].
    Example: "EXCHANGE_S_ENTERPRISE"
    Example: "EXCHANGE_S_ENTERPRISE","YAMMER_ENTERPRISE"

 .-location
    Defines the location for the license.
    Variable type [string].
    Example: AR (for default)

 .-v or -Verbose
    Sets Verbose mode ON.

.EXAMPLES

    .\PS_O365_LicenseModifyService.ps1 -tipoUsuario O365 -accion Enable -location AR -AccountSkuId enterprise:ENTERPRISEPACK -servicios "YAMMER_ENTERPRISE"

    .\PS_O365_LicenseModifyService.ps1 -tipoUsuario UPN.csv -accion Enable

    .\PS_O365_LicenseModifyService.ps1 -tipoUsuario O365 -accion Disable

    .\PS_O365_LicenseModifyService.ps1 -tipoUsuario UPN.csv -accion Disable -AccountSkuId enterprise:ENTERPRISEPACK

#>

[CmdLetBinding()]

param(
 [string] $tipoUsuario,
 [string] $accion,
 [string] $AccountSkuId,
 [string] $servicio,
 [string] $Location = "AR"
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


<# Error Output #>

function errorParameter
{
    param(
        [string] $texto
    )
    process
    {
        Write-host "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | $($texto)" -ForegroundColor Red
        Write-host "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | fin del script"
        exit 1
    }
}


<# Buscar usuarios en Office 365 #>

function BuscarUsuarios {
    param()
    
    $CT = (Get-Date)
    Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Buscando todos los usuarios de la cuenta de O365..."
    #Busca todos los usuarios de O365
    $users = Get-MsolUser -all
    #$users = Get-MsolUser -MaxResults 15
    Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Final de la busqueda. (Duración: $(("{0:hh\:mm\:ss}" -f ((Get-Date) - $CT))))"
    return $users
}


<# Busca usuarios en Office 365 mediante una lista previa #>

function BuscarUsuariosLista {
    param(
        [Parameter(Mandatory=$true)]
        [object] $listado
    )

    $CT = (Get-Date)
    # Busca todos los usuarios a partir de un listado
    Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Buscando lista de usuarios en la cuenta de O365..."
    $users = $listado | ForEach-Object {
        Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Buscando usuario:  $($_.UserPrincipalName)..."
        Get-MsolUser -UserPrincipalName $_.UserPrincipalName
    }
    Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Final de la busqueda. (Duración: $(("{0:hh\:mm\:ss}" -f ((Get-Date) - $CT))))"
    return $users
}


<#
******************************************************************************
Script
******************************************************************************
#>

$currentTime = (Get-Date)
Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Start Script..."

<# accion #>

if($accion)
{
    if($accion -eq "Enable"){$action = $true}
    elseif($accion -eq "Disable"){$action = $false}
    else{errorParameter -texto "Definición incorrecta. Las opciones de -accion son: Enable o Disable."}
}
else
{
    $TMP = $true,$false
    echo "------------------------------"
    Write-Host "Select an Action"
    echo "------------------------------"
    Write-Host "1 - Enable Service"
    Write-Host "2 - Disable Service"
    $pos = Read-Host "Enter N°"
    $action = $TMP[$pos-1]
}

$tipoUsuario = "UPNtest.csv"
<# Verificador de tipo de usuario #>

if (($tipoUsuario -eq "O365") -or ($tipoUsuario -like "*.csv"))
{
    if ($tipoUsuario -ne "O365")
    {
        #Importacion del CSV
        try
        {
            $ErrorActionPreference = "Stop"
            $UPNs = Import-Csv $tipoUsuario
            Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Importando CSV..."
        }
        catch
        {
            errorParameter -texto "Error con la importacion del archivo CSV. Verificar su formato o nombre."
        }
        finally
        {
           $ErrorActionPreference = "Continue"
        }
    }
}
else{errorParameter -texto "Definición incorrecta. Las opciones de -tipoUsuario son: O365 o nombre del archivo CSV (ejem: UPN.csv)"}


<# Busqueda de usuarios #>

if ($tipoUsuario -eq "O365"){$usuarios = BuscarUsuarios}
else{$usuarios = BuscarUsuariosLista -listado $UPNs}


<# Busqueda de Licencias #>

Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Searching Licenses..."
$licencias = Get-MsolAccountSku
if($AccountSkuId)
{
    $license = $licencias | Where-Object {$_.AccountSkuId -Eq $AccountSkuId}
    if(!$license) {errorParameter -texto "Definición incorrecta. La licencia $($AccountSkuId) no existe"}
}
else
{
    Write-Host "Select a license"
    echo "------------------------------"
    fEnumerateList -array $licencias | select Item,AccountSkuId | ft
    $pos = Read-Host "Enter N°"
    $license = $licencias[$pos-1]
}

<# Servicios a Modificar #>

if($servicio){if (!($servicio -in $license.ServiceStatus.ServicePlan.ServiceName)) {errorParameter -texto "El servicio $($servicio) no existe."}}
else
{
    echo "------------------------------"
    Write-Host "Select the services to Modify"
    echo "------------------------------"
    fEnumerateList -array $license.ServiceStatus.ServicePlan | select Item,ServiceName | ft
    $pos = Read-Host "Enter N°"
    $servicio = $license.ServiceStatus.ServicePlan[$pos -1].ServiceName
}


<# Ejecucion #>

ForEach ($usr in $usuarios) {
    if ($usr.IsLicensed -eq $true)
    {
        $UserLicense = $usr.licenses | Where-Object {$_.AccountSkuId -Eq $license.AccountSkuId}
        if($UserLicense)
        {
            $UserLicensesService = $userLicense | Select-Object -ExpandProperty ServiceStatus | Where-Object {$_.ServicePlan.ServiceName -like $servicio}
            if($action)
            {
                if ($UserLicensesService.ProvisioningStatus -like "Disabled")
                {
                    Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Usuario activado: $($usr.UserPrincipalName.ToString())"
                    $DisabledPlans = @()
                    $DisabledPlans += $UserLicense.ServiceStatus | Where-Object {($_.ProvisioningStatus -Eq "Disabled") -and ($_.ServicePlan.ServiceName -ne $servicio)} | % {$_.ServicePlan.ServiceName}
                    $ServicePlan = New-MsolLicenseOptions -AccountSkuId $license.AccountSkuId -DisabledPlans $DisabledPlans
                    Set-MsolUserLicense -UserPrincipalName $usr.UserPrincipalName -LicenseOptions $ServicePlan
                }
                else{Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Usuario sin cambio: $($usr.UserPrincipalName.ToString())"}
            }
            else
            {
                if ($UserLicensesService.ProvisioningStatus -notlike "Disabled")
                {
                    Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Usuario desactivado: $($usr.UserPrincipalName.ToString())"
                    $DisabledPlans = @()
                    $DisabledPlans += $UserLicense.ServiceStatus | Where-Object {$_.ProvisioningStatus -Eq "Disabled"} | % {$_.ServicePlan.ServiceName}
                    $DisabledPlans += $servicio
                    $ServicePlan = New-MsolLicenseOptions -AccountSkuId $license.AccountSkuId -DisabledPlans $DisabledPlans
                    Set-MsolUserLicense -UserPrincipalName $usr.UserPrincipalName -LicenseOptions $ServicePlan
                }
                else{Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Usuario sin cambio: $($usr.UserPrincipalName.ToString())"}
            }
        }
        else{Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Usuario sin licencia: $($usr.UserPrincipalName.ToString())"}
    }
    else{Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Usuario sin licencia: $($usr.UserPrincipalName.ToString())"}
}

Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | End of Script. (Duración: $(("{0:hh\:mm\:ss}" -f ((Get-Date) - $currentTime))))"



