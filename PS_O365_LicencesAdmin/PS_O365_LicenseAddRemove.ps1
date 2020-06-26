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
    Example: Add | Remove

 .-AccountSkuId
    Defines the license ID.
    Variable type [string].
    Example: enterprise:ENTERPRISEPACK

 .-servicios
    Defines which services  can be disable. Multiple items can be entered with comma.
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

    .\PS_O365_LicenseAddRemove.ps1 -tipoUsuario O365 -accion Add -location AR -AccountSkuId BancoComafi:ENTERPRISEPACK -servicios "EXCHANGE_S_ENTERPRISE","YAMMER_ENTERPRISE"

    .\PS_O365_LicenseAddRemove.ps1 -tipoUsuario UPN.csv -accion Add

    .\PS_O365_LicenseAddRemove.ps1 -tipoUsuario O365 -accion Remove

    .\PS_O365_LicenseAddRemove.ps1 -tipoUsuario UPN.csv -accion Remove -AccountSkuId BancoComafi:ENTERPRISEPACK

#>

[CmdLetBinding()]

param(
 [string] $tipoUsuario,
 [string] $accion,
 [string] $AccountSkuId,
 [object] $servicios = @(),
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

<# Asignar licencia #>

function fAddLicense {
    param(
        [Parameter(Mandatory=$true)]
        [object] $lic, #license
        [Parameter(Mandatory=$true)]
        [object] $serv, #license
        [Parameter(Mandatory=$true)]
        [object] $usrs, #usuarios
        [Parameter(Mandatory=$true)]
        [string] $loc #location
    )
    
    $CT = (Get-Date)


    <# Servicios a deshabilitar #>

    if($serv.count)
    {
        $services = @()
        $serv | ForEach-Object {
            if ($_ -in $lic.ServiceStatus.ServicePlan.ServiceName) {
                $services += $_
            }
            else {
                errorParameter -texto "El servicio $($_) no existe."
            }
        }
        $DisabledPlans = $true
    }
    else
    {
        $pos = Read-Host "Disable any service? (yes/no) [no]"
        if(($pos.ToLower() -eq "y") -or ($pos.ToLower() -eq "yes"))
        {
            echo "------------------------------"
            Write-Host "Select the services to disable (for multiples example: 1,2,4,7)"
            echo "------------------------------"
            fEnumerateList -array $lic.ServiceStatus.ServicePlan | select Item,ServiceName | ft
            $pos = Read-Host "Enter N°"
            $services = @()
            $pos.Split(",") | ForEach-Object {
                $services += $lic.ServiceStatus.ServicePlan[$_ -1].ServiceName
            }
            $DisabledPlans = $true

        }
        else {$DisabledPlans = $false}
    }


    <# Generacion del plan #>

    if($DisabledPlans){$ServicePlan = New-MsolLicenseOptions -AccountSkuId $lic.AccountSkuId -DisabledPlans $services}
    else{$ServicePlan = New-MsolLicenseOptions -AccountSkuId $lic.AccountSkuId}


    <# Agregado de Licencia #>

    ForEach($user in $usrs) {
        # Verifico si posee una licencia
        if ($user.IsLicensed -eq $true)
        {
            # Busco el tipo de licencia a trabajar
            $UserLicenses = $user.licenses | Where-Object {$_.AccountSkuId -Eq $lic.AccountSkuId}
            # Verifico si es el tipo de licencia que requiero
            if($UserLicenses){Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Usuario: $($user.UserPrincipalName.ToString()) | Asignado"}
            else
            {
                Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Usuario: $($user.UserPrincipalName.ToString()) | Asignando"
                Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses $lic.AccountSkuId -LicenseOptions $ServicePlan
            }
        }
        else
        {
            Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Usuario: $($user.UserPrincipalName.ToString()) | Asignando"
            Set-MsolUser -UserPrincipalName $user.UserPrincipalName -UsageLocation $Location
            Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses $lic.AccountSkuId -LicenseOptions $ServicePlan
        }
    }
    Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Final de asignacion de licencias. (Duración: $(("{0:hh\:mm\:ss}" -f ((Get-Date) - $CT))))"
}


<# Remover licencia #>

function fRemoveLicense {
    param(
        [Parameter(Mandatory=$true)]
        [object] $lic, #license
        [Parameter(Mandatory=$true)]
        [object] $usrs #usuarios
    )
    $CT = (Get-Date)

    <# Remocion de Licencia #>

    ForEach($user in $usrs) {
        if ($user.IsLicensed -eq $true)
        {
            # Busco el tipo de licencia a trabajar
            $UserLicenses = $user.licenses | Where-Object {$_.AccountSkuId -Eq $lic.AccountSkuId}
            # Verifico si es el tipo de licencia que requiero
            if($UserLicenses)
            {
                Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Usuario: $($user.UserPrincipalName.ToString()) | Removiendo"
                Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -RemoveLicenses $lic.AccountSkuId
            }
            else {Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Usuario: $($user.UserPrincipalName.ToString()) | Sin licencia"}
        }
        else {Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Usuario: $($user.UserPrincipalName.ToString()) | Sin licencia"}
    }
    Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Final de remocion de licencias. (Duración: $(("{0:hh\:mm\:ss}" -f ((Get-Date) - $CT))))"
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
    if($accion -eq "Add"){$accion = "1"}
    elseif($accion -eq "Remove"){$accion = "2"}
    else{errorParameter -texto "Definición incorrecta. Las opciones de -accion son: Add o Remove."}
}
else
{
    echo "------------------------------"
    Write-Host "Select an Action"
    echo "------------------------------"
    Write-Host "1 - Add License"
    Write-Host "2 - Remove License"
    $pos = Read-Host "Enter N°"
    $accion = $pos
}


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
else
{
    errorParameter -texto "Definición incorrecta. Las opciones de -tipoUsuario son: O365 o nombre del archivo CSV (ejem: UPN.csv)"
}


<# Busqueda de usuarios #>

if ($tipoUsuario -eq "O365")
{
    $usuarios = BuscarUsuarios # Busca todos los usuarios de O365
}
else
{
    $usuarios = BuscarUsuariosLista -listado $UPNs # Busca los usuarios mediante una lista
}


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


<# Ejecucion #>

switch ($accion) 
    { 
        "1" {
            fAddLicense -usrs $usuarios -lic $license -serv $servicios -loc $Location
        } # 1
        "2" {
            fRemoveLicense -usrs $usuarios -lic $license
        } # 2
        default {"Opcion no determinada"}
    }

Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | End of Script. (Duración: $(("{0:hh\:mm\:ss}" -f ((Get-Date) - $currentTime))))"
