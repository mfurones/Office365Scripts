<#
 .SYNOPSIS
    Allow to generate a report of Licenses and Services from Office 365 by powershell.

 .DESCRIPTION
    This script allows to generate two types of reports, by License or Services. It can generato for a custom list o for all users in Office 365.
    The output of the reports is in CSV format.

 .PRE-REQUISITES
    Not Mandatory
    * Have the CSV file with the UPN complete in the same path than the script.

 .PARAMETERS
 .-tipoUsuario
    Defines the CSV file name or if use all O365 users.
    Variable type [string].
    Example: O365 --> for all users.
    Example: NombreEjemplo.csv --> for custom users.

 .-tipoReporte
    Defines the report type.
    Variable type [string] | License o Services.
    Example: License o Services

 .-v or -Verbose
    Sets Verbose mode ON.

.EXAMPLE

    .\PS_O365_License_Report.ps1 -tipoUsuario O365 -tipoReporte Licencias

    .\PS_O365_License_Report.ps1 -tipoUsuario UPN.csv -tipoReporte Licencias

    .\PS_O365_License_Report.ps1 -tipoUsuario O365 -tipoReporte Servicios

    .\PS_O365_License_Report.ps1 -tipoUsuario UPN.csv -tipoReporte Servicios

#>

[CmdLetBinding()]

param(
 [string] $tipoUsuario,
 [string] $tipoReporte
 )


<#
******************************************************************************
Aviso de error & exit
******************************************************************************
#>

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

<#
******************************************************************************
Buscar usuarios en Office 365
******************************************************************************
#>

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

<#
******************************************************************************
Busca usuarios en Office 365 mediante una lista previa
******************************************************************************
#>

function BuscarUsuariosLista {
    param(
        [Parameter(Mandatory=$true)]
        [object] $listado
    )

    $CT = (Get-Date)
    # Busca todos los usuarios a partir de un listado
    $users = $listado | ForEach-Object {
        Get-MsolUser -UserPrincipalName $_.UserPrincipalName
    }
    Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Final de la busqueda. (Duración: $(("{0:hh\:mm\:ss}" -f ((Get-Date) - $CT))))"
    return $users
}


<#
******************************************************************************
Lista los tipos de Licencia que tienen habilitados.
******************************************************************************
#>

function LicenciasHabilitadas {
    param(
        [Parameter(Mandatory=$true)]
        [object] $USRList,
        [Parameter(Mandatory=$true)]
        [object] $ASIL
    )

    $CT = (Get-Date)
    Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Concatenando los registros de los Usuarios..."
    # Listado total de reportes
    $reporte = @()
    #recorro la lista de usuarios
    foreach ($usr in $USRList)
    {
        # Genero un registro de usuario
        $linea = $usr | select FirstName, LastName, UserPrincipalName, IsLicensed

        # Verifico si tiene licencia
        if ($usr.IsLicensed -eq $true)
        {
            # Expando las propiedades de licencia
            $license = $usr | Select-Object -ExpandProperty Licenses
            foreach ($ASI in $ASIL)
            {
                $status = $license | Where-Object {$_.AccountSkuId -like $ASI.AccountSkuId}
                if ($status)
                {
                $linea | Add-Member -MemberType NoteProperty -Name $ASI.AccountSkuId.split(":")[1] -Value $true -Force
                }
                else
                {
                $linea | Add-Member -MemberType NoteProperty -Name $ASI.AccountSkuId.split(":")[1] -Value $false -Force
                }
            }
        }
        # Agrego el registro de usuario al reporte
        $reporte = $reporte + $linea
    }

    Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Exportando..."
    # Exporto el reporte como CSV

    $Archivo = "Licencias_Reporte.csv"
    $reporte | export-csv $Archivo

    Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Exportacion finalizada. (Duración: $(("{0:hh\:mm\:ss}" -f ((Get-Date) - $CT))))"
}


<#
******************************************************************************
Reporte de servicio de una licencia en particular
******************************************************************************
#>

function LicenciasServicios {
    param(
        [Parameter(Mandatory=$true)]
        [object] $USRList,
        [Parameter(Mandatory=$true)]
        [string] $AccountSkuId
    )

    $CT = (Get-Date)
    Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Concatenando los registros de los Usuarios..."

    # Servicios de la licencia
    $ServicePlans = (Get-MsolAccountSku | Where {$_.AccountSkuId -eq $AccountSkuId}).ServiceStatus.ServicePlan.ServiceName
    # Listado total de reportes
    $reporte = @()
    #recorro la lista de usuarios
    foreach ($usr in $USRList)
    {
        # Genero un registro de usuario
        $linea = $usr | select FirstName, LastName, UserPrincipalName, IsLicensed
        # Verifico si tiene licencia
        if ($usr.IsLicensed -eq $true)
        {
            # filtro el tipo de licencia que quiero
            $license = $usr | Select-Object -ExpandProperty Licenses | Where-Object {$_.AccountSkuId -like $AccountSkuId} 
            if ($license)
            {
                # agrego linea con un true al tipo de licencia al registro del usuario
                $linea | Add-Member -MemberType NoteProperty -Name $AccountSkuId.Split(":")[1] -Value $true -Force
                # Expando las propiedades de licencia
                # Recorro cada uno de los servicios de la licencia
                $license | Select-Object -ExpandProperty ServiceStatus | ForEach-Object{
                    # Agrego cada servicio al registro del usuario
                    $linea | Add-Member -MemberType NoteProperty -Name $_.ServicePlan.ServiceName -Value $_.ProvisioningStatus -Force
                }
            }
            else
            {
                 # agrego linea con un false al tipo de licencia
                $linea | Add-Member -MemberType NoteProperty -Name $AccountSkuId.Split(":")[1] -Value $false -Force
                # Completo los campos en blanco del cuadro de servicios
                $ServicePlans | ForEach-Object {$linea | Add-Member -MemberType NoteProperty -Name $_ -Value "" -Force}
            }
        }
        # Agrego el registro de usuario al reporte
        $reporte = $reporte + $linea
    }

    Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Exportando..."
    # Exporto el reporte como CSV

    $Archivo = "Licencias_Reporte_$($AccountSkuId.split(":")[1]).csv"
    $reporte | export-csv $Archivo

    Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Exportacion finalizada.  (Duración: $(("{0:hh\:mm\:ss}" -f ((Get-Date) - $CT))))"
}


<#
******************************************************************************
Verificacion de parametros
******************************************************************************
#>

if (!$tipoUsuario -and !$tipoReporte)
{
    errorParameter -texto "Error. Parametros sin definir."
}

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

if (!("Licencias","Servicios"  -contains $tipoReporte))
{
    errorParameter -texto "Definición incorrecta. Las opciones de -tipoReporte son: Licencias o Servicios."
}


<#
******************************************************************************
Ejecucion de Script
******************************************************************************
#>

$currentTime = (Get-Date)
Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Inicio del script de Reportes..."

if ($tipoUsuario -eq "O365")
{
    $usuarios = BuscarUsuarios # Busca todos los usuarios de O365
}
else
{
    $usuarios = BuscarUsuariosLista -listado $UPNs # Busca los usuarios mediante una lista
}

Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Generando lista de Licencias..."
$AccountSkuIdList = Get-MsolAccountSku | select AccountSkuId

if ($tipoReporte -eq "Licencias")
{
    #Listado de Licencias
    LicenciasHabilitadas -USRList $usuarios -ASIL $AccountSkuIdList
}
else
{
    # Listado de todos los Servicios por Licencia
    $AccountSkuIdList | ForEach-Object { LicenciasServicios -USRList $usuarios -AccountSkuId $_.AccountSkuId}
}

Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Fin del script de Reportes. (Duración: $(("{0:hh\:mm\:ss}" -f ((Get-Date) - $currentTime))))"

