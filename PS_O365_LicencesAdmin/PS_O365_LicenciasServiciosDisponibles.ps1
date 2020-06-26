<#
 .SYNOPSIS
    Allow to generate a list of Licenses and Services available in Office 365 by powershell.

 .DESCRIPTION
    This script allows to generate a list of Licenses and Services available in Office 365. Can be visualised in console mode or generate an output to a CSV file.

 .PARAMETERS

 .-v or -Verbose
    Sets Verbose mode ON.

.EXAMPLE

    .\PS_O365_License_Report.ps1

#>

[CmdLetBinding()]

param(

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

<#
******************************************************************************
Script
******************************************************************************
#>

$currentTime = (Get-Date)
Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Start Script..."

<# Menu de accion #>

Write-Host "Select an option"
echo "------------------------------"
Write-Host "1 - Types of Licenses available"
Write-Host "2 - Types of services by license"
$pos = Read-Host "Enter N°"
$accion = $pos

Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Searching Licenses..."

$licencias = Get-MsolAccountSku

switch ($accion) 
    { 
        1 {
            Write-Host $($licencias | Out-String)
            $pos = Read-Host "Export results to CSV file? (yes/no) [no]"
            if(($pos.ToLower() -eq "y") -or ($pos.ToLower() -eq "yes"))
            {
                Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Exporting to: PS_O365_LicenciasReport.csv"
                $licencias | Export-Csv "PS_O365_LicenciasReport.csv"
            }
        } # 1
        2 {
            
            Write-Host "Select a license"
            echo "------------------------------"
            fEnumerateList -array $licencias | select Item,AccountSkuId | ft
            $pos = Read-Host "Enter N°"
            $license = $licencias[$pos-1]
            Write-Host $($license.ServiceStatus | Out-String)
            $pos = Read-Host "Export results to CSV file? (yes/no) [no]"
            if(($pos.ToLower() -eq "y") -or ($pos.ToLower() -eq "yes"))
            {
                Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Generating report..."
                $reporte = @()
                $license.ServiceStatus | ForEach-Object {
                    $linea = $_.ServicePlan | Select-Object ServiceName
                    $linea | Add-Member -MemberType NoteProperty -Name "ProvisioningStatus" -Value $_.ProvisioningStatus -Force
                    $reporte += $linea
                }
                Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | Exporting to: PS_O365_ServicesReport$($license.AccountSkuId.Split(":")[1]).csv"
                $reporte | Export-Csv "PS_O365_ServicesReport$($license.AccountSkuId.Split(":")[1]).csv"
            }
        } # 2
        default {"Opcion no determinada"}
    }


Write-Verbose "$(Get-Date -UFormat "%d/%m/%Y") | $(Get-Date -Format T) | End of Script. (Duración: $(("{0:hh\:mm\:ss}" -f ((Get-Date) - $currentTime))))"
