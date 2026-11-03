<#
.SYNOPSIS
    Lista los servers (Azure VMs + Azure Arc) por tag AUMSchedule.

.DESCRIPTION
    Ejecuta un query de Azure Resource Graph sobre todas las suscripciones accesibles
    y exporta los servers cuyo tag AUMSchedule coincide con los schedules indicados.

.PARAMETER Schedules
    Valores del tag AUMSchedule a buscar (sin distinguir mayúsculas).
.PARAMETER OutputPath
    Ruta del CSV de salida.
.PARAMETER TenantId
    Tenant al que conectarse (ID o dominio, ej. contoso.onmicrosoft.com).
.PARAMETER UseDeviceCode
    Inicia sesión con código de dispositivo (útil si la ventana de login no aparece).

.EXAMPLE
    .\Patchreview.ps1
    .\Patchreview.ps1 -Schedules 'PRD-EMEA-Infra-Ring1' -OutputPath C:\Reportes\Ring1.csv
#>
[CmdletBinding()]
param(
    [string[]]$Schedules = @(
        'NON-PRD-APAC-App-Ring2', 'NON-PRD-EMEA-App-Ring2', 'NON-PRD-LATM-App-Ring2', 'NON-PRD-NAMR-App-Ring2',
        'PRD-EMEA-Infra-Ring1', 'PRD-EMEA-Infra-Ring2', 'PRD-LATM-Infra-Ring1', 'PRD-NAMR-Infra-Ring1'
    ),
    [string]$OutputPath = ".\AUMSchedule_Servers_$(Get-Date -Format 'yyyyMMdd').csv",
    [string]$TenantId,
    [switch]$UseDeviceCode
)

$ErrorActionPreference = 'Stop'

# --- Módulos -------------------------------------------------------------------
foreach ($m in 'Az.Accounts', 'Az.ResourceGraph') {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Host "Instalando $m ..." -ForegroundColor Yellow
        Install-Module $m -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $m
}

# --- Conexión ------------------------------------------------------------------
# Siempre pide credenciales y no reutiliza sesiones guardadas (se administran varios tenants)
Disable-AzContextAutosave -Scope Process | Out-Null
Clear-AzContext -Scope Process -Force -ErrorAction SilentlyContinue
$connect = @{}
if ($TenantId)      { $connect.Tenant = $TenantId }
if ($UseDeviceCode) { $connect.UseDeviceAuthentication = $true }
Connect-AzAccount @connect | Out-Null
$ctx = Get-AzContext
Write-Host "Cuenta: $($ctx.Account.Id) | Tenant: $($ctx.Tenant.Id)" -ForegroundColor Green

# --- Query ---------------------------------------------------------------------
$scheduleList = ($Schedules | ForEach-Object { "'$_'" }) -join ', '
$query = @"
resources
| where type in~ ('microsoft.compute/virtualmachines', 'microsoft.hybridcompute/machines')
| mv-expand bagexpansion=array tags
| where tostring(tags[0]) =~ 'AUMSchedule'
| extend AUMSchedule = tostring(tags[1])
| where AUMSchedule in~ ($scheduleList)
| extend OS = coalesce(tostring(properties.storageProfile.osDisk.osType), tostring(properties.osType)),
         Platform = iff(type =~ 'microsoft.hybridcompute/machines', 'Azure Arc', 'Azure VM')
| join kind=leftouter (
    resourcecontainers
    | where type =~ 'microsoft.resources/subscriptions'
    | project subscriptionId, SubscriptionName = name
  ) on subscriptionId
| project name, AUMSchedule, Platform, OS, resourceGroup, SubscriptionName, location, id
| order by AUMSchedule asc, name asc
"@

# --- Ejecución con paginación --------------------------------------------------
Write-Host "Consultando Azure Resource Graph..." -ForegroundColor Cyan
$results   = [System.Collections.Generic.List[object]]::new()
$skipToken = $null
do {
    $params = @{ Query = $query; First = 1000; UseTenantScope = $true }
    if ($skipToken) { $params.SkipToken = $skipToken }
    $page = Search-AzGraph @params
    $results.AddRange([object[]]@($page.Data))
    $skipToken = $page.SkipToken
} while ($skipToken)

$results | Select-Object name, AUMSchedule, Platform, OS, resourceGroup, SubscriptionName, location, id |
    Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

# --- Resumen -------------------------------------------------------------------
Write-Host "`nListo: $((Resolve-Path $OutputPath).Path) ($($results.Count) servers)" -ForegroundColor Green
$results | Group-Object AUMSchedule | Sort-Object Name |
    Format-Table @{ n = 'AUMSchedule'; e = { $_.Name } }, Count -AutoSize
