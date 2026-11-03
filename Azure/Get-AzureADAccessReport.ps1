<#
.SYNOPSIS
    H2 Campaign Sync 2026 - Lista de cuentas activas + permisos de Azure AD (Entra ID).

.DESCRIPTION
    Genera un solo CSV con usuarios, cuentas de servicio, aplicaciones (service principals)
    e identidades administradas (cuentas de sistema), con las columnas que piden:

      Account Type | Account Name | User Name | Email | Emp ID | SAML ID / UserID |
      Access/Roles | Status | EndPoint | Account Creation Date | Last Sign-In | Source

    Access/Roles incluye:
      - Roles de Entra ID asignados (directos, vía grupo y elegibles en PIM si hay licencia P2)
      - Aplicaciones empresariales asignadas al usuario (y el rol de la app)
      - Para aplicaciones: permisos de API (aplicación y delegados)
      - Opcional (-IncludeGroups): membresía de grupos

.PARAMETER OutputPath
    Ruta del CSV de salida.
.PARAMETER IncludeDisabled
    Incluye cuentas deshabilitadas (por defecto solo activas).
.PARAMETER IncludeGroups
    Agrega los grupos de cada usuario a Access/Roles (más lento en tenants grandes).
.PARAMETER IncludeMicrosoftApps
    Incluye los service principals propios de Microsoft (cientos, normalmente no se piden).
.PARAMETER ServiceAccountPattern
    Regex sobre UPN / nombre para marcar usuarios como "Service Account".

.EXAMPLE
    .\Get-AzureADAccessReport.ps1
    .\Get-AzureADAccessReport.ps1 -IncludeGroups -OutputPath C:\Reportes\H2.csv
#>
[CmdletBinding()]
param(
    [string]$OutputPath = ".\H2_Campaign_Sync_2026_AzureAD_$(Get-Date -Format 'yyyyMMdd').csv",
    [switch]$IncludeDisabled,
    [switch]$IncludeGroups,
    [switch]$IncludeMicrosoftApps,
    [string]$ServiceAccountPattern = '^(svc|srv|sa|service)[-_.]|serviceaccount|noreply|no-reply|^sync_|^adsync'
)

$ErrorActionPreference = 'Stop'

# --- Módulos -------------------------------------------------------------------
$modules = 'Microsoft.Graph.Authentication', 'Microsoft.Graph.Users', 'Microsoft.Graph.Groups',
           'Microsoft.Graph.Applications', 'Microsoft.Graph.Identity.DirectoryManagement',
           'Microsoft.Graph.Identity.Governance', 'Microsoft.Graph.Identity.SignIns'

# Windows PowerShell 5.1 tiene un límite de 4096 funciones y Microsoft.Graph lo rebasa
$MaximumFunctionCount = 32768
$MaximumVariableCount = 32768

# Todos los submódulos de Microsoft.Graph deben cargarse en la MISMA versión
function Get-CommonGraphVersion {
    $common = $null
    foreach ($m in $modules) {
        $v = @(Get-Module -ListAvailable -Name $m | ForEach-Object { $_.Version.ToString() })
        $common = if ($null -eq $common) { $v } else { @($common | Where-Object { $_ -in $v }) }
    }
    $common | Sort-Object { [version]$_ } -Descending | Select-Object -First 1
}

$version = Get-CommonGraphVersion
if (-not $version) {
    Write-Host "Instalando/actualizando módulos Microsoft.Graph (misma versión)..." -ForegroundColor Yellow
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    foreach ($m in $modules) { Install-Module $m -Scope CurrentUser -Force -AllowClobber -Repository PSGallery }
    $version = Get-CommonGraphVersion
}

# Evita choques con versiones ya cargadas en esta sesión
Get-Module Microsoft.Graph.* | Remove-Module -Force -ErrorAction SilentlyContinue
foreach ($m in $modules) {
    try {
        Import-Module $m -RequiredVersion $version -ErrorAction Stop
    } catch {
        Write-Host "`nNo se pudo cargar $m $version" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host "Prueba en una ventana NUEVA de PowerShell (idealmente PowerShell 7: winget install Microsoft.PowerShell)." -ForegroundColor Yellow
        exit 1
    }
}
Write-Host "Microsoft.Graph $version cargado" -ForegroundColor Green

# --- Conexión (solo lectura) ---------------------------------------------------
Connect-MgGraph -NoWelcome -Scopes 'User.Read.All', 'Application.Read.All', 'Directory.Read.All',
    'RoleManagement.Read.Directory', 'AuditLog.Read.All', 'GroupMember.Read.All'
$org      = Get-MgOrganization
$tenantId = $org.Id
$domain   = ($org.VerifiedDomains | Where-Object IsDefault).Name
$endpoint = "Azure AD / Entra ID ($domain)"
Write-Host "Conectado a $domain ($tenantId)" -ForegroundColor Green

# Tenants dueños de las apps de primera parte de Microsoft
$msTenants = 'f8cdef31-a31e-4b4a-93e4-5f571e91255a', '72f988bf-86f1-41af-91ab-2d7cd011db47'

# --- Helpers -------------------------------------------------------------------
$access = @{}   # principalId -> lista de accesos
function Add-Access([string]$Id, [string]$Text) {
    if (-not $access.ContainsKey($Id)) { $access[$Id] = [System.Collections.Generic.List[string]]::new() }
    if (-not $access[$Id].Contains($Text)) { $access[$Id].Add($Text) }
}

$groupCache = @{}   # groupId -> @{ Name; Members }
function Get-GroupInfo([string]$GroupId) {
    if (-not $groupCache.ContainsKey($GroupId)) {
        try {
            $g = Get-MgGroup -GroupId $GroupId -Property DisplayName
            $m = @(Get-MgGroupTransitiveMember -GroupId $GroupId -All | Select-Object -ExpandProperty Id)
            $groupCache[$GroupId] = @{ Name = $g.DisplayName; Members = $m }
        } catch {
            $groupCache[$GroupId] = $null
        }
    }
    $groupCache[$GroupId]
}

function Get-AppRoleName($ResourceSp, [string]$AppRoleId) {
    if ($AppRoleId -eq '00000000-0000-0000-0000-000000000000') { return 'Default Access' }
    $r = $ResourceSp.AppRoles | Where-Object Id -eq $AppRoleId
    if ($r) { if ($r.Value) { $r.Value } else { $r.DisplayName } } else { $AppRoleId }
}

# --- 1. Usuarios ---------------------------------------------------------------
Write-Host "Obteniendo usuarios..." -ForegroundColor Cyan
$userProps = 'Id', 'DisplayName', 'UserPrincipalName', 'Mail', 'EmployeeId', 'UserType',
             'AccountEnabled', 'CreatedDateTime', 'OnPremisesSyncEnabled', 'SignInActivity'
try {
    $users = Get-MgUser -All -Property $userProps
} catch {
    Write-Warning "Sin acceso a SignInActivity (requiere Entra ID P1). Se omite 'Last Sign-In'."
    $users = Get-MgUser -All -Property ($userProps | Where-Object { $_ -ne 'SignInActivity' })
}
if (-not $IncludeDisabled) { $users = $users | Where-Object AccountEnabled }
$userIds = [System.Collections.Generic.HashSet[string]]::new([string[]]@($users.Id))

# --- 2. Service principals (apps + identidades administradas) -----------------
Write-Host "Obteniendo aplicaciones y service principals..." -ForegroundColor Cyan
$allSps = Get-MgServicePrincipal -All -Property Id, AppId, DisplayName, ServicePrincipalType,
          AccountEnabled, AppOwnerOrganizationId, AppRoles, Oauth2PermissionScopes
$spById = @{}
foreach ($sp in $allSps) { $spById[$sp.Id] = $sp }

$appCreated = @{}
Get-MgApplication -All -Property AppId, CreatedDateTime | ForEach-Object { $appCreated[$_.AppId] = $_.CreatedDateTime }

$targetSps = $allSps | Where-Object {
    $_.ServicePrincipalType -eq 'ManagedIdentity' -or
    $IncludeMicrosoftApps -or
    ($_.AppOwnerOrganizationId -notin $msTenants)
}
if (-not $IncludeDisabled) { $targetSps = $targetSps | Where-Object AccountEnabled }

# --- 3. Roles de Entra ID ------------------------------------------------------
Write-Host "Obteniendo roles de directorio..." -ForegroundColor Cyan
$roleDefs = @{}
Get-MgRoleManagementDirectoryRoleDefinition -All | ForEach-Object { $roleDefs[$_.Id] = $_.DisplayName }

function Add-RoleAccess([string]$PrincipalId, [string]$Label) {
    if ($userIds.Contains($PrincipalId) -or $spById.ContainsKey($PrincipalId)) {
        Add-Access $PrincipalId $Label
    } else {
        # El principal es un grupo: se expande a sus miembros
        $g = Get-GroupInfo $PrincipalId
        if ($g) { foreach ($mId in $g.Members) { Add-Access $mId "$Label (via group: $($g.Name))" } }
    }
}

Get-MgRoleManagementDirectoryRoleAssignment -All | ForEach-Object {
    Add-RoleAccess $_.PrincipalId "Entra Role: $($roleDefs[$_.RoleDefinitionId])"
}
try {
    Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All | ForEach-Object {
        Add-RoleAccess $_.PrincipalId "Entra Role (PIM Eligible): $($roleDefs[$_.RoleDefinitionId])"
    }
} catch {
    Write-Warning "No se pudieron leer roles elegibles de PIM (requiere Entra ID P2). Se omiten."
}

# --- 4. Asignaciones a apps y permisos de API ---------------------------------
Write-Host "Obteniendo asignaciones de aplicaciones y permisos de API..." -ForegroundColor Cyan
$i = 0
foreach ($sp in $targetSps) {
    $i++
    Write-Progress -Activity 'Aplicaciones' -Status $sp.DisplayName -PercentComplete ($i / @($targetSps).Count * 100)

    # Quién tiene asignada esta app (usuarios / grupos)
    Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $sp.Id -All | ForEach-Object {
        $label = "App: $($sp.DisplayName) [$(Get-AppRoleName $sp $_.AppRoleId)]"
        if ($_.PrincipalType -eq 'Group') {
            $g = Get-GroupInfo $_.PrincipalId
            if ($g) { foreach ($mId in $g.Members) { Add-Access $mId "$label (via group: $($g.Name))" } }
        } else {
            Add-Access $_.PrincipalId $label
        }
    }

    # Permisos de aplicación que tiene esta app sobre otras APIs (ej. Microsoft Graph)
    Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All | ForEach-Object {
        $res = $spById[$_.ResourceId]
        Add-Access $sp.Id "API Permission (Application): $($_.ResourceDisplayName)/$(Get-AppRoleName $res $_.AppRoleId)"
    }
}
Write-Progress -Activity 'Aplicaciones' -Completed

# Permisos delegados (consentimientos)
Get-MgOauth2PermissionGrant -All | ForEach-Object {
    $res   = $spById[$_.ResourceId]
    $who   = if ($_.ConsentType -eq 'AllPrincipals') { 'all users' } else { 'single user' }
    $label = "API Permission (Delegated, $who): $($res.DisplayName)/$(($_.Scope -split ' ' | Where-Object { $_ }) -join ',')"
    Add-Access $_.ClientId $label
}

# --- 5. Grupos (opcional) -----------------------------------------------------
if ($IncludeGroups) {
    Write-Host "Obteniendo grupos por usuario..." -ForegroundColor Cyan
    $i = 0
    foreach ($u in $users) {
        $i++
        Write-Progress -Activity 'Grupos' -Status $u.UserPrincipalName -PercentComplete ($i / @($users).Count * 100)
        Get-MgUserMemberOf -UserId $u.Id -All |
            Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.group' } |
            ForEach-Object { Add-Access $u.Id "Group: $($_.AdditionalProperties['displayName'])" }
    }
    Write-Progress -Activity 'Grupos' -Completed
}

# --- 6. Armar reporte ----------------------------------------------------------
Write-Host "Generando reporte..." -ForegroundColor Cyan
$report = [System.Collections.Generic.List[object]]::new()

foreach ($u in $users) {
    $type = if ($u.UserType -eq 'Guest') { 'User (Guest)' }
            elseif ($u.UserPrincipalName -match $ServiceAccountPattern -or
                    $u.DisplayName -match $ServiceAccountPattern) { 'Service Account' }
            else { 'User' }
    $report.Add([pscustomobject][ordered]@{
        'Account Type'          = $type
        'Account Name'          = $u.DisplayName
        'User Name'             = $u.UserPrincipalName
        'Email'                 = $u.Mail
        'Emp ID'                = $u.EmployeeId
        'SAML ID / UserID'      = $u.Id
        'Access/Roles'          = if ($access[$u.Id]) { $access[$u.Id] -join '; ' } else { 'Standard User (no admin roles)' }
        'Status'                = if ($u.AccountEnabled) { 'Active' } else { 'Disabled' }
        'EndPoint'              = $endpoint
        'Account Creation Date' = $u.CreatedDateTime
        'Last Sign-In'          = $u.SignInActivity.LastSignInDateTime
        'Source'                = if ($u.OnPremisesSyncEnabled) { 'Synced from on-prem AD' } else { 'Cloud' }
    })
}

foreach ($sp in $targetSps) {
    $type = switch ($true) {
        ($sp.ServicePrincipalType -eq 'ManagedIdentity')    { 'System (Managed Identity)'; break }
        ($sp.AppOwnerOrganizationId -eq $tenantId)          { 'Application (Internal)'; break }
        ($sp.AppOwnerOrganizationId -in $msTenants)         { 'Application (Microsoft)'; break }
        default                                             { 'Application (Third-party)' }
    }
    $report.Add([pscustomobject][ordered]@{
        'Account Type'          = $type
        'Account Name'          = $sp.DisplayName
        'User Name'             = $sp.DisplayName
        'Email'                 = ''
        'Emp ID'                = ''
        'SAML ID / UserID'      = $sp.AppId
        'Access/Roles'          = if ($access[$sp.Id]) { $access[$sp.Id] -join '; ' } else { 'None' }
        'Status'                = if ($sp.AccountEnabled) { 'Active' } else { 'Disabled' }
        'EndPoint'              = $endpoint
        'Account Creation Date' = $appCreated[$sp.AppId]
        'Last Sign-In'          = ''
        'Source'                = 'Cloud'
    })
}

$report | Sort-Object 'Account Type', 'Account Name' |
    Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

# --- Resumen -------------------------------------------------------------------
Write-Host "`nListo: $((Resolve-Path $OutputPath).Path)" -ForegroundColor Green
$report | Group-Object 'Account Type' | Sort-Object Name |
    Format-Table @{ n = 'Account Type'; e = { $_.Name } }, Count -AutoSize

Disconnect-MgGraph | Out-Null
