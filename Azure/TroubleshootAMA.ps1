$servers = @'
Sub,RG,Name,Provider
634e0965-c22c-4e55-ba34-7d320d396e24,euaew_shared_dfs_prod_rg,EUAEW2WANFCORE01,Microsoft.Compute/virtualMachines
f8a7128f-5ef7-466b-97d2-91abe3c6a4a5,euaew_nsap_scm_tdq_rg,EUAEW2WSCMTDQ01,Microsoft.Compute/virtualMachines
f8a7128f-5ef7-466b-97d2-91abe3c6a4a5,euaew_nsap_sectigotestservers_tst_rg,EUAEW2WSTSAPPT1,Microsoft.Compute/virtualMachines
634e0965-c22c-4e55-ba34-7d320d396e24,inf_ops_arc_hh_emea,EUD0012WDIP01,Microsoft.HybridCompute/machines
634e0965-c22c-4e55-ba34-7d320d396e24,inf_ops_arc_hh_emea,EUD0012WDIP02,Microsoft.HybridCompute/machines
634e0965-c22c-4e55-ba34-7d320d396e24,inf_ops_arc_hh_emea,EUD0012WPRT01,Microsoft.HybridCompute/machines
634e0965-c22c-4e55-ba34-7d320d396e24,inf_ops_arc_emea,EUEQX2WKTBDC01,Microsoft.HybridCompute/machines
634e0965-c22c-4e55-ba34-7d320d396e24,inf_ops_arc_namr,LAM0092WKTBDIP1,Microsoft.HybridCompute/machines
634e0965-c22c-4e55-ba34-7d320d396e24,naae2_nsap_fpass_prod_rg,NAAE21WFPASSWEB,Microsoft.Compute/virtualMachines
634e0965-c22c-4e55-ba34-7d320d396e24,naae2_nsap_complianceguardian_prd_rg,NAAE22WACGDBSP1,Microsoft.Compute/virtualMachines
634e0965-c22c-4e55-ba34-7d320d396e24,naae2_nsap_apcmp_prod,NAAE22WCMPDBS1,Microsoft.Compute/virtualMachines
634e0965-c22c-4e55-ba34-7d320d396e24,naae2_nsap_leebartender_prd_rg,NAAE22WLBADBSP1,Microsoft.Compute/virtualMachines
f8a7128f-5ef7-466b-97d2-91abe3c6a4a5,naae2_nsap_leebartender_tdq_rg,NAAE22WLBADBST1,Microsoft.Compute/virtualMachines
634e0965-c22c-4e55-ba34-7d320d396e24,naae2_nsap_leecrd_prd_rg,NAAE22WLCRDBSP1,Microsoft.Compute/virtualMachines
f8a7128f-5ef7-466b-97d2-91abe3c6a4a5,naae2_nsap_leecrd_tdq_rg,NAAE22WLCRDBST1,Microsoft.Compute/virtualMachines
f8a7128f-5ef7-466b-97d2-91abe3c6a4a5,naae2_nsap_logilitylee_dev_rg,NAAE22WLLGDBSD1,Microsoft.Compute/virtualMachines
f8a7128f-5ef7-466b-97d2-91abe3c6a4a5,naae2_nsap_logilitylee_dev_rg,NAAE22WLLGDBSD2,Microsoft.Compute/virtualMachines
634e0965-c22c-4e55-ba34-7d320d396e24,naae2_nsap_logilitylee_prd_rg,NAAE22WLLGDBSP1,Microsoft.Compute/virtualMachines
634e0965-c22c-4e55-ba34-7d320d396e24,naae2_nsap_logilitylee_prd_rg,NAAE22WLLGDBSP2,Microsoft.Compute/virtualMachines
634e0965-c22c-4e55-ba34-7d320d396e24,naae2_nsap_continuitypatrol_prd_rg,NAAE22WPCPDBSP1,Microsoft.Compute/virtualMachines
634e0965-c22c-4e55-ba34-7d320d396e24,naae2_nsap_dbasbx_prd_rg,NAAE22WPOCDBSP3,Microsoft.Compute/virtualMachines
f8a7128f-5ef7-466b-97d2-91abe3c6a4a5,naae2_nsap_dbasbx_tdq_rg,NAAE22WPOCDBST1,Microsoft.Compute/virtualMachines
f8a7128f-5ef7-466b-97d2-91abe3c6a4a5,naae2_nsap_dbasbx_tdq_rg,NAAE22WPOCDBST2,Microsoft.Compute/virtualMachines
634e0965-c22c-4e55-ba34-7d320d396e24,naae2_nsap_sdpmagicbox_prd_rg,NAAE22WSDPDBSP1,Microsoft.Compute/virtualMachines
f8a7128f-5ef7-466b-97d2-91abe3c6a4a5,naae2_nsap_sdpmagicbox_tdq_rg,NAAE22WSDPDBST1,Microsoft.Compute/virtualMachines
634e0965-c22c-4e55-ba34-7d320d396e24,naae2_nsap_continuitypatrol_drs_rg,NAASC2WPCPWEBP1,Microsoft.Compute/virtualMachines
85a526d2-43e0-475d-9e93-a993bc3d5215,naasc_nsap_continuitypatrol_drs_rg,NAASC2WPCPWEBP1,Microsoft.Compute/virtualMachines
'@ | ConvertFrom-Csv

$currentSub = $null
$report = foreach ($s in $servers) {
    if ($s.Sub -ne $currentSub) {
        Set-AzContext -SubscriptionId $s.Sub -WarningAction SilentlyContinue | Out-Null
        $currentSub = $s.Sub
    }
    $id = "/subscriptions/$($s.Sub)/resourceGroups/$($s.RG)/providers/$($s.Provider)/$($s.Name)"

    # Extensión AMA: estado + identidad configurada + proxy
    $amaState = 'NO-AMA'; $idType = ''; $idValue = ''; $proxy = ''
    try {
        $ext = Get-AzResource -ResourceId "$id/extensions/AzureMonitorWindowsAgent" -ExpandProperties -ErrorAction Stop
        $amaState = $ext.Properties.provisioningState
        $mi       = $ext.Properties.settings.authentication.managedIdentity
        $idType   = $mi.'identifier-name'
        $idValue  = if ($mi.'identifier-value') { ($mi.'identifier-value' -split '/')[-1] } else { '<VACIO>' }
        $proxy    = $ext.Properties.settings.proxy.address
    } catch { }

    # Asociaciones DCR
    $dcrs = try {
        (Get-AzDataCollectionRuleAssociation -ResourceUri $id -ErrorAction Stop |
            Where-Object DataCollectionRuleId |
            ForEach-Object { ($_.DataCollectionRuleId -split '/')[-1] }) -join ', '
    } catch { 'ERROR' }

    [pscustomobject]@{
        Server      = $s.Name
        Sub         = $s.Sub.Substring(0,4)
        Type        = if ($s.Provider -like '*Hybrid*') { 'Arc' } else { 'VM' }
        AMA         = $amaState
        IdentityBy  = $idType
        Identity    = $idValue
        Proxy       = $proxy
        DCRs        = if ($dcrs) { $dcrs } else { 'NONE' }
        SentinelDCR = $dcrs -match 'KTB-AMA-SecuriyEvents-Sentinel'
    }
}

$report | Format-Table -AutoSize
$report | Export-Csv ./ama_validation.csv -NoTypeInformation