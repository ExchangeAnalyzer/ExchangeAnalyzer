#requires -Modules ExchangeAnalyzer

#This function tests each Exchange site to determine whether more than one CAS URL/namespace
#exists for each HTTPS service.
Function Run-CAS001()
{
    [CmdletBinding()]
    param()

    $TestID = "CAS001"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $ErrorList = @()

    $sites = @($ClientAccessServers | Group-Object -Property:Site | Select Name)

    # Get the URLs for each site, and if more than one URL exists for a HTTPS service in the same site it is
    # considered a fail.
    foreach ($site in $sites)
    {
        $SiteName = ($Site.Name).Split("/")[-1]

        Write-Verbose "Processing $SiteName"

        $SiteOWAInternalUrls = @()
        $SiteOWAExternalUrls = @()

        $SiteECPInternalUrls = @()
        $SiteECPExternalUrls = @()

        $SiteOABInternalUrls = @()
        $SiteOABExternalUrls = @()
    
        $SiteRPCInternalUrls = @()
        $SiteRPCExternalUrls = @()

        $SiteEWSInternalUrls = @()
        $SiteEWSExternalUrls = @()

        $SIteMAPIInternalUrls = @()
        $SiteMAPIExternalUrls = @()

        $SiteActiveSyncInternalUrls = @()
        $SiteActiveSyncExternalUrls = @()

        $SiteAutodiscoverUrls = @()

        $CASinSite = @($ClientAccessServers | Where {$_.Site -eq $site.Name})

        Write-Verbose "Getting OWA Urls for site $SiteName"
        foreach ($CAS in $CASinSite)
        {
            Write-Verbose "Server: $($CAS.Name)"
            $CASOWAUrls = @($CASURLs | Where {$_.Name -ieq $CAS.Name} | Select OWAInternal,OWAExternal)
            foreach ($CASOWAUrl in $CASOWAUrls)
            {
                if (!($SiteOWAInternalUrls -Contains $CASOWAUrl.OWAInternal.ToLower()) -and ($CASOWAUrl.OWAInternal.ToLower() -ne $null))
                {
                    $SiteOWAInternalUrls += $CASOWAUrl.OWAInternal.ToLower()
                }
                if (!($SiteOWAExternalUrls -Contains $CASOWAUrl.OWAExternal.ToLower()) -and ($CASOWAUrl.OWAExternal.ToLower() -ne $null))
                {
                    $SiteOWAExternalUrls += $CASOWAUrl.OWAExternal.ToLower()
                }
            }
        }

        if ($SiteOWAInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteOWAExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting ECP Urls for site $SiteName"
        foreach ($CAS in $CASinSite)
        {
            $CASECPUrls = @($CASURLs | Where {$_.Name -ieq $CAS.Name} | Select ECPInternal,ECPExternal)
            foreach ($CASECPUrl in $CASECPUrls)
            {
                if (!($SiteECPInternalUrls -Contains $CASECPUrl.ECPInternal.ToLower()) -and ($CASECPUrl.ECPInternal.ToLower() -ne $null))
                {
                    $SiteECPInternalUrls += $CASECPUrl.ECPInternal.ToLower()
                }
                if (!($SiteECPExternalUrls -Contains $CASECPUrl.ECPInternal.ToLower()) -and ($CASECPUrl.ECPInternal.ToLower() -ne $null))
                {
                    $SiteECPExternalUrls += $CASECPUrl.ECPExternal.ToLower()
                }
            }
        }

        if ($SiteECPInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteECPExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting OAB Urls for site $SiteName"
        foreach ($CAS in $CASinSite)
        {
            $CASOABUrls = @($CASURLs | Where {$_.Name -ieq $CAS.Name} | Select OABInternal,OABExternal)
            foreach ($CASOABUrl in $CASOABUrls)
            {
                if (!($SiteOABInternalUrls -Contains $CASOABUrl.OABInternal.ToLower()) -and ($CASOABUrl.OABInternal.ToLower() -ne $null))
                {
                    $SiteOABInternalUrls += $CASOABUrl.OABInternal.ToLower()
                }
                if (!($SiteOABExternalUrls -Contains $CASOABUrl.OABExternal.ToLower()) -and ($CASOABUrl.OABExternal.ToLower() -ne $null))
                {
                    $SiteOABExternalUrls += $CASOABUrl.OABExternal.ToLower()
                }
            }
        }

        if ($SiteOABInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteOABExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting RPC Urls for site $SiteName"
        foreach ($CAS in $CASinSite)
        {
            $OA = @($CASURLs | Where {$_.Name -ieq $CAS.Name} | Select OAInternal,OAExternal)
            [string]$OAInternalHostName = $OA.OAInternal
            [string]$OAExternalHostName = $OA.OAExternal

            [string]$OAInternalUrl = "https://$($OAInternalHostName.ToLower())/rpc"
            [string]$OAExternalUrl = "https://$($OAExternalHostName.ToLower())/rpc"

            if (!($SiteRPCInternalUrls -Contains $OAInternalUrl) -and ($OAInternalHostName -ne $null))
            {
                $SiteRPCInternalUrls += $OAInternalUrl
            }
            if (!($SiteRPCExternalUrls -Contains $OAExternalUrl) -and ($OAExternalHostName -ne $null) -and ($OAExternalHostName -ne ""))
            {
                $SiteRPCExternalUrls += $OAExternalUrl
            }
        }

        if ($SiteRPCInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteRPCExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting EWS Urls for site $SiteName"
        foreach ($CAS in $CASinSite)
        {
            $CASEWSUrls = @($CASURLs | Where {$_.Name -ieq $CAS.Name} | Select EWSInternal,EWSExternal)
            foreach ($CASEWSUrl in $CASEWSUrls)
            {
                if (!($SiteEWSInternalUrls -Contains $CASEWSUrl.EWSInternal.ToLower()) -and ($CASEWSUrl.EWSInternal.ToLower() -ne $null))
                {
                    $SiteEWSInternalUrls += $CASEWSUrl.EWSInternal.ToLower()
                }
                if (!($SiteEWSExternalUrls -Contains $CASEWSUrl.EWSExternal.ToLower()) -and ($CASEWSUrl.EWSExternal.ToLower() -ne $null))
                {
                    $SiteEWSExternalUrls += $CASEWSUrl.EWSExternal.ToLower()
                }
            }
        }

        if ($SiteEWSInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteEWSExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting MAPI Urls for site $SiteName"
        foreach ($CAS in $CASinSite)
        {
            $CASMAPIUrls = @($CASURLs | Where {$_.Name -ieq $CAS.Name} | Select MAPIInternal,MAPIExternal)
            foreach ($CASMAPIUrl in $CASMAPIUrls)
            {
                if (!($SiteMAPIInternalUrls -Contains $CASMAPIUrl.MAPIInternal.ToLower()) -and ($CASMAPIUrl.MAPIInternal.ToLower() -ne $null))
                {
                    $SiteMAPIInternalUrls += $CASMAPIUrl.MAPIInternal.ToLower()
                }
                if (!($SiteMAPIExternalUrls -Contains $CASMAPIUrl.MAPIExternal.ToLower()) -and ($CASMAPIUrl.MAPIExternal.ToLower() -ne $null))
                {
                    $SiteMAPIExternalUrls += $CASMAPIUrl.MAPIExternal.ToLower()
                }
            }
        }

        if ($SiteMAPIInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteMAPIExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting ActiveSync Urls for site $SiteName"
        foreach ($CAS in $CASinSite)
        {
            $CASActiveSyncUrls = @($CASURls | Where {$_.Name -eq $CAS.Name} | Select EASInternal,EASExternal)
            foreach ($CASActiveSyncUrl in $CASActiveSyncUrls)
            {
                if (!($SiteActiveSyncInternalUrls -Contains $CASActiveSyncUrl.EASInternal.ToLower()) -and ($CASActiveSyncUrl.EASInternal.ToLower() -ne $null))
                {
                    $SiteActiveSyncInternalUrls += $CASActiveSyncUrl.EASInternal.ToLower()
                }
                if (!($SiteActiveSyncExternalUrls -Contains $CASActiveSyncUrl.EASExternal.ToLower()) -and ($CASActiveSyncUrl.EASExternal.ToLower() -ne $null))
                {
                    $SiteActiveSyncExternalUrls += $CASActiveSyncUrl.EASExternal.ToLower()
                }
            }
        }

        if ($SiteActiveSyncInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteActiveSyncExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting AutoDiscover Urls for site $SiteName"
        foreach ($CAS in $CASinSite)
        {
            #$CASServer = Get-ClientAccessServer $CAS.Name
            $AutoDUrl = @($CASURLs | Where {$_.Name -ieq $CAS.Name} | Select AutoD)
            [string]$AutodiscoverSCP = $AutoDUrl
            $CASAutodiscoverUrl = $AutodiscoverSCP.Replace("/Autodiscover.xml","")
            if (!($SiteAutodiscoverUrls -Contains $CASAutodiscoverUrl.ToLower())) {$SiteAutodiscoverUrls += $CASAutodiscoverUrl.ToLower()}
        }

        if ($SiteAutodiscoverUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        #If the site is not in FailedList by now, add it to $PassedList
        if ($FailedList -notcontains $SiteName) { $PassedList += $SiteName }
    }

    #Roll the object to be returned to the results
    $ReportObj = Get-TestResultObject -ExchangeAnalyzerTests $ExchangeAnalyzerTests `
                                      -TestId $TestID `
                                      -PassedList $PassedList `
                                      -FailedList $FailedList `
                                      -ErrorList $ErrorList `
                                      -Verbose:($PSBoundParameters['Verbose'] -eq $true)

    return $ReportObj
}

Run-CAS001