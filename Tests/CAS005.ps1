#requires -Modules ExchangeAnalyzer

<# 

Logic:
    Get CAS Namespace URIs (Get-ExchangeUrls)
    Test against internal DNS (native resolver)
    Test against external DNS (8.8.8.8 for example)

    Warn if >5m TTL
    Fail if >1h TTL
    Warn if unable to query external DNS

#>
# This test evaluates the TTLs on CAS namespace DNS records
Function Run-CAS005() {
    [CmdletBinding()]
    param()

    $TestID = "CAS005"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $WarningList = @()
    $InfoList = @()
    $ErrorList = @()

    # How low a TTL should be considered a warning?
    $warningTTL = New-TimeSpan -Minutes 5
    $failureTTL = New-TimeSpan -Hours 1
    
    # Array of known open DNS resolvers for external resolution purposes. 
    # This list should remain short so as to reduce delay in the event
    # that external resolution is blocked entirely and as such all resolvers
    # will fail.

    $externalResolvers = @(
        "8.8.8.8",        # Google
        "8.8.4.4",        # Google
        "208.67.222.222", # OpenDNS
        "208.67.220.220"  # OpenDNS
        )

    # Get all the namespacese we'll need to check, then dedupe it
    $NamespacesToCheck = @()
    $ExternalNamespacesToCheck = @()
    foreach ($CAS in $CASURLs) {
        # The System.Uri class allows us to format the string as a URI
        # and remove all the https:// and /owa parts
        $NamespacesToCheck += 
            $CAS.OAInternal,
            $CAS.OAExternal,
            ([System.Uri]$CAS.OWAInternal).Host,
            ([System.Uri]$CAS.OWAExternal).Host,
            ([System.Uri]$CAS.ECPInternal).Host,
            ([System.Uri]$CAS.ECPExternal).Host,
            ([System.Uri]$CAS.OABInternal).Host,
            ([System.Uri]$CAS.OABExternal).Host,
            ([System.Uri]$CAS.EWSInternal).Host,
            ([System.Uri]$CAS.EWSExternal).Host,
            ([System.Uri]$CAS.MAPIInternal).Host,
            ([System.Uri]$CAS.MAPIExternal).Host,
            ([System.Uri]$CAS.EASInternal).Host,
            ([System.Uri]$CAS.EASExternal).Host,
            ([System.Uri]$CAS.AutoDSCP).Host
        $ExternalNamespacesToCheck +=
            $CAS.OAExternal,
            ([System.Uri]$CAS.OWAExternal).Host,
            ([System.Uri]$CAS.ECPExternal).Host,
            ([System.Uri]$CAS.OABExternal).Host,
            ([System.Uri]$CAS.EWSExternal).Host,
            ([System.Uri]$CAS.MAPIExternal).Host
    }

    #Remove duplicates
    $NamespacesToCheck = @($NamespacesToCheck | Select -Unique)
    $ExternalNamespacesToCheck = @($ExternalNamespacesToCheck | Select -Unique)

    Write-Verbose "----- $TestID`: Found $($NamespacesToCheck.Count) namespaces ($($ExternalNamespacesToCheck.Count) external) to check:"

    # Test internal namespaces
    $resolvParameters = @{
        "DnsOnly" = $true;
        "NoHostsFile" = $true;
        "Type" = "A"
        }

    foreach ($namespace in $NamespacesToCheck) {
        Write-Verbose "----- $TestID`: Testing internal namespace '$namespace' against internal resolvers"
        
        $record = $null
        try {
            # Attempt to resolve the record using the default reoslvers
            $record = (Resolve-DnsName -Name $namespace @resolvParameters -Verbose:$false -ErrorAction Stop)[-1]
        } catch {
            # unable to resolve the record for some reason, either it 
            # doesn't exist or the resolver configuration is broken
            Write-Verbose "----- $TestID`: Unable to resolve namespace '$namespace' against internal resolvers"
            $ErrorList += "Internal: '$namespace': (Unable to resolve)"
        }

        if ($record) {
            $recordTTL = New-TimeSpan -seconds $record.TTL
            if ($recordTTL -ge $failureTTL) {
                # If the TTL of the record exceeds the failed trigger
                $FailedList += "$namespace' (Internal TTL: $($recordTTL.ToString()))"
            } elseif ($recordTTL -ge $warningTTL) {
                # Otherwise, if the TTL of the record exceeds the warning trigger
                $WarningList += "$namespace' (Internal TTL: $($recordTTL.ToString()))"
            } else {
                # TTL is OK
                $passedList += "$namespace' (Internal TTL: $($recordTTL.ToString()))"
            }
        }

    }


    # Test external namespaces
    $resolvParameters = @{
        "DnsOnly" = $true;
        "NoHostsFile" = $true;
        "Type" = "A";
        "Server" = $null;
        }
        
    # Find an external resolver that works
    #TODO: Find a better way of doing the resolver discovery below
    foreach ($resolver in $externalResolvers) {
        if ($resolvParameters.Server -eq $null) {
            # Test a resolver
            try {        
                if (Resolve-DnsName -Name "." -Server $resolver -Verbose:$false -ErrorAction Stop) {
                    $resolvParameters.Server = $resolver
                }
            } catch { 
                # Resolver didn't work 
            }
            
        }
    }

    if ($resolvParameters.Server -ne $null) {
        Write-Verbose "----- $TestID`: Using external resolver $($resolvParameters.Server)"
        foreach ($namespace in $ExternalNamespacesToCheck) {
            Write-Verbose "----- $TestID`: Testing external namespace '$namespace' against resolver $($resolvParameters.Server)"
        
            $record = $null
            try {
                # Attempt to resolve the record using the default reoslvers
                $record = (Resolve-DnsName -Name $namespace @resolvParameters -Verbose:$false -ErrorAction Stop)[-1]
            } catch {
                # unable to resolve the record for some reason, either it 
                # doesn't exist or the resolver configuration is broken
                Write-Verbose "----- $TestID`: Unable to resolve namespace '$namespace' against external resolvers"
                $ErrorList += "External: '$namespace': (Unable to resolve)"
            }

            if ($record) {
                $recordTTL = New-TimeSpan -seconds $record.TTL
                if ($recordTTL -ge $failureTTL) {
                    # If the TTL of the record exceeds the failed trigger
                    $FailedList += "$namespace (External TTL: $($recordTTL.ToString()))"
                } elseif ($recordTTL -ge $warningTTL) {
                    $WarningList += "$namespace (External TTL: $($recordTTL.ToString()))"
                } else {
                    $PassedList += "$namespace (External TTL: $($recordTTL.ToString()))"
                }
            }

        }
    } else {
        Write-Verbose "----- $TestID`: Could not find a functional external resolver"
        $ErrorList += "External DNS resolution failed"
    }

    #Roll the object to be returned to the results
    $ReportObj = Get-TestResultObject -ExchangeAnalyzerTests $ExchangeAnalyzerTests `
                                      -TestId $TestID `
                                      -PassedList $PassedList `
                                      -FailedList $FailedList `
                                      -WarningList $WarningList `
                                      -InfoList $InfoList `
                                      -ErrorList $ErrorList `
                                      -Verbose:($PSBoundParameters['Verbose'] -eq $true)

    return $ReportObj
}

Run-CAS005