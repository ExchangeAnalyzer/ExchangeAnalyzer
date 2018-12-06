#requires -Modules ExchangeAnalyzer

#This function verifies the Active Directory Forest level is Windows 2008 or greater
Function Run-AD002()
{
    [CmdletBinding()]
    param()

    $TestID = "AD002"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $ErrorList = @()

    # Domain Check - Current Forest
    $Domains = @()
    $Forest = @()
    $Domains = @((get-adforest).domains)
    $Forest = @((get-adforest).name)
    $AllDomains = $domains
    $AllForests = $forest

    # Check for other forest domains (via trusts)
    If($? -and $Domains -ne $Null) {
        ForEach($Domain in $Domains) { 
            # Get list of AD Domain Trusts in each domain
            $ADDomainTrusts = Get-ADObject -Filter {ObjectClass -eq "trustedDomain"} -Server $Domain -Properties * -EA 0
            If($? -and $ADDomainTrusts -ne $Null) {
                If($ADDomainTrusts -is [array]) {
                    [int]$ADDomainTrustsCount = $ADDomainTrusts.Count 
                } Else {
                    [int]$ADDomainTrustsCount = 1
                }
                ForEach($Trust in $ADDomainTrusts) { 
                    [string]$TrustName = $Trust.Name
                    If ($TrustName -ne $Forests) {
                        $TrustAttributesNumber = $Trust.TrustAttributes
                        if (($TrustAttributesNumber -eq "8")) {
                            $newdomains = (get-adforest $trustname).domains
                            $newforest = (get-adforest $trustname).name
                            $alldomains += $newdomains
                            $allforests += $newforest
                        }
                    }
                }
            }
        }
    }

    #Determine newest and oldest Exchange versions in the org and set min/max functional levels
    #based on supportability matrix: https://technet.microsoft.com/library/ff728623(v=exchg.150).aspx

    $ExchangeVersions = @{
                        Newest = ($ExchangeServers | Sort AdminDisplayVersion -Descending)[0].AdminDisplayVersion
                        Oldest = ($ExchangeServers | Sort AdminDisplayVersion -Descending)[-1].AdminDisplayVersion
                        }

    if ($ExchangeVersions.Newest -like "Version 15.2*")
    {
        $MinFunctionalLevel = 6
        $MinFunctionalLevelText = "Windows Server 2012 R2"
    }
    elseif ($ExchangeVersions.Newest -like "Version 15.1*")
    {
        $MinFunctionalLevel = 3
        $MinFunctionalLevelText = "Windows Server 2008"
    }
    else
    {
        $MinFunctionalLevel = 2
        $MinFunctionalLevelText = "Windows Server 2003"
    }

    if ($ExchangeVersions.Oldest -like "Version 8.0*")
    {
        $MaxFunctionalLevel = 5
        $MaxFunctionalLevelText = "Windows Server 2012"
    }
    elseif ($ExchangeVersions.Oldest -like "Version 15.0*")
    {
        $MaxFunctionalLevel = 6
        $MaxFunctionalLevelText = "Windows Server 2012 R2"
    }
    else
    {
        $MaxFunctionalLevel = 7
        $MaxFunctionalLevelText = "Windows Server 2016"
    }

    Write-Verbose "The Forest Functional level must be:"
    Write-Verbose " - Minimum: $MinFunctionalLevelText"
    Write-Verbose " - Maximum: $MaxFunctionalLevelText"

    foreach ($forest in $allforests)
    {
        $DC = @((get-adforest $forest).GlobalCatalogs)[0]
        Write-Verbose "Using GC $DC"
        $dse = ([ADSI] "LDAP://$dc/RootDSE")
        $flevel = $dse.forestFunctionality

        switch ($flevel)
        {
            2 {$fleveltext = "Windows Server 2003"}
            3 {$fleveltext = "Windows Server 2008"}
            4 {$fleveltext = "Windows Server 2008 R2"}
            5 {$fleveltext = "Windows Server 2012"}
            6 {$fleveltext = "Windows Server 2012 R2"}
            7 {$fleveltext = "Windows Server 2016"}
        }

        if ($flevel -ge $MinFunctionalLevel -and $dlevel -le $MaxFunctionalLevel)
        {
            $PassedList += "$($forest) ($fleveltext)"
        }
        else
        {
            $FailedList += "$($forest) ($fleveltext)"
        }
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

Run-AD002