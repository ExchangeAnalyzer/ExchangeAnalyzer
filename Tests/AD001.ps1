#requires -Modules ExchangeAnalyzer
#requires -Modules ActiveDirectory

#This function verifies the Active Directory Domain level is at the correct level
Function Run-AD001()
{
    [CmdletBinding()]
    param()

    $TestID = "AD001"
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

    Write-Verbose "$($domains.count) domain(s) found."

    # Check for other forest domains (via trusts)
    If($? -and $Domains -ne $Null) {
        Write-Verbose "Checking for domain trusts to other forests."
        ForEach($Domain in $Domains) { 
            # Get list of AD Domain Trusts in each domain
            $ADDomainTrusts = Get-ADObject -Filter {ObjectClass -eq "trustedDomain"} -Server $Domain -Properties * -EA 0
            If($? -and $ADDomainTrusts -ne $Null) {
                Write-Verbose "Domain trusts found."
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
            else
            {
                Write-Verbose "No domain trusts found."
            }
        }
    }
    else
    {
        Write-Verbose "An error occurred or no domains were found."
    }

    foreach ($server in $exchangeservers) { 
        $admin = $server.admindisplayversion
        Write-Verbose $admin
        #[string]$ver=[string]$admin.major+'.'+[string]$admin.minor
        #Write-Verbose $ver
        if ($admin -like "Version 15.0*")
        {
            $Ex2013 = $true
            Write-Verbose "Exchange 2013 detected."    
        }
        if ($admin -like "Version 15.1*")
        {
            $Ex2016 = $true
            Write-Verbose "Exchange 2016 detected."
        }
    }
    
    if ($ex2013 -eq $true) {
    Write-Verbose "At least one Exchange 2013 server detected."
    # All Exchange 2013 servers, no Exchange 2016 servers found
        foreach ($domain in $alldomains) {
            $pdc = (get-addomain $domain).pdcemulator
            Write-Verbose "Using PDCE $pdc"
            $dse = ([ADSI] "LDAP://$pdc/RootDSE")
            $dlevel = $dse.domainFunctionality
            $flevel = $dse.forestFunctionality
            if ($dlevel -ge "2") {
                $PassedList += $($domain)
            }
            if (($dlevel -lt "2") -and ($dlevel -gt "0")) {
                $FailedList += $($domain)
            }
        }
    }

    if ($ex2016 -eq $true) {
    Write-Verbose "At least one Exchange 2016 server detected."
    # Exchange 2016 servers found
        foreach ($domain in $alldomains) {
            $pdc = (get-addomain $domain).pdcemulator
            Write-Verbose "Using PDCE $pdc"
            $dse = ([ADSI] "LDAP://$pdc/RootDSE")
            $dlevel = $dse.domainFunctionality
            $flevel = $dse.forestFunctionality
            if ($dlevel -ge "3") {
                $PassedList += $($domain)
            }
            if (($dlevel -lt "3") -and ($dlevel -gt "0")) {
                $FailedList += $($domain)
            }
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

Run-AD001