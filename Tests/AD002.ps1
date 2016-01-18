#requires -Modules ExchangeAnalyzer
#requires -Modules ActiveDirectory

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

    $# Domain Check - Current Forest
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

    foreach ($server in $exchangeservers) { 
        $admin = $server.admindisplayversion
        [string]$ver=[string]$admin.major+'.'+[string]$admin.minor
        if ($Ver -like "15.0") {$Ex2013 = $true}
        if ($Ver -like "15.1") {$Ex2016 = $true}
    }

    if ($ex2013 -eq $true) {
    # All Exchange 2013 servers, no Exchange 2016 servers found
        Foreach ($forest in $allforests) {
            $DC = (get-adforest $forest).GlobalCatalogs
            $dse = ([ADSI] "LDAP://$dc/RootDSE")
            $flevel = $dse.forestFunctionality
            if ($flevel -ge "2") {
                $PassedList += $($ADforest)
            }
            if (($flevel -lt "2") -and ($flevel -gt "0")) {
                $FailedList += $($ADforest)
            }
        }
    }

    if ($ex2016 -eq $true) {
    # Exchange 2016 servers found
         Foreach ($forest in $allforests) {
            $DC = (get-adforest $forest).GlobalCatalogs
            $dse = ([ADSI] "LDAP://$dc/RootDSE")
            $flevel = $dse.forestFunctionality
            if ($flevel -ge "3") {
                $PassedList += $($Forest)
            }
            if (($flevel -lt "3") -and ($flevel -gt "0")) {
                $FailedList += $($Forest)
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

Run-AD002