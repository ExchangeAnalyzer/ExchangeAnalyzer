#requires -Modules ExchangeAnalyzer

#This function verifies the Active Directory Domain level is Windows 2008 or greater
Function Run-AD001()
{
    [CmdletBinding()]
    param()

    $TestID = "AD001"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $ErrorList = @()

    $domaindnsname = ($addomain).dnsroot
    $domainmode = ($addomain).domainmode

    if (($domainmode -like "*2012*") -or ($domainmode -like "*2008*")) {
        $PassedList += $($domaindnsname)
    } else {
        $FailedList += $($domaindnsname)
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