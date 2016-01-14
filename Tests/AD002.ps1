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

    $forestname = ($adforest).name
    $forestmode = ($adforest).forestmode

    if (($forestmode -like "*2012*") -or ($forestmode -like "*2008*")) {
        $PassedList += $($forestname)
    } else {
        $FailedList += $($forestname)
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