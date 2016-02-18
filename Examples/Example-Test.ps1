#This is your test
Function Run-TESTID()
{
    [CmdletBinding()]
    param()

    $TestID = "TESTID"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $WarningList = @()
    $InfoList = @()
    $ErrorList = @()

    #Your test logic goes here and populates the results
    $PassedList += "Foo"
    $FailedList += "Bar" #An empty failed list will result in a test passing.
    #Add any errors to $ErrorList

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

Run-TESTID

