#requires -Modules ExchangeAnalyzer

#This function tests whether Exchange 2010 or earlier servers exist and flags as a 
#warning in the Exchange Analyzer report.
Function Run-EXSRV001()
{
    [CmdletBinding()]
    param()

    $TestID = "EXSRV001"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $ErrorList = @()

    $SupportedServers = @($ExchangeServers | Where {$_.AdminDisplayVersion -like "Version 15.*"})
    $UnsupportedServers = @($ExchangeServers | Where {$_.AdminDisplayVersion -notlike "Version 15.*"})

    if ($SupportedServers.Count -gt 0)
    {
        foreach ($SupportedServer in $SupportedServers)
        {
            Write-Verbose "$($SupportedServer) is supported by ExchangeAnalyzer"
            $PassedList += $($SupportedServer.Name)
        }
    }
    
    if ($UnsupportedServers.Count -gt 0)
    {
        foreach ($UnsupportedServer in $UnsupportedServers)
        {
            Write-Verbose "$($UnsupportedServer) is not supported by ExchangeAnalyzer"
            $FailedList += $($UnsupportedServer.Name)
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

Run-EXSRV001


