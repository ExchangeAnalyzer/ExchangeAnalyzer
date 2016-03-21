#This is your test
Function Run-POP001()
{
    [CmdletBinding()]
    param()

    $TestID = "POP001"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $WarningList = @()
    $InfoList = @()
    $ErrorList = @()

    $POPServers = @($ExchangeServers | Where {$_.IsClientAccessServer -or $_.IsMailboxServer})

    foreach ($Server in $POPServers)
    {
        Write-Verbose "Checking POP services for $($Server)"

        try
        {
            #This test won't return StartupType. Need to replace with a WMI query once the WMI framework
            #has been built into the ExchangeAnalyzer module.
            $PopServices = @(Get-Service -ComputerName $Server MSExchangePOP* -ErrorAction STOP)
            foreach ($PopService in $PopServices)
            {
                $tmpString = "$($Server): $($PopService.DisplayName) is $($PopService.Status)"
                Write-Verbose $tmpString
                $InfoList += $tmpString
            }
        }
        catch
        {
            Write-Verbose "Unable to determine POP service status"
            Write-Verbose $_.Exception.Message

            $ErrorList += "$($Server) - unable to determine POP service status."
        }
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

Run-POP001

