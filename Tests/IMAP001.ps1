Function Run-IMAP001()
{
    [CmdletBinding()]
    param()

    $TestID = "IMAP001"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $WarningList = @()
    $InfoList = @()
    $ErrorList = @()

    $IMAPServers = @($ExchangeServers | Where {$_.IsClientAccessServer -or $_.IsMailboxServer})

    foreach ($Server in $IMAPServers)
    {
        Write-Verbose "Checking IMAP services for $($Server)"

        try
        {
            #This test won't return StartupType. Need to replace with a WMI query once the WMI framework
            #has been built into the ExchangeAnalyzer module.
            $IMAPServices = @(Get-Service -ComputerName $Server MSExchangeIMAP* -ErrorAction STOP)
            foreach ($IMAPService in $IMAPServices)
            {
                $tmpString = "$($Server): $($IMAPService.DisplayName) is $($IMAPService.Status)"
                Write-Verbose $tmpString
                $InfoList += $tmpString
            }
        }
        catch
        {
            Write-Verbose "Unable to determine IMAP service status"
            Write-Verbose $_.Exception.Message

            $ErrorList += "$($Server) - unable to determine IMAP service status."
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

Run-IMAP001

