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
            $POPServices = @("MSExchangePop3","MSExchangePOP3BE")
            foreach ($POPService in $POPServices)
            {
                $Service = Get-ExAWMIObject -Computer $Server -Class Win32_Service -Filter "Name='$($POPService)'"

                #Storing values in property bags may not be necessary, as this test information is
                #not re-used elsewhere. May be possible to remove this later if speed is impacted.
                Set-ExAServerProperty -Server $Server -Property "$($Service)State" -Value $Service.State
                Set-ExAServerProperty -Server $Server -Property "$($Service)StartMode" -Value $Service.StartMode               
                
                $tmpString = "$($Server): $($Service.Name) is $($Service.State) (Start Mode: $($Service.StartMode))"
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

