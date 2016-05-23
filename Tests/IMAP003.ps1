Function Run-IMAP003()
{
    [CmdletBinding()]
    param()

    $TestID = "IMAP003"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $WarningList = @()
    $InfoList = @()
    $ErrorList = @()

    #Check IMAP settings for ProtocolLogEnabled configuration. Protocol logging for IMAP
    #does not have age/retention settings, so it can grow to consume all available disk
    #space on the server. Because external cleanup options are possible, eg running a
    #script as a scheduled task, if IMAP protocol logging is enabled it is considered
    #a warning not a fail.

    foreach ($IMAPSetting in $AllIMAPSettings)
    {
        Write-Verbose "Checking IMAP protocol logging settings for $($IMAPSetting.Server)"
        
        switch ($IMAPSetting.ProtocolLogEnabled -eq $true)
        {
            $true {
                $tmpString = "$($IMAPSetting.Server) has protocol logging enabled"
                Write-Verbose $tmpString
                $WarningList += $IMAPSetting.Server
                }
            $false {
                $tmpString = "$($IMAPSetting.Server) has protocol logging disabled"
                Write-Verbose $tmpString
                $PassedList += $IMAPSetting.Server
                }
            default {
                $ErrorList += "$($IMAPSetting.Server) protocol log setting could not be determined"
                }
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

Run-IMAP003

