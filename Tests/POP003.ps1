#This is your test
Function Run-POP003()
{
    [CmdletBinding()]
    param()

    $TestID = "POP003"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $WarningList = @()
    $InfoList = @()
    $ErrorList = @()

    #Check POP settings for ProtocolLogEnabled configuration. Protocol logging for POP
    #does not have age/retention settings, so it can grow to consume all available disk
    #space on the server. Because external cleanup options are possible, eg running a
    #script as a scheduled task, if POP protocol logging is enabled it is considered
    #a warning not a fail.

    foreach ($PopSetting in $AllPopSettings)
    {
        Write-Verbose "Checking POP protocol logging settings for $($PopSetting.Server)"
        
        switch ($PopSetting.ProtocolLogEnabled -eq $true)
        {
            $true {
                $tmpString = "$($PopSetting.Server) has protocol logging enabled"
                Write-Verbose $tmpString
                $WarningList += $PopSetting.Server
                }
            $false {
                $tmpString = "$($PopSetting.Server) has protocol logging disabled"
                Write-Verbose $tmpString
                $PassedList += $PopSetting.Server
                }
            default {
                $ErrorList += "$($PopSetting.Server) protocol log setting could not be determined"
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

Run-POP003

