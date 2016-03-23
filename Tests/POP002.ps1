#This is your test
Function Run-POP002()
{
    [CmdletBinding()]
    param()

    $TestID = "POP002"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $WarningList = @()
    $InfoList = @()
    $ErrorList = @()

    #Check POP settings for SecureLogin

    foreach ($PopSetting in $AllPopSettings)
    {
        Write-Verbose "Checking POP settings for $($PopSetting.Server)"
        
        if ($PopSetting.LoginType -ieq "SecureLogin")
        {
            Write-Verbose "$($PopSetting) requires secure login"
            $PassedList += $PopSetting.Server
        }
        else
        {
            #SecureLogin is not enabled, but server is secure if no port binding exists for
            #unencrypted connections, therefore forcing connections on the SSL port only
            
            if ($PopSetting.UnencryptedOrTLSBindings.Count -gt 0)
            {
                $tmpString = "$($PopSetting.Server) allows plain text login over insecure ports"
                Write-Verbose $tmpString
                $FailedList += $tmpString
            }
            else
            {
                $tmpString = "$($PopSetting.Server) allows plain text login, but has no insecure port bindings"
                Write-Verbose $tmpString
                $WarningList += $tmpString
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

Run-POP002

