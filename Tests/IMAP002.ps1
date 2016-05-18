Function Run-IMAP002()
{
    [CmdletBinding()]
    param()

    $TestID = "IMAP002"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $WarningList = @()
    $InfoList = @()
    $ErrorList = @()

    #Check IMAP settings for SecureLogin

    foreach ($IMAPSetting in $AllIMAPSettings)
    {
        Write-Verbose "Checking IMAP settings for $($IMAPSetting.Server)"
        
        if ($IMAPSetting.LoginType -ieq "SecureLogin")
        {
            Write-Verbose "$($IMAPSetting) requires secure login"
            $PassedList += $IMAPSetting.Server
        }
        else
        {
            #SecureLogin is not enabled, but server is secure if no port binding exists for
            #unencrypted connections, therefore forcing connections on the SSL port only
            
            if ($IMAPSetting.UnencryptedOrTLSBindings.Count -gt 0)
            {
                $tmpString = "$($IMAPSetting.Server) allows plain text login over insecure ports"
                Write-Verbose $tmpString
                $FailedList += $tmpString
            }
            else
            {
                $tmpString = "$($IMAPSetting.Server) allows plain text login, but has no insecure port bindings"
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

Run-IMAP002

