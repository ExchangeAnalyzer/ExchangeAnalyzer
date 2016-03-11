#requires -Modules ExchangeAnalyzer

#This function tests each Exchange server to verify if SSL 3.0 is disabled
Function Run-CAS003()
{
    [CmdletBinding()]
    param()

    $TestID = "CAS003"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $ErrorList = @()

    foreach ($server in $exchangeservers) {
        # Set the inital value
        $name = $server.name
        $up = $true
        $success = $false
        try {
            $Registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$name)
        } catch {
            $up = $false
        }

        if ($up -eq $true) {
            # Check the registry path
            $check1 = $registry.OpenSubKey("System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0")

            if ($check1 -ne $null) {
                # Check the next registry pat
                $check2 = $registry.OpenSubKey("System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Server")
        
                if ($check2 -ne $null) {

                    # Check to see if Enabled value is present with a '0' value
                    $check3 = $registry.OpenSubKey("System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Server").GetValue("Enabled")
                    if ($Check3 -eq "0") {
                        $Success = $true
                    } else {
                        $Success = $false
                    }
                }
            }
        
            # Decide if the test has failed based off of the values missing or present and what value is there if present
            If ($Success -eq $true) {
                $PassedList += $name
                write-verbose "SSL 3.0 is disabled on the server $name!!"
            } else {
                $FailedList += $name
                write-verbose "SSL 3.0 is not disabed on server $name!!"
            }
        } else {
            $FailedList += $name
            write-verbose "The server $name is down and SSL 3.0 settings cannot be verified."
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

Run-CAS003