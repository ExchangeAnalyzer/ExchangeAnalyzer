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
    $WarningList = @()
    $InfoList = @()
    $ErrorList = @()

    foreach ($server in $exchangeservers) {
        # Set the inital value
        $name = $server.name
        $up = $true
        $SSLv3Enabled = $null
        try {
            $Registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$name)
        } catch {
            $up = $false
        }

        if ($up -eq $true) {
            # Check the registry path
            $check1 = $registry.OpenSubKey("System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0")

            if ($check1 -ne $null) {
                # Check the next registry path
                $check2 = $registry.OpenSubKey("System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Server")
        
                if ($check2 -ne $null) {

                    # Check to see if Enabled value is present with a '0' value
                    $check3 = $registry.OpenSubKey("System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Server").GetValue("Enabled")
                    if ($Check3 -eq "0") {
                        $SSLv3Enabled = $true
                    } else {
                        $SSLv3Enabled = $false
                    }
                }
            }
        
            # Decide if the test has failed based off of the values missing or present and what value is there if present
            Switch ($SSLv3Enabled)
            {
                $true {
                        $PassedList += $name
                        Write-Verbose "SSL 3.0 is disabled on the server $name"
                    }
                $false {
                        $FailedList += $name
                        Write-Verbose "SSL 3.0 is not disabed on server $name"    
                    }
                default {
                        Write-Verbose "SSL 3.0 status could not be determined for $name"
                        $WarningList += "$name - SSL 3.0 status unknown"
                    }
            }

        }
        else
        {
            $WarningList += "$name - unable to connect to registry"
            Write-Verbose "Unable to connect to registry of $name"
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

Run-CAS003