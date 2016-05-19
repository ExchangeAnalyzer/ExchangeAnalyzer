#This function tests each Exchange server to verify a supported version of the .NET framework istalled
Function Run-EXSRV003()
{
    [CmdletBinding()]
    param()

    $TestID = "EXSRV003"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $WarningList = @()
    $InfoList = @()
    $ErrorList = @()

    foreach ($server in $exchangeservers) {
        # Set the inital value
        $ServerName = $server.name
        $up = $true

        try {
            $Registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$ServerName)
        } catch {
            $up = $false
        }

        if ($up -eq $true) {
            # Check the registry path
            $RegKey = $Registry.OpenSubKey("SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full")

            if ($RegKey -ne $null) {
                # Get the value
                [int]$NETFxVersion = $RegKey.GetValue("Release")
            }
        
            # Results of the check
            switch ($NETFxVersion)
            {
                {($_ -ge 378389) -and ($_ -lt 378675)} {$WarningList += "$ServerName - .NET Framework 4.5";Write-Verbose ".NET FrameWork version 4.5 detected on server $ServerName.  4.5.2 is strongly recommended."}
                {($_ -ge 378675) -and ($_ -lt 379893)} {$WarningList += "$ServerName - .NET Framework 4.5.1";Write-Verbose ".NET FrameWork version 4.5.1 detected on server $ServerName.  4.5.2 is strongly recommended."}
                {($_ -ge 379893) -and ($_ -lt 393297)} {$PassedList += "$ServerName - .NET Framework 4.5.2";Write-Verbose ".NET FrameWork version 4.5.2 detected on server $ServerName.  This is the recommended version."}
		        {($_ -ge 393297) -and ($_ -lt 394271)} {$FailedList += "$ServerName - .NET Framework 4.6";Write-Verbose ".NET FrameWork version 4.6 detected on server $ServerName.  4.6 is not supported for Exchange.  4.5.2 is strongly recommended."}
                {($_ -ge 394271)} {$FailedList += "$ServerName - .NET Framework 4.6.1";Write-Verbose ".NET FrameWork version 4.6.1 or later detected on server $ServerName.  4.6.1 is not supported for Exchange.  4.5.2 is strongly recommended."}
                default {$WarningList += $ServerName;Write-Verbose("Unable to determine .NET FrameWork version on server $ServerName")}
            }
        }
        else
        {
            $WarningList += "$ServerName - unable to connect to registry"
            Write-Verbose "Unable to connect to registry of $ServerName"
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

Run-EXSRV003