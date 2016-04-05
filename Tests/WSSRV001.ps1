#requires -Modules ExchangeAnalyzer

#This function checks to see if the Exchange Install Drive has greater than 30% free space
Function Run-WSSRV001()
{

   [CmdletBinding()]
    param()

    $TestID = 'WSSRV001'
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $WarningList = @()
    $InfoList = @()
    $ErrorList = @()

    foreach ($server in $ExchangeServers) 
    {
        $serverName = $server.name

        Write-Verbose "Checking $serverName"

        # Null out variables for clean results
        $reg = $null

		$reg = Get-ExAServerProperty -Server $serverName -Property 'ExchangeInstallPath'

        if ($reg -eq $null)
        {
            Write-Verbose "ExchangeInstallPath property null for $serverName"
            $FailedList += @($serverName)
            continue
        }

        [string]$exchangeInstallDrive = [IO.Path]::GetPathRoot($reg)
        # trim the trailing \
        $exchangeInstallDrive = $exchangeInstallDrive.Substring(0, 2)

        # Get free Space for all drives on the Exchange Server
        ##TODO: Consider publishing this to the property bag if we need to test multiple things about this disk
        $logicalDisk = Get-ExAWmiObject -Computer $serverName -Class 'Win32_LogicalDisk' -Property @('DeviceID', 'FreeSpace', 'Size') -Filter "DeviceID='$ExchangeInstallDrive'"
    
        $free = $LogicalDisk.FreeSpace
        $size = $LogicalDisk.Size

        Write-Verbose "Server $serverName Disk $exchangeInstallDrive is $size with $free free space"

        # Calculate percent free space
        [int]$percentFree = ($free / $size) * 100

        
        if ($percentFree -lt 15) 
        {
            Write-Verbose "Install Drive on $serverName has less than 15% free space."
            $FailedList += $($ServerName)
        } 
        elseif (($percentFree -gt 15) -and ($percentFree -lt 30)) 
        {
            Write-Verbose "Install Drive on $serverName has 15% - 30% free space."
            $WarningList += $($ServerName)
        }
        else
        {
            Write-Verbose "Install Drive on $serverName has more than 30% free space."
            $PassedList += $($ServerName)
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

Run-WSSRV001