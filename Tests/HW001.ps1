#requires -Modules ExchangeAnalyzer

#This function checks to see if the Exchange Server is Physical or Virtual.
Function Run-HW001()
{

   [CmdletBinding()]
    param()

    $TestID = "HW001"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $WarningList = @()
    $InfoList = @()
    $ErrorList = @()

    foreach ($server in $exchangeservers) {
        $name = $server.name

        # Get system information on the server
        try {
            $virtual = Get-CIMInstance -ComputerName $name -Class Win32_ComputerSystem -ErrorAction Stop
        }catch {
            $tryWMI = $true
            Write-Verbose "$($TestID): Was not able to acquire information for $name via CIM"
        }

        if( $tryWMI ) {
            ## WMI depends on RPC. CIM depends on WinRM. There are pluses and minuses to either, but in recent versions of Windows,
            ## CIM is preferred. But CIM failed, so we try WMI before we give up.
            Try {
                $virtual = Get-WmiObject -ComputerName $name -Class Win32_ComputerSystem -ErrorAction Stop
            } catch {
                Write-Verbose "$($TestID): Was not able to acquire information for $name via WMI"
                $FailedList += $($name)
                Break
            }
        }

        # Check to see if the server is VMWare
        if($virtual.Manufacturer -like "*VMWare*") {

            # Server is running on a VMWare Server
            $VirtualType = "VMWare"
            write-verbose "The server $name is virtualized on the $Virtualtype platform."
            $WarningList += $($name)
        } elseif($virtual.Manufacturer -like "*Microsoft Corporation*") {

            # Server is running on a Hyper-V Server
            $VirtualType = "Hyper-V"
            write-verbose "The server $name is virtualized on the $Virtualtype platform."
            $WarningList += $($name)
        } elseif($virtual.Manufacturer.Length -gt 0) {

            # Server is running on a Physical Server
            $InfoList = "Physical"
            write-verbose "The server $name is running on physical hardware."
        } else {

            # This server is on an unknown platform
            $VirtualType = "Unknown"
            $Infolist += $($name)
            write-verbose "The server $name is running on an unknown platform."
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

Run-HW001