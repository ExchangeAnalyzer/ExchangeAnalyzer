#requires -Modules ExchangeAnalyzer

#This function tests each CAS URL to determine whether it contains a server FQDN
Function Run-CAS002()
{
    [CmdletBinding()]
    param()

    $TestID = "CAS002"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $ErrorList = @()

    foreach ($CAS in $ClientAccessServers)
    {
        $HasUrlsWithFQDN = $false        
        $serverFQDN = $CAS.Fqdn.ToLower()
        $serverURLs = @($CASURLs | Where {$_.Name -ieq $CAS.Name})
        $propertyNames = @($serverURLs | Get-Member -Type NoteProperty | Where {$_.Name -ne "Name"} | Select Name)
        foreach ($name in $propertyNames)
        {
            Write-Verbose "Checking URL $($serverURLs."$($name.name)")"
            if ($serverURLs."$($name.name)" -icontains $serverFQDN)
            {
                $HasUrlsWithFQDN = $true
            }
        }

        if ($HasUrlsWithFQDN)
        {
            $FailedList += $($CAS.Name)
        }
        else
        {
            $PassedList += $($CAS.Name)
        }
    }

    $ReportObj = Get-TestResultObject -ExchangeAnalyzerTests $ExchangeAnalyzerTests `
                                      -TestId $TestID `
                                      -PassedList $PassedList `
                                      -FailedList $FailedList `
                                      -ErrorList $ErrorList `
                                      -Verbose:($PSBoundParameters['Verbose'] -eq $true)
    return $ReportObj
}

Run-CAS002
