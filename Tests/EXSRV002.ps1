#requires -Modules ExchangeAnalyzer

#This function tests each Exchange server to determine whether it is running the latest
#build for that version of Exchange.
Function Run-EXSRV002()
{
    [CmdletBinding()]
    param()

    $TestID = "EXSRV002"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $ErrorList = @()

    Write-Verbose "Scraping TechNet page"
    $TechNetBuilds = Get-ExchangeBuildNumbers -Verbose:($PSBoundParameters['Verbose'] -eq $true)
    if ($TechNetBuilds -like "An error occurred*")
    {
        $ErrorList += $TechNetBuilds
    }
    else
    {
        $Exchange2013Builds = @()
        $Exchange2016Builds = @()

        #Process results to rename properties, convert release date strings
        #to proper date values, and exclude legacy versions
        foreach ($build in $TechNetBuilds)
        {
            if ($build.'Build number' -like "15.00.*")
            {
                $BuildProperties = [Ordered]@{
                        'Product Name'="Exchange Server 2013"
                        'Description'=$build.'Product name'
                        'Build Number'=$build.'Build number'
                        'Release Date'=$(Get-Date $build.'Release date')
                        }
                $buildObject = New-Object -TypeName PSObject -Prop $BuildProperties
                $Exchange2013Builds += $buildObject
            }
            elseif ($build.'Build number' -like "15.01.*")
            {
                $BuildProperties = [Ordered]@{
                        'Product Name'="Exchange Server 2016"
                        'Description'=$build.'Product name'
                        'Build Number'=$build.'Build number'
                        'Release Date'=$(Get-Date $build.'Release date')
                        }
                $buildObject = New-Object -TypeName PSObject -Prop $BuildProperties
                $Exchange2016Builds += $buildObject
            }
        }

        $Exchange2013Builds = $Exchange2013Builds | Sort 'Product Name','Release Date' -Descending
        $Exchange2016Builds = $Exchange2016Builds | Sort 'Product Name','Release Date' -Descending
    
        foreach($server in $ExchangeServers)
        {
            Write-Verbose "Checking $server"
            $adv = $server.AdminDisplayVersion
            if ($adv -like "Version 15.*")
            {
                Write-Verbose "$server is at least Exchange 2013"

                $build = ($adv -split "Build ").Trim()[1]
                $build = $build.SubString(0,$build.Length-1)
                $arrbuild = $build.Split(".")
            
                [int]$tmp = $arrbuild[0]
                $buildpart1 = "{0:D4}" -f $tmp
            
                [int]$tmp = $arrbuild[1]
                $buildpart2 = "{0:D3}" -f $tmp
            
                $MinorVersion = "$buildpart1.$buildpart2"

                if ($adv -like "Version 15.0*")
                {
                    $MajorVersion = "15.00"
                    $buildnumber = "$MajorVersion.$MinorVersion"
                    $CASndex = $Exchange2013Builds."Build Number".IndexOf("$buildnumber")
                    $buildage = New-TimeSpan -Start ($Exchange2013Builds[$CASndex]."Release Date") -End $now
                }
                if ($adv -like "Version 15.1*")
                {
                    $MajorVersion = "15.01"
                    $buildnumber = "$MajorVersion.$MinorVersion"
                    $CASndex = $Exchange2016Builds."Build Number".IndexOf("$buildnumber")
                    $buildage = New-TimeSpan -Start ($Exchange2013Builds[$CASndex]."Release Date") -End $now
                }

                Write-Verbose "$server is N-$CASndex"
            
                if ($CASndex -eq 0)
                {
                    $PassedList += $($Server.Name)
                }
                else
                {
                    $tmpstring = "$($Server.Name) ($($buildage.Days) days old)"
                    Write-Verbose "Adding to fail list: $tmpstring"
                    $FailedList += $tmpstring
                }
            }
            else
            {
                #Skip servers earlier than v15.0
                Write-Verbose "$server is earlier than Exchange 2013"
            }
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

Run-EXSRV002