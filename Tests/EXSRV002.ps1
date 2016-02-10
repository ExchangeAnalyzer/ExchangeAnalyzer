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
    $WarningList = @()
    $InfoList = @()
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
                    $buildindex = $Exchange2013Builds."Build Number".IndexOf("$buildnumber")
                    Write-Verbose "Exchange version is: $($Exchange2013Builds[$buildindex]."Product Name")"
                    $buildage = New-TimeSpan -Start ($Exchange2013Builds[$buildindex]."Release Date") -End $now

                    #Fixes issue when $buildindex is -1 due to being last item in the array
                    if ($buildindex -eq "-1")
                    {
                        $buildindex = $Exchange2013Builds.Count - 1
                    }

                }
                if ($adv -like "Version 15.1*")
                {
                    $MajorVersion = "15.01"
                    $buildnumber = "$MajorVersion.$MinorVersion"
                    $buildindex = $Exchange2016Builds."Build Number".IndexOf("$buildnumber")
                    Write-Verbose "Exchange version is: $($Exchange2016Builds[$buildindex]."Product Name")"
                    $buildage = New-TimeSpan -Start ($Exchange2016Builds[$buildindex]."Release Date") -End $now

                    #Fixes issue when $buildindex is -1 due to being last item in the array
                    if ($buildindex -eq "-1")
                    {
                        $buildindex = $Exchange2016Builds.Count - 1
                    }

                }

                Write-Verbose "$server is N-$buildindex"
                $tmpstring = "$($Server.Name) ($($buildage.Days) days old)"

                if ($buildindex -eq 0)
                {
                    Write-Verbose "Adding to passed list: $tmpstring"
                    $PassedList += $($Server.Name)
                }
                elseif ($buildindex -eq 1)
                {
                    Write-Verbose "Adding to warning list: $tmpstring"
                    $WarningList += $tmpstring
                }
                else
                {        
                    Write-Verbose "Adding to fail list: $tmpstring"
                    $FailedList += $tmpstring
                }
            }
            else
            {
                #Skip servers earlier than v15.0
                Write-Verbose "$server is earlier than Exchange 2013 and will not be checked."
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

Run-EXSRV002