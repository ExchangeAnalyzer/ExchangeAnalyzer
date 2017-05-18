#requires -Modules ExchangeAnalyzer

#This functions tests each mailbox database to determine whether the database has been
#backed up in the last 24 hours.
Function Run-DB001()
{
    [CmdletBinding()]
    param()

    $TestID = "DB001"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $WarningList = @()
    $InfoList = @()
    $ErrorList = @()

    $mailboxdatabases = @($ExchangeDatabases | Where {$_.AdminDisplayVersion -like "Version 15.*" -and $_.Recovery -ne $true})

    if ($mailboxdatabases)
    {
	    Write-Verbose "$($mailboxdatabases.count) mailbox databases found"
        
        #Check each database for most recent backup timestamp
        foreach ($db in $mailboxdatabases)
        {
	        Write-Verbose "Checking $db"
            if ($db.Mounted -eq $false)
            {
                $tmpString = "$($db.name) is dismounted"
                Write-Verbose $tmpString
                $WarningList += $tmpString
            }
            elseif ($db.Mounted -eq $null)
            {
                #Indicates Information Store service was stopped on the server
                $tmpString = "$($db.Name) could not be reached"
                Write-Verbose $tmpString
                $WarningList += $tmpString
            }
            else
            {
                if ( -not $db.LastFullBackup -and -not $db.LastIncrementalBackup -and -not $db.LastDifferentialBackup)
	                                            {
		        #No backup timestamp was present. This means either the database has
		        #never been backed up, or it was unreachable when this script ran
		        $LastBackups = @{
                                  Never="n/a"
                                }
                
                #Write-Verbose "Last backup of $($db.name) was $($LatestBackup.Value)."
	        }
	            else
	        {
                if (-not $db.LastIncrementalBackup)
                {
                    $LastInc = "Never"
                }
                else
                {
                    [int]$LastInc = ($now.ToUniversalTime() - $db.LastIncrementalBackup.ToUniversalTime()).TotalHours
                }

                if (-not $db.LastDifferentialBackup)
                {
                    $LastDiff = "Never"
                }
                else
                {
                    [int]$LastDiff = ($now.ToUniversalTime() - $db.LastDifferentialBackup.ToUniversalTime()).TotalHours
                }

                if (-not $db.LastFullBackup)
                {
                    $LastFull = "Never"
                }
                else
                {
                    [int]$LastFull = ($now.ToUniversalTime() - $db.LastFullBackup.ToUniversalTime()).TotalHours
                }

                $LastBackups = @{
                        Incremental=$LastInc
                        Differential=$LastDiff
                        Full=$LastFull
                        }
            }

                $LatestBackup = ($LastBackups.GetEnumerator() | Sort-Object -Property Value)[0]
                if ($($LatestBackup.Value) -eq "n/a")
                {
                    Write-Verbose "$($db.name) has never been backed up."
                }
                else
                {
                    Write-Verbose "Last backup of $($db.name) was $($LatestBackup.Key) $($LatestBackup.Value) hours ago"
                }
            
                if ($($LatestBackup.Value) -eq "n/a")
                {
                    $FailedList += "$($db.Name) (Never backed up)"
                }
                elseif ($($LatestBackup.Value.ToInt32($null)) -gt 24)
                {
                    $FailedList += "$($db.Name) ($($LatestBackup.Value) hrs ago)"
                }
                elseif ($($LatestBackup.Value) -ieq "Never")
                {
                    $FailedList += "$($db.name) (Never backed up)"
                }
                else
                {
                    $PassedList += "$($db.Name)"
                }
            }
        }
    }
    else
    {
	    Write-Verbose "No mailbox databases found"
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

Run-DB001