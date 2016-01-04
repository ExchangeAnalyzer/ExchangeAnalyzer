#This function rolls up test results into an object
Function Get-TestResultObject()
{
    [CmdletBinding()]
    param (
	    [Parameter()]
	    $ExchangeAnalyzerTests,
        
        [Parameter()]
	    $TestID,

	    [Parameter()]
	    $PassedList,

	    [Parameter()]
	    $FailedList,

	    [Parameter()]
	    $ErrorList
	)

    Write-Verbose "Rolling test result object for $TestID"
     
    if ($PassedList)
    {
        $TestComments = ($ExchangeAnalyzerTests.Test | Where {$_.Id -eq $TestID}).IfPassedComments
        $TestOutcome = "Passed"
    }

    if ($FailedList)
    {
        $TestComments = ($ExchangeAnalyzerTests.Test | Where {$_.Id -eq $TestID}).IfFailedComments
        $TestOutcome = "Failed"
    }

    if (-not $PassedList -and -not $FailedList)
    {
        $TestComments = "Test could not run."
        $TestOutcome = "Warning"
    }

    if ($ErrorList)
    {
        $TestComments = "Errors were encountered. $($ErrorList)"
        $TestOutcome = "Warning"
    }

    $result = [Ordered]@{
        TestID = $TestID
        TestCategory = ($ExchangeAnalyzerTests.Test | Where {$_.Id -eq $TestID}).Category
        TestName = ($ExchangeAnalyzerTests.Test | Where {$_.Id -eq $TestID}).Name
        TestOutcome = $TestOutcome
        PassedObjects = $PassedList
        FailedObjects = $FailedList
        Comments = $TestComments
        Reference = ($ExchangeAnalyzerTests.Test | Where {$_.Id -eq $TestID}).Reference
    }
    
    $TestResultObj = New-Object -TypeName PSObject -Property $result 
 
    return $TestResultObj
}

#This function scrapes the TechNet page for Exchange Server build numbers and release
#dates to match the build numbers for Exchange Servers in the organization.
##Reference: Lee Holmes article on extracting tables from web pages was very useful for developing this
#Link: #http://www.leeholmes.com/blog/2015/01/05/extracting-tables-from-powershells-invoke-webrequest/
Function Get-ExchangeBuildNumbers()
{
    [CmdletBinding()]
    param ()
    
    Write-Verbose "Fetching Exchange build numbers from TechNet"
    
    $URL = "https://technet.microsoft.com/en-us/library/hh135098(v=exchg.160).aspx"
    try
    {
        $WebPage = Invoke-WebRequest -Uri $URL -ErrorAction STOP
        $tables = @($WebPage.Parsedhtml.getElementsByTagName("TABLE"))

        Write-Verbose "Parsing results from web request"
        foreach ($table in $tables)
        {
            $rows = @($table.Rows)

            foreach($row in $rows)
            {
                $cells = @($row.Cells)

                ## If we’ve found a table header, remember its titles
                if($cells[0].tagName -eq "TH")
                {
                    $titles = @($cells | ForEach-Object { ("" + $_.InnerText).Trim() })
                    continue
                }

                ## If we haven’t found any table headers, make up names "P1", "P2", etc.
                if(-not $titles)
                {
                    $titles = @(1..($cells.Count + 2) | ForEach-Object { "P$_" })
                }

                ## Now go through the cells in the the row. For each, try to find the
                ## title that represents that column and create a hashtable mapping those
                ## titles to content

                $resultObject = [Ordered] @{}

                for($counter = 0; $counter -lt $cells.Count; $counter++)
                {
                    $title = $titles[$counter]
                    if(-not $title) { continue }
                    $resultObject[$title] = ("" + $cells[$counter].InnerText).Trim()
                }

                ## And finally cast that hashtable to a PSCustomObject
                [PSCustomObject] $resultObject
            }
    }
    }
    catch
    {
        Write-Warning $_.Exception.Message
        $ExchangeBuildNumbers = "An error occurred. $($_.Exception.Message)"
    }

    return $ExchangeBuildNumbers
}

#This function retrieves the Client Access URLs for HTTPS services
Function Get-ExchangeURLs()
{
    [CmdletBinding()]
    param (
        [Parameter()]
	    $ClientAccessServers
    )

    $results = @()

    foreach ($CAS in $ClientAccessServers)
    {
        Write-Verbose "Fetching URLs for $CAS"
        $OA = Get-OutlookAnywhere -Server $CAS -AdPropertiesOnly | Select InternalHostName,ExternalHostName
        $OWA = Get-OWAVirtualDirectory -Server $CAS -AdPropertiesOnly | Select InternalURL,ExternalURL
        $ECP = Get-ECPVirtualDirectory -Server $CAS -AdPropertiesOnly | Select InternalURL,ExternalURL
        $OAB = Get-OABVirtualDirectory -Server $CAS -AdPropertiesOnly | Select InternalURL,ExternalURL
        $EWS = Get-WebServicesVirtualDirectory -Server $CAS -AdPropertiesOnly | Select InternalURL,ExternalURL
        $MAPI = Get-MAPIVirtualDirectory -Server $CAS -AdPropertiesOnly | Select InternalURL,ExternalURL
        $EAS = Get-ActiveSyncVirtualDirectory -Server $CAS -AdPropertiesOnly | Select InternalURL,ExternalURL
        Write-Verbose "Fetching AutoD SCP for $CAS"
        $AutoD = Get-ClientAccessServer $CAS.Name | Select AutoDiscoverServiceInternalUri

        Write-Verbose "Creating object for CAS Urls"
        $props = [Ordered]@{
            Name = $CAS.Name
            OAInternal = $OA.InternalHostName
            OAExternal = $OA.ExternalHostName
            OWAInternal = $OWA.InternalURL.AbsoluteUri
            OWAExternal = $OWA.ExternalURL.AbsoluteUri
            ECPInternal = $ECP.InternalURL.AbsoluteUri
            ECPExternal = $ECP.ExternalURL.AbsoluteUri
            OABInternal = $OAB.InternalURL.AbsoluteUri
            OABExternal = $OAB.ExternalURL.AbsoluteUri
            EWSInternal = $EWS.InternalURL.AbsoluteUri
            EWSExternal = $EWS.ExternalURL.AbsoluteUri
            MAPIInternal = $MAPI.InternalURL.AbsoluteUri
            MAPIExternal = $MAPI.ExternalURL.AbsoluteUri
            EASInternal = $EAS.InternalURL.AbsoluteUri
            EASExternal = $EAS.ExternalURL.AbsoluteUri
            AutoDSCP = $AutoD.AutoDiscoverServiceInternalUri
            }
            
        $CASObj = New-Object -TypeName PSObject -Property $props

        $results += $CASObj
    }

    return $results
}

Export-ModuleMember -Function * -Alias * -Variable *