<#
.SYNOPSIS
Run-ExchangeAnalyzer.ps1 - An Exchange Server Configuration Analyzer
#>

#region Start parameters

[CmdletBinding()]
param ()
#endregion


#region Start variables

#...................................
# Variables
#...................................

$now = Get-Date											
$shortdate = $now.ToShortDateString()					#Short date format for reports, logs, emails

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$report = @()
$reportFile = "$($myDir)\ExchangeAnalyzerReport.html"


#endregion


#region Get Tests from XML

#Check for presence of Tests.xml file and exit if not found.
if (!(Test-Path "$MyDir\Tests.xml"))
{
    Write-Warning "Tests.xml file not found."
    EXIT
}

[xml]$TestsFile = Get-Content "$MyDir\Tests.xml"
$ExchangeAnalyzerTests = @($TestsFile.Tests)

#endregion Get Tests from XML


#region Functions
#...................................
# Functions
#...................................

#This function rolls up test results into an object
Function Get-TestResultObject($TestID, $PassedList, $FailedList)
{
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
    Write-Verbose "Fetching Exchange build numbers from TechNet"
    
    $URL = "https://technet.microsoft.com/en-us/library/hh135098(v=exchg.160).aspx"
    $WebPage = Invoke-WebRequest -Uri $URL
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
    return $ExchangeBuildNumbers
}

#This function tests each Exchange server to determine whether it is running the latest
#build for that version of Exchange.
Function Run-EXSRV001()
{
    $TestID = "EXSRV001"
    Write-Verbose "Starting test $TestID"

    $PassedList = @()
    $FailedList = @()

    $TechNetBuilds = Get-ExchangeBuildNumbers
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
                $Index = $Exchange2013Builds."Build Number".IndexOf("$buildnumber")
                $buildage = New-TimeSpan -Start ($Exchange2013Builds[$index]."Release Date") -End $now
            }
            if ($adv -like "Version 15.1*")
            {
                $MajorVersion = "15.01"
                $buildnumber = "$MajorVersion.$MinorVersion"
                $Index = $Exchange2016Builds."Build Number".IndexOf("$buildnumber")
                $buildage = New-TimeSpan -Start ($Exchange2013Builds[$index]."Release Date") -End $now
            }

            Write-Verbose "$server is N-$index"
            
            if ($index -eq 0)
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

    $ReportObj = Get-TestResultObject $TestID $PassedList $FailedList

    return $ReportObj
}

#This function tests each Exchange site to determine whether more than one CAS URL/namespace
#exists for each HTTPS service.
Function Run-CAS001()
{
    $TestID = "CAS001"
    Write-Verbose "Starting test $TestID"

    $PassedList = @()
    $FailedList = @()

    $ClientAccessServers = @($ExchangeServers | Where {$_.IsClientAccessServer -and $_.AdminDisplayVersion -like "Version 15.*"})
    $sites = @($ClientAccessServers | Group-Object -Property:Site | Select Name)

    # Get the URLs for each site, and if more than one URL exists for a HTTPS service in the same site it is
    # considered a fail.
    foreach ($site in $sites)
    {
        $SiteName = ($Site.Name).Split("/")[-1]

        Write-Verbose "Processing $SiteName"

        $SiteOWAInternalUrls = @()
        $SiteOWAExternalUrls = @()

        $SiteECPInternalUrls = @()
        $SiteECPExternalUrls = @()

        $SiteOABInternalUrls = @()
        $SiteOABExternalUrls = @()
    
        $SiteRPCInternalUrls = @()
        $SiteRPCExternalUrls = @()

        $SiteEWSInternalUrls = @()
        $SiteEWSExternalUrls = @()

        $SIteMAPIInternalUrls = @()
        $SiteMAPIExternalUrls = @()

        $SiteActiveSyncInternalUrls = @()
        $SiteActiveSyncExternalUrls = @()

        $SiteAutodiscoverUrls = @()

        $CASinSite = @($ClientAccessServers | Where {$_.Site -eq $site.Name})

        Write-Verbose "Getting OWA Urls"

        foreach ($CAS in $CASinSite)
        {
            $CASOWAUrls = @(Get-OWAVirtualDirectory -Server $CAS -AdPropertiesOnly | Select InternalURL,ExternalURL)
            foreach ($CASOWAUrl in $CASOWAUrls)
            {
                if (!($SiteOWAInternalUrls -Contains $CASOWAUrl.InternalURL.AbsoluteUri.ToLower()) -and ($CASOWAUrl.InternalURL.AbsoluteUri.ToLower() -ne $null))
                {
                    $SiteOWAInternalUrls += $CASOWAUrl.InternalURL.AbsoluteUri.ToLower()
                }
                if (!($SiteOWAExternalUrls -Contains $CASOWAUrl.ExternalURL.AbsoluteUri.ToLower()) -and ($CASOWAUrl.ExternalURL.AbsoluteUri.ToLower() -ne $null))
                {
                    $SiteOWAExternalUrls += $CASOWAUrl.ExternalUrl.AbsoluteUri.ToLower()
                }
            }
        }

        if ($SiteOWAInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteOWAExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting ECP Urls"

        foreach ($CAS in $CASinSite)
        {
            $CASECPUrls = @(Get-ECPVirtualDirectory -Server $CAS -AdPropertiesOnly | Select InternalURL,ExternalURL)
            foreach ($CASECPUrl in $CASECPUrls)
            {
                if (!($SiteECPInternalUrls -Contains $CASECPUrl.InternalURL.AbsoluteUri.ToLower()) -and ($CASECPUrl.InternalURL.AbsoluteUri.ToLower() -ne $null))
                {
                    $SiteECPInternalUrls += $CASECPUrl.InternalURL.AbsoluteUri.ToLower()
                }
                if (!($SiteECPExternalUrls -Contains $CASECPUrl.ExternalURL.AbsoluteUri.ToLower()) -and ($CASECPUrl.ExternalURL.AbsoluteUri.ToLower() -ne $null))
                {
                    $SiteECPExternalUrls += $CASECPUrl.ExternalUrl.AbsoluteUri.ToLower()
                }
            }
        }

        if ($SiteECPInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteECPExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting OAB Urls"

        foreach ($CAS in $CASinSite)
        {
            $CASOABUrls = @(Get-OABVirtualDirectory -Server $CAS -AdPropertiesOnly | Select InternalURL,ExternalURL)
            foreach ($CASOABUrl in $CASOABUrls)
            {
                if (!($SiteOABInternalUrls -Contains $CASOABUrl.InternalURL.AbsoluteUri.ToLower()) -and ($CASOABUrl.InternalURL.AbsoluteUri.ToLower() -ne $null))
                {
                    $SiteOABInternalUrls += $CASOABUrl.InternalURL.AbsoluteUri.ToLower()
                }
                if (!($SiteOABExternalUrls -Contains $CASOABUrl.ExternalURL.AbsoluteUri.ToLower()) -and ($CASOABUrl.ExternalURL.AbsoluteUri.ToLower() -ne $null))
                {
                    $SiteOABExternalUrls += $CASOABUrl.ExternalUrl.AbsoluteUri.ToLower()
                }
            }
        }

        if ($SiteOABInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteOABExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting RPC Urls"

        foreach ($CAS in $CASinSite)
        {
            $OA = Get-OutlookAnywhere -Server $CAS -AdPropertiesOnly | Select InternalHostName,ExternalHostName
            [string]$OAInternalHostName = $OA.InternalHostName
            [string]$OAExternalHostName = $OA.ExternalHostName

            [string]$OAInternalUrl = "https://$OAInternalHostName/rpc"
            [string]$OAExternalUrl = "https://$OAExternalHostName/rpc"

            if (!($SiteRPCInternalUrls -Contains $OAInternalUrl) -and ($OAInternalHostName -ne $null))
            {
                $SiteRPCInternalUrls += $OAInternalUrl
            }
            if (!($SiteRPCExternalUrls -Contains $OAExternalUrl) -and ($OAExternalHostName -ne $null) -and ($OAExternalHostName -ne ""))
            {
                $SiteRPCExternalUrls += $OAExternalUrl
            }
        }

        if ($SiteRPCInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteRPCExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting EWS Urls"

        foreach ($CAS in $CASinSite)
        {
            $CASEWSUrls = @(Get-WebServicesVirtualDirectory -Server $CAS -AdPropertiesOnly | Select InternalURL,ExternalURL)
            foreach ($CASEWSUrl in $CASEWSUrls)
            {
                if (!($SiteEWSInternalUrls -Contains $CASEWSUrl.InternalURL.AbsoluteUri.ToLower()) -and ($CASEWSUrl.InternalURL.AbsoluteUri.ToLower() -ne $null))
                {
                    $SiteEWSInternalUrls += $CASEWSUrl.InternalURL.AbsoluteUri.ToLower()
                }
                if (!($SiteEWSExternalUrls -Contains $CASEWSUrl.ExternalURL.AbsoluteUri.ToLower()) -and ($CASEWSUrl.ExternalURL.AbsoluteUri.ToLower() -ne $null))
                {
                    $SiteEWSExternalUrls += $CASEWSUrl.ExternalUrl.AbsoluteUri.ToLower()
                }
            }
        }

        if ($SiteEWSInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteEWSExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting MAPI Urls"

        foreach ($CAS in $CASinSite)
        {
            $CASMAPIUrls = @(Get-MAPIVirtualDirectory -Server $CAS -AdPropertiesOnly | Select InternalURL,ExternalURL)
            foreach ($CASMAPIUrl in $CASMAPIUrls)
            {
                if (!($SiteMAPIInternalUrls -Contains $CASMAPIUrl.InternalURL.AbsoluteUri.ToLower()) -and ($CASMAPIUrl.InternalURL.AbsoluteUri.ToLower() -ne $null))
                {
                    $SiteMAPIInternalUrls += $CASMAPIUrl.InternalURL.AbsoluteUri.ToLower()
                }
                if (!($SiteMAPIExternalUrls -Contains $CASMAPIUrl.ExternalURL.AbsoluteUri.ToLower()) -and ($CASMAPIUrl.ExternalURL.AbsoluteUri.ToLower() -ne $null))
                {
                    $SiteMAPIExternalUrls += $CASMAPIUrl.ExternalUrl.AbsoluteUri.ToLower()
                }
            }
        }

        if ($SiteMAPIInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteMAPIExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting ActiveSync Urls"

        foreach ($CAS in $CASinSite)
        {
            $CASActiveSyncUrls = @(Get-ActiveSyncVirtualDirectory -Server $CAS -AdPropertiesOnly | Select InternalURL,ExternalURL)
            foreach ($CASActiveSyncUrl in $CASActiveSyncUrls)
            {
                if (!($SiteActiveSyncInternalUrls -Contains $CASActiveSyncUrl.InternalURL.AbsoluteUri.ToLower()) -and ($CASActiveSyncUrl.InternalURL.AbsoluteUri.ToLower() -ne $null))
                {
                    $SiteActiveSyncInternalUrls += $CASActiveSyncUrl.InternalURL.AbsoluteUri.ToLower()
                }
                if (!($SiteActiveSyncExternalUrls -Contains $CASActiveSyncUrl.ExternalURL.AbsoluteUri.ToLower()) -and ($CASActiveSyncUrl.ExternalURL.AbsoluteUri.ToLower() -ne $null))
                {
                    $SiteActiveSyncExternalUrls += $CASActiveSyncUrl.ExternalUrl.AbsoluteUri.ToLower()
                }
            }
        }

        if ($SiteActiveSyncInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteActiveSyncExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting AutoDiscover Urls"

        foreach ($CAS in $CASinSite)
        {
            $CASServer = Get-ClientAccessServer $CAS.Name
            [string]$AutodiscoverSCP = ($CASServer).AutoDiscoverServiceInternalUri.AbsoluteUri.ToLower()
            $CASAutodiscoverUrl = $AutodiscoverSCP.Replace("/Autodiscover.xml","")
            if (!($SiteAutodiscoverUrls -Contains $CASAutodiscoverUrl)) {$SiteAutodiscoverUrls += $CASAutodiscoverUrl}
        }

        if ($SiteAutodiscoverUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        #If the site is not in FailedList by now, add it to $PassedList
        if ($FailedList -notcontains $SiteName) { $PassedList += $SiteName }
    }

    #Roll the object to be returned to the results
    $ReportObj = Get-TestResultObject $TestID $PassedList $FailedList

    return $ReportObj
}

#endregion


#region Main Script
#...................................
# Main Script
#...................................


#region -Basic Data Collection
#Collect information about the Exchange organization, databases, DAGs, and servers to be
#re-used throughout the script.

Write-Verbose "Collecting data about the Exchange organization"

$ExchangeOrganization = Get-OrganizationConfig
$ExchangeServers = @(Get-ExchangeServer)
$ExchangeDatabases = @(Get-MailboxDatabase)
$ExchangeDAGs = @(Get-DatabaseAvailabilityGroup)

#endregion -Basic Data Collection


#region -Start Exchange Server Tests

#region --EXSRV001: Exchange Servers are running the latest build
$EXSRV001 = Run-EXSRV001
$report += $EXSRV001
#endregion --EXSRV001

#endregion - End Exchange Server Tests

#region -Start Client Access Tests

#region --CAS001: Check if multiple namespaces exist for a protocol within the same AD site.
$CAS001 = Run-CAS001
$report += $CAS001
#endregion --CAS001

#endregion -Client Access tests

#region -Generate Report

Write-Verbose "Generating HTML report"

#HTML HEAD with styles
$htmlhead="<html>
			<style>
			BODY{font-family: Arial; font-size: 8pt;}
			H1{font-size: 16px;}
			H2{font-size: 14px;}
			H3{font-size: 12px;}
			TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
			TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
			TD{border: 1px solid black; padding: 5px; }
			td.pass{background: #7FFF00;}
			td.warn{background: #FFE600;}
			td.fail{background: #FF0000; color: #ffffff;}
			td.info{background: #85D4FF;}
			</style>
			<body>
			<h1 align=""center"">Exchange Analyzer Report</h1>
			<h3 align=""center"">Generated: $now</h3>"


#Build a list of report categories
$reportcategories = $report | Group-Object -Property TestCategory | Select Name

#Create report HTML for each category
foreach ($reportcategory in $reportcategories)
{
    $categoryHtmlTable = $null
    
    #Create HTML table headings
    $categoryHtmlHeader = "<h3>Category: $($reportcategory.Name)</h3>
					        <p>
					        <table>
					        <tr>
					        <th>Test ID</th>
					        <th>Test Category</th>
					        <th>Test Name</th>
					        <th>Test Outcome</th>
					        <th>Passed Objects</th>
					        <th>Failed Objects</th>
					        <th>Comments</th>
					        <th>Reference</th>
					        </tr>"

    $categoryHtmlTable += $categoryHtmlHeader

    #Generate each HTML table row
    foreach ($reportline in ($report | Where {$_.TestCategory -eq $reportcategory.Name}))
    {
        $HtmlTableRow = "<tr>"
        $htmltablerow += "<td>$($reportline.TestID)</td>"
		$htmltablerow += "<td>$($reportline.TestCategory)</td>"
		$htmltablerow += "<td>$($reportline.TestName)</td>"
    
        Switch ($reportline.TestOutcome)
        {	
            "Passed" {$htmltablerow += "<td class=""pass"">$($reportline.TestOutcome)</td>"}
            "Failed" {$htmltablerow += "<td class=""fail"">$($reportline.TestOutcome)</td>"}
            default {$htmltablerow += "<td>$($reportline.TestOutcome)</td>"}
		}
		
        if ($($reportline.PassedObjects).Count -gt 0)
        {
            $ul = "<ul>"
            foreach ($object in $reportline.PassedObjects)
            {
                $ul += "<li>$object</li>"
            }
            $ul += "</ul>"
            $htmltablerow += "<td>$ul</td>"
        }
        else
        {
            $htmltablerow += "<td>n/a</td>"
        }

        if ($($reportline.FailedObjects).Count -gt 0)
        {
            $ul = "<ul>"
            foreach ($object in $reportline.FailedObjects)
            {
                $ul += "<li>$object</li>"
            }
            $ul += "</ul>"
            $htmltablerow += "<td>$ul</td>"
        }
        else
        {
            $htmltablerow += "<td>n/a</td>"
        }
		
        $htmltablerow += "<td>$($reportline.Comments)</td>"
		$htmltablerow += "<td><a href=""$($reportline.Reference)"">$($reportline.Reference)</a></td>"
    
        $categoryHtmlTable += $HtmlTableRow
    }

    $categoryHtmlTable += "</table></p>"

    #Add the category to the full report
    $bodyHtml += $categoryHtmlTable
}

$htmltail = "</body>
			</html>"

#Roll the final HTML by assembling the head, body, and tail
$reportHtml = $htmlhead + $bodyHtml + $htmltail
$reportHtml | Out-File $reportFile -Force

#endregion Generate Report


#endregion Main Script
Write-Verbose "Finished."
#...................................
# Finished
#...................................