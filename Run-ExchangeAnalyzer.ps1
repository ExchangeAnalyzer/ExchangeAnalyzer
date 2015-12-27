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

#This function retrieves the Client Access URLs for HTTPS services
Function Get-ExchangeURLs()
{
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

#This function tests each Exchange server to determine whether it is running the latest
#build for that version of Exchange.
Function Run-EXSRV001()
{
    $TestID = "EXSRV001"
    Write-Verbose "----- Starting test $TestID"

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

    $ReportObj = Get-TestResultObject $TestID $PassedList $FailedList

    return $ReportObj
}

#This function tests each Exchange site to determine whether more than one CAS URL/namespace
#exists for each HTTPS service.
Function Run-CAS001()
{
    $TestID = "CAS001"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()

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
            $CASOWAUrls = @($CASURLs | Where {$_.Name -ieq $CAS.Name} | Select OWAInternal,OWAExternal)
            foreach ($CASOWAUrl in $CASOWAUrls)
            {
                if (!($SiteOWAInternalUrls -Contains $CASOWAUrl.OWAInternal.ToLower()) -and ($CASOWAUrl.OWAInternal.ToLower() -ne $null))
                {
                    $SiteOWAInternalUrls += $CASOWAUrl.OWAInternal.ToLower()
                }
                if (!($SiteOWAExternalUrls -Contains $CASOWAUrl.OWAExternal.ToLower()) -and ($CASOWAUrl.OWAExternal.ToLower() -ne $null))
                {
                    $SiteOWAExternalUrls += $CASOWAUrl.OWAExternal.ToLower()
                }
            }
        }

        if ($SiteOWAInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteOWAExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting ECP Urls"
        foreach ($CAS in $CASinSite)
        {
            $CASECPUrls = @($CASURLs | Where {$_.Name -ieq $CAS.Name} | Select ECPInternal,ECPExternal)
            foreach ($CASECPUrl in $CASECPUrls)
            {
                if (!($SiteECPInternalUrls -Contains $CASECPUrl.ECPInternal.ToLower()) -and ($CASECPUrl.ECPInternal.ToLower() -ne $null))
                {
                    $SiteECPInternalUrls += $CASECPUrl.ECPInternal.ToLower()
                }
                if (!($SiteECPExternalUrls -Contains $CASECPUrl.ECPInternal.ToLower()) -and ($CASECPUrl.ECPInternal.ToLower() -ne $null))
                {
                    $SiteECPExternalUrls += $CASECPUrl.ECPExternal.ToLower()
                }
            }
        }

        if ($SiteECPInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteECPExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting OAB Urls"
        foreach ($CAS in $CASinSite)
        {
            $CASOABUrls = @($CASURLs | Where {$_.Name -ieq $CAS.Name} | Select OABInternal,OABExternal)
            foreach ($CASOABUrl in $CASOABUrls)
            {
                if (!($SiteOABInternalUrls -Contains $CASOABUrl.OABInternal.ToLower()) -and ($CASOABUrl.OABInternal.ToLower() -ne $null))
                {
                    $SiteOABInternalUrls += $CASOABUrl.OABInternal.ToLower()
                }
                if (!($SiteOABExternalUrls -Contains $CASOABUrl.OABExternal.ToLower()) -and ($CASOABUrl.OABExternal.ToLower() -ne $null))
                {
                    $SiteOABExternalUrls += $CASOABUrl.OABExternal.ToLower()
                }
            }
        }

        if ($SiteOABInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteOABExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting RPC Urls"
        foreach ($CAS in $CASinSite)
        {
            $OA = @($CASURLs | Where {$_.Name -ieq $CAS.Name} | Select OAInternal,OAExternal)
            [string]$OAInternalHostName = $OA.OAInternal
            [string]$OAExternalHostName = $OA.OAExternal

            [string]$OAInternalUrl = "https://$($OAInternalHostName.ToLower())/rpc"
            [string]$OAExternalUrl = "https://$($OAExternalHostName.ToLower())/rpc"

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
            $CASEWSUrls = @($CASURLs | Where {$_.Name -ieq $CAS.Name} | Select EWSInternal,EWSExternal)
            foreach ($CASEWSUrl in $CASEWSUrls)
            {
                if (!($SiteEWSInternalUrls -Contains $CASEWSUrl.EWSInternal.ToLower()) -and ($CASEWSUrl.EWSInternal.ToLower() -ne $null))
                {
                    $SiteEWSInternalUrls += $CASEWSUrl.EWSInternal.ToLower()
                }
                if (!($SiteEWSExternalUrls -Contains $CASEWSUrl.EWSExternal.ToLower()) -and ($CASEWSUrl.EWSExternal.ToLower() -ne $null))
                {
                    $SiteEWSExternalUrls += $CASEWSUrl.EWSExternal.ToLower()
                }
            }
        }

        if ($SiteEWSInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteEWSExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting MAPI Urls"
        foreach ($CAS in $CASinSite)
        {
            $CASMAPIUrls = @($CASURLs | Where {$_.Name -ieq $CAS.Name} | Select MAPIInternal,MAPIExternal)
            foreach ($CASMAPIUrl in $CASMAPIUrls)
            {
                if (!($SiteMAPIInternalUrls -Contains $CASMAPIUrl.MAPIInternal.ToLower()) -and ($CASMAPIUrl.MAPIInternal.ToLower() -ne $null))
                {
                    $SiteMAPIInternalUrls += $CASMAPIUrl.MAPIInternal.ToLower()
                }
                if (!($SiteMAPIExternalUrls -Contains $CASMAPIUrl.MAPIExternal.ToLower()) -and ($CASMAPIUrl.MAPIExternal.ToLower() -ne $null))
                {
                    $SiteMAPIExternalUrls += $CASMAPIUrl.MAPIExternal.ToLower()
                }
            }
        }

        if ($SiteMAPIInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteMAPIExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting ActiveSync Urls"
        foreach ($CAS in $CASinSite)
        {
            $CASActiveSyncUrls = @($CASURls | Where {$_.Name -eq $CAS.Name} | Select EASInternal,EASExternal)
            foreach ($CASActiveSyncUrl in $CASActiveSyncUrls)
            {
                if (!($SiteActiveSyncInternalUrls -Contains $CASActiveSyncUrl.EASInternal.ToLower()) -and ($CASActiveSyncUrl.EASInternal.ToLower() -ne $null))
                {
                    $SiteActiveSyncInternalUrls += $CASActiveSyncUrl.EASInternal.ToLower()
                }
                if (!($SiteActiveSyncExternalUrls -Contains $CASActiveSyncUrl.EASExternal.ToLower()) -and ($CASActiveSyncUrl.EASExternal.ToLower() -ne $null))
                {
                    $SiteActiveSyncExternalUrls += $CASActiveSyncUrl.EASExternal.ToLower()
                }
            }
        }

        if ($SiteActiveSyncInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteActiveSyncExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting AutoDiscover Urls"
        foreach ($CAS in $CASinSite)
        {
            #$CASServer = Get-ClientAccessServer $CAS.Name
            $AutoDUrl = @($CASURLs | Where {$_.Name -ieq $CAS.Name} | Select AutoD)
            [string]$AutodiscoverSCP = $AutoDUrl
            $CASAutodiscoverUrl = $AutodiscoverSCP.Replace("/Autodiscover.xml","")
            if (!($SiteAutodiscoverUrls -Contains $CASAutodiscoverUrl.ToLower())) {$SiteAutodiscoverUrls += $CASAutodiscoverUrl.ToLower()}
        }

        if ($SiteAutodiscoverUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        #If the site is not in FailedList by now, add it to $PassedList
        if ($FailedList -notcontains $SiteName) { $PassedList += $SiteName }
    }

    #Roll the object to be returned to the results
    $ReportObj = Get-TestResultObject $TestID $PassedList $FailedList

    return $ReportObj
}

#This function tests each CAS URL to determine whether it contains a server FQDN
Function Run-CAS002()
{
    $TestID = "CAS002"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()

    foreach ($CAS in $ClientAccessServers)
    {
        $HasUrlsWithFQDN = $false        
        $serverFQDN = $CAS.Fqdn.ToLower()
        $serverURLs = @($CASURLs | Where {$_.Name -ieq $CAS.Name})
        $propertyNames = @($serverURLs | Get-Member -Type NoteProperty | Where {$_.Name -ne "Name"} | Select Name)
        foreach ($name in $propertyNames)
        {
            Write-Verbose "Checking URL $($name.Name)"
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
$ExchangeDatabases = @(Get-MailboxDatabase -Status)
$ExchangeDAGs = @(Get-DatabaseAvailabilityGroup)

#endregion -Basic Data Collection


#region -Start Exchange Server Tests

#region --EXSRV001: Exchange Servers are running the latest build
$EXSRV001 = Run-EXSRV001
$report += $EXSRV001
#endregion --EXSRV001

#endregion - End Exchange Server Tests

#region -Start Client Access Tests

#Get all Exchange HTTPS URLs to use for CAS tests
Write-Verbose "Determining Client Access servers"
$ClientAccessServers = @($ExchangeServers | Where {$_.IsClientAccessServer -and $_.AdminDisplayVersion -like "Version 15.*"})
Write-Verbose "Collecting Exchange URLs"
$CASURLs = Get-ExchangeURLs


#region --CAS001: Check if multiple namespaces exist for a protocol within the same AD site.
$CAS001 = Run-CAS001
$report += $CAS001
#endregion --CAS001

#region --CAS002: Check that CAS URLs don't contain server FQDNs.
$CAS002 = Run-CAS002
$report += $CAS002
#endregion --CAS002

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
            ul{list-style: inside; padding-left: 0px;}
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
		$htmltablerow += "<td><a href=""$($reportline.Reference)"">More Info</a></td>"
    
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

iex $reportFile
#...................................
# Finished
#...................................