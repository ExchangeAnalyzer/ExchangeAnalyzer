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

    #Faking the test for now
    
    foreach($server in $ExchangeServers)
    {
        $PassedList += $($server.Name)
    }
    
    $FailedList += "Foo"

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