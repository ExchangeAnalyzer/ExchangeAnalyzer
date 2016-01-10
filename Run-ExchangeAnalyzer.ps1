<#
.SYNOPSIS
Run-ExchangeAnalyzer.ps1 - An Exchange Server Configuration Analyzer
#>

#region Start parameters

[CmdletBinding()]
param ()
#endregion

#Import-Module ExchangeAnalyzer

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
if (!(Test-Path "$MyDir\Data\Tests.xml"))
{
    Write-Warning "Tests.xml file not found."
    EXIT
}

[xml]$TestsFile = Get-Content "$MyDir\Data\Tests.xml"
$ExchangeAnalyzerTests = @($TestsFile.Tests)

#endregion Get Tests from XML


#region Main Script
#...................................
# Main Script
#...................................


#region -Basic Data Collection
#Collect information about the Exchange organization, databases, DAGs, and servers to be
#re-used throughout the script.

$ProgressActivity = "Initializing"

$msgString = "Collecting data about the Exchange organization"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 0
Write-Verbose $msgString

try
{
    Write-Progress -Activity $ProgressActivity -Status "Get-OrganizationConfig" -PercentComplete 1
    $ExchangeOrganization = Get-OrganizationConfig -ErrorAction STOP
    
    Write-Progress -Activity $ProgressActivity -Status "Get-ExchangeServer" -PercentComplete 2
    $ExchangeServers = @(Get-ExchangeServer -ErrorAction STOP)
    Write-Verbose "$($ExchangeServers.Count) Exchange servers found."

    Write-Progress -Activity $ProgressActivity -Status "Get-MailboxDatabase" -PercentComplete 3
    $ExchangeDatabases = @(Get-MailboxDatabase -Status -ErrorAction STOP)
    Write-Verbose "$($ExchangeDatabases.Count) databases found."

    Write-Progress -Activity $ProgressActivity -Status "Get-DatabaseAvailabilityGroup" -PercentComplete 4
    $ExchangeDAGs = @(Get-DatabaseAvailabilityGroup -Status -ErrorAction STOP)
    Write-Verbose "$($ExchangeDAGs.Count) DAGs found."
}
catch
{
    Write-Warning "An error has occurred during basic data collection."
    Write-Warning $_.Exception.Message
    EXIT
}

#Get all Exchange HTTPS URLs to use for CAS tests
$msgString = "Determining Client Access servers"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 5
Write-Verbose $msgString
$ClientAccessServers = @($ExchangeServers | Where {$_.IsClientAccessServer -and $_.AdminDisplayVersion -like "Version 15.*"})
Write-Verbose "$($ClientAccessServers.Count) Client Access servers found."

$msgString = "Collecting Exchange URLs from Client Access servers"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 6
Write-Verbose $msgString
$CASURLs = @(Get-ExchangeURLs $ClientAccessServers)
Write-Verbose "$($CASURLs.Count) URLs collected."


#endregion -Basic Data Collection

#region -Run tests
#The tests listed in Tests.xml will be performed as long as the corresponding PowerShell
#script for that test ID is found in the \Tests folder.
$ProgressActivity = "Running Tests"
$NumberOfTests = ($ExchangeAnalyzerTests.Test).Count
$TestCount = 0
foreach ($Test in $ExchangeAnalyzerTests.ChildNodes.Id)
{
	$TestDescription = ($exchangeanalyzertests.Childnodes | Where {$_.Id -eq $Test}).Description
    $TestCount += 1
    $pct = $TestCount/$NumberOfTests * 100
	Write-Progress -Activity $ProgressActivity -Status "(Test $TestCount of $NumberOfTests) $($Test): $TestDescription" -PercentComplete $pct

    if (Test-Path "$MyDir\Tests\$($Test).ps1")
    {
        $testresult = Invoke-Expression -Command "$MyDir\Tests\$($Test).ps1"
        $report += $testresult
    }
    else
    {
        Write-Warning "$($Test) script wasn't found in $MyDir\Tests folder."
    }
}


#endregion -Run tests

#region -Generate Report
$ProgressActivity = "Finishing"
$msgString = "Generating HTML report"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 99
Write-Verbose $msgString

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
            "Warning" {$HtmlTableRow += "<td class=""warn"">$($reportline.TestOutcome)</td>"}
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
		
        if ($($reportline.Reference) -eq "")
        {
            $htmltablerow += "<td>No additional info</td>"
        }
        else
        {
            $htmltablerow += "<td><a href=""$($reportline.Reference)"" target=""_blank"">More Info</a></td>"
        }
        
    
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


$msgString = "Finished"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 100
Write-Verbose $msgString

iex $reportFile
#...................................
# Finished
#...................................