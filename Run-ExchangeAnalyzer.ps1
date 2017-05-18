<#
.SYNOPSIS
Exchange Analyzer - An Exchange Server 2013/2016 Best Practices Analyzer

.DESCRIPTION 
Exchange Analyzer is a PowerShell tool that scans an Exchange Server 2013 or 2016 organization
and reports on compliance with best practices.

Please refer to the installation and usage instructions at http://exchangeanalyzer.com

.OUTPUTS
Results are output to a HTML report.

.PARAMETER FileName
Specifies the relative or absolute path to the output file.

If this parameter is not supplied, the default ExchangeAnalyzerReport-date-time.html
output path will be utilized.

.PARAMETER Verbose
Verbose output is displayed in the Exchange management shell.

.EXAMPLE
.\Run-ExchangeAnalyzer.ps1
Runs the Exchange Analyzer.

.EXAMPLE
.\Run-ExchangeAnalyzer.ps1 -FileName C:\ExchangeReports\ContosoExchange.html
Runs the Exchange Analyzer outputting results to C:\ExchangeReports\ContosoExchange.html

.EXAMPLE
.\Run-ExchangeAnalyzer.ps1 -Verbose
Runs the Exchange Analyzer with -Verbose output.

.LINK
https://exchangeanalyzer.com

.NOTES

*** Credits ***

----- Core Team -----

- Paul Cunningham
    * Website:	https://paulcunningham.me
    * Twitter:	https://twitter.com/paulcunningham

- Michael B Smith
    * Website: http://theessentialexchange.com/
    * Twitter: https://twitter.com/essentialexch

- Brian Desmond
    * Website: http://www.briandesmond.com/
    * Twitter: https://twitter.com/brdesmond

- Damian Scoles
    * Website: https://justaucguy.wordpress.com/

----- Additional Contributions -----

https://github.com/ExchangeAnalyzer/ExchangeAnalyzer/wiki/Contributors


*** Change Log ***

Refer to releases:
https://github.com/ExchangeAnalyzer/ExchangeAnalyzer/releases

*** License ***

The MIT License (MIT)

Copyright (c) 2017 Paul Cunningham, exchangeanalyzer.com

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
#>


#requires -Modules ExchangeAnalyzer
#requires -Modules ActiveDirectory
#requires -Modules DnsClient

#region Start parameters

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]$FileName
)
#endregion


#region Start variables

#...................................
# Variables
#...................................

$now = Get-Date											
$shortdate = $now.ToShortDateString()					#Short date format for reports, logs, emails

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$report = @()

# What file types may be provided when creating the output file?
$supportedOutputFileTypes = @("html","htm")

#endregion


#region Get Tests from XML

#Check for presence of Tests.xml file and exit if not found.
if (!(Test-Path "$($MyDir)\Data\Tests.xml"))
{
    Write-Warning "Tests.xml file not found."
    EXIT
}

[xml]$TestsFile = Get-Content "$($MyDir)\Data\Tests.xml"
$ExchangeAnalyzerTests = @($TestsFile.Tests)

#endregion Get Tests from XML


#region Main Script
#...................................
# Main Script
#...................................

#region -File Name Generation
# Generate an output filename for the script

if ($FileName) {
    # If the user has passed the filename parameter to the script, use that.
    try {
        $FileNameExtension = $FileName.Split(".")[-1]
        if ($supportedOutputFileTypes -icontains $FileNameExtension) {
            if ([System.IO.Path]::IsPathRooted($FileName)) {
                # Path provided by user is absolute; use it as is.
                $reportFile = $FileName
                # Ensure the folder exists
                $ReportFileFolder = Split-Path $FileName -Parent
                if (-not (Test-Path $ReportFileFolder -PathType Container -ErrorAction SilentlyContinue)) {
                    # Folder does not exist, create it
                    Write-Verbose "$ReportFileFolder does not exist. Attempting to create it."
                    try {
                        $null = New-Item -ItemType Directory -Force -Path $ReportFileFolder
                        Write-Verbose "$ReportFilefolder was created."
                    } catch {
                        throw "Folder $ReportFileFolder does not exist, and was unable to be created."
                    }
                }
            } else {
                # Path provided by user is relative; base it in $MyDir.
                $reportFile = Join-Path $myDir $FileName
            }
        } else {
            throw "Unsupported file type: $FileNameExtension"
        }
    } catch {
        throw "Unable to validate passed -FileName as a relative or absolute path. Exception: $($_.ToString())"
    }
        
} else {
    # If the user did not pass a filename, generate one based on date/time.
    $reportFile = "$($MyDir)\ExchangeAnalyzerReport-$(Get-Date -UFormat %Y%m%d-%H%M).html"
}
#endregion

#region -Basic Data Collection
#Collect information about the Exchange organization, databases, DAGs, and servers to be
#re-used throughout the script.

#publish this for use by anyone
$global:ExAPropertyBag = @{}

$ProgressActivity = "Initializing"

$msgString = "Collecting data about the Exchange organization"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 0
Write-Verbose $msgString

try
{
    Write-Progress -Activity $ProgressActivity -Status "Get-OrganizationConfig" -PercentComplete 1
    $ExchangeOrganization = Get-OrganizationConfig -ErrorAction STOP
    
    Write-Progress -Activity $ProgressActivity -Status "Get-ExchangeServer" -PercentComplete 1
    $ExchangeServersAll = @(Get-ExchangeServer -ErrorAction STOP)
    $ExchangeServers = @($ExchangeServersAll | Where-Object {$_.AdminDisplayVersion -like "Version 15.*"})
    Write-Verbose "$($ExchangeServers.Count) Exchange servers found."

    #Check for supported servers before continuing
    if (($ExchangeServers | Where-Object {$_.AdminDisplayVersion -like "Version 15.*"}).Count -eq 0)
    {
        Write-Warning "No Exchange 2013 or later servers were found. Exchange Analyzer is exiting."
        EXIT
    }

    Write-Progress -Activity $ProgressActivity -Status "Get-MailboxDatabase" -PercentComplete 2
    $ExchangeDatabases = @(Get-MailboxDatabase -Status -ErrorAction STOP)
    Write-Verbose "$($ExchangeDatabases.Count) databases found."

    #Do not use -Status switch here as it causes an error to be thrown. DAG status should be
    #queried later after filtering DAG list to only v15.x DAGs.
    Write-Progress -Activity $ProgressActivity -Status "Get-DatabaseAvailabilityGroup" -PercentComplete 2
    $ExchangeDAGs = @(Get-DatabaseAvailabilityGroup -ErrorAction STOP)
    Write-Verbose "$($ExchangeDAGs.Count) DAGs found."

    Write-Progress -Activity $ProgressActivity -Status 'Exchange Registry Settings' -PercentComplete 2
    foreach ($s in $ExchangeServers)
    {
        Set-ExAServerProperty -Server $s.Name -Property 'ExchangeInstallPath' -Value (Get-ExARegistryValue -Host $s.Name -Hive LocalMachine -Key 'SOFTWARE\Microsoft\ExchangeServer\V15\Setup' -Value 'MsiInstallPath')
    }

    Write-Progress -Activity $ProgressActivity -Status "Get-ADDomain" -PercentComplete 3
    $ADDomain = Get-ADDomain -ErrorAction STOP
 
    Write-Progress -Activity $ProgressActivity -Status "Get-ADForest" -PercentComplete 3
    $ADForest = Get-ADForest -ErrorAction STOP
 
    Write-Progress -Activity $ProgressActivity -Status "Get-ADDomainController" -PercentComplete 3
    $ADDomainControllers = @(Get-ADDomainController -filter * -ErrorAction STOP)
    Write-Verbose "$($ADDomainControllers.Count) Domain Controller(s) found."
}
catch
{
    Write-Warning "An error has occurred during basic data collection."
    Write-Warning $_.Exception.Message
    EXIT
}

#Get all Exchange HTTPS URLs to use for CAS tests
$msgString = "Determining Client Access servers"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 4
Write-Verbose $msgString
$ClientAccessServers = @($ExchangeServers | Where-Object {$_.IsClientAccessServer -and $_.AdminDisplayVersion -like "Version 15.*"})
Write-Verbose "$($ClientAccessServers.Count) Client Access servers found."

$msgString = "Collecting Exchange URLs from Client Access servers"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 4
Write-Verbose $msgString
$CASURLs = @(Get-ExchangeURLs $ClientAccessServers -Verbose:($PSBoundParameters['Verbose'] -eq $true))
Write-Verbose "CAS URLs collected from $($CASURLs.Count) servers."

#Get all POP settings for CAS/MBX servers
$msgString = "Collecting POP settings from Client Access and Mailbox servers"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 4
Write-Verbose $msgString
#This needs to be processed as a foreach to work in PS remoting
$AllPopSettings = @($ExchangeServers | ForEach-Object {Get-PopSettings -Server $_.Identity})

#Get all IMAP settings for CAS/MBX servers
$msgString = "Collecting IMAP settings from Client Access and Mailbox servers"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 4
Write-Verbose $msgString
#This needs to be processed as a foreach to work in PS remoting
$AllImapSettings = @($ExchangeServers | ForEach-Object {Get-ImapSettings -Server $_.Identity})


#endregion -Basic Data Collection

#region -Run tests
#The tests listed in Tests.xml will be performed as long as the corresponding PowerShell
#script for that test ID is found in the \Tests folder.
$ProgressActivity = "Running Tests"
$NumberOfTests = ($ExchangeAnalyzerTests.Test).Count
$TestCount = 0
foreach ($Test in $ExchangeAnalyzerTests.ChildNodes.Id)
{
	$TestDescription = ($exchangeanalyzertests.Childnodes | Where-Object {$_.Id -eq $Test}).Description
    $TestCount += 1
    $pct = $TestCount/$NumberOfTests * 100
	Write-Progress -Activity $ProgressActivity -Status "Test $($TestCount) of $($NumberOfTests): $TestDescription" -PercentComplete $pct

    if (Test-Path "$($MyDir)\Tests\$($Test).ps1")
    {
        #Escape any spaces in path prior to running Invoke-Expression
        $command = "$($MyDir)\Tests\$($Test).ps1" -replace ' ','` '
        $testresult = Invoke-Expression -Command $command
        $report += $testresult
    }
    else
    {
        Write-Warning "$($Test) script wasn't found in $($MyDir)\Tests folder."
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
			BODY{font-family: Arial; font-size: 10pt;}
			H1{font-size: 22px;}
			H2{font-size: 20px; padding-top: 10px;}
			H3{font-size: 16px; padding-top: 8px;}
			TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt; table-layout: fixed;}
            TABLE.testresults{width: 850px;}
            TABLE.summary{text-align: center; width: auto;}
			TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
            TH.summary{width: 80px;}
            TH.test{width: 120px;}
            TH.description{width: 150px;}
            TH.outcome{width: 50px}
            TH.comments{width: 120px;}
            TH.details{width: 270px;}
            TH.reference{width: 60px;}
			TD{border: 1px solid black; padding: 5px; vertical-align: top; }
			td.pass{background: #7FFF00;}
			td.warn{background: #FFE600;}
			td.fail{background: #FF0000; color: #ffffff;}
			td.info{background: #85D4FF;}
            ul{list-style: inside; padding-left: 0px;}
			</style>
			<body>"

#HTML intro
$IntroHtml="<h1>Exchange Analyzer Report</h1>
			<p><strong>Generated:</strong> $now</p>
            <p><strong>Organization:</strong> $($ExchangeOrganization.Name)</p>
            <p>The following guidelines apply to this report:
            <ul>
                <li>This tests included in this report are documented on the <a href=""https://github.com/ExchangeAnalyzer/ExchangeAnalyzer/wiki/Exchange-Analyzer-Tests"">Exchange Analyzer Wiki</a>.</li>
                <li>Click the ""More Info"" link for each test to learn more about that test, what a pass or fail means, and recommendations for how to respond.</li>
                <li>A test can fail if it can't complete successfully, or if a condition was encountered that requires manual assessment.</li>
                <li>For some organizations a failed test may be due to a deliberate design or operational decision.</li>
                <li>Please review the <a href=""https://github.com/ExchangeAnalyzer/ExchangeAnalyzer/wiki/Frequently-Asked-Questions"">Frequently Asked Questions</a> if you have any further questions.</li>
            </ul>
            </p>"

#Count of test results
$PassedItems = @($report | Where-Object {$_.TestOutcome -eq "Passed"})
$TotalPassed = $PassedItems.Count

$WarningItems = @($report | Where-Object {$_.TestOutcome -eq "Warning"})
$TotalWarning = $WarningItems.Count

$FailedItems = @($report | Where-Object {$_.TestOutcome -eq "Failed"})
$TotalFailed = $FailedItems.Count

$InfoItems = @($report | Where-Object {$_.TestOutcome -eq "Info"})
$TotalInfo = $InfoItems.Count

#HTML summary table
$SummaryTableHtml  = "<h2>Summary:</h2>
                      <p>
                      <table class=""summary"">
                      <tr>
                      <th class=""summary"">Passed</th>
                      <th class=""summary"">Warning</th>
                      <th class=""summary"">Failed</th>
                      <th class=""summary"">Info</th>
                      </tr>
                      <tr>
                      <td class=""pass"">$TotalPassed</td>
                      <td class=""warn"">$TotalWarning</td>
                      <td class=""fail"">$TotalFailed</td>
                      <td class=""info"">$TotalInfo</td>
                      </tr>
                      </table>
                      </p>"

if ($TotalFailed -gt 0)
{
    $SummaryFailedItemsHtml = "<p>Failed items:</p>
                              <ul>"

    foreach ($FailedItem in $FailedItems)
    {
        $SummaryFailedItemsHtml += "<li>$($FailedItem.Comments)</li>"
    }

    $SummaryFailedItemsHtml += "</ul>
                                </p>"

    $SummaryTableHtml += $SummaryFailedItemsHtml
}

if ($TotalWarning -gt 0)
{
    $SummaryWarningItemsHtml = "<p>Warning items:</p>
                                <ul>"

    foreach ($WarningItem in $WarningItems)
    {
        $SummaryWarningItemsHtml += "<li>$($WarningItem.Comments)</li>"
    }

    $SummaryWarningItemsHtml += "</ul>
                                 </p>"

    $SummaryTableHtml += $SummaryWarningItemsHtml
}


#Build table for summary of Exchange Servers in organization
$ExchangeServersSummaryHtml = $null
$ExchangeServersSummaryHtml += "<p>Summary of Exchange Servers:</p>
                                <p>
                                <table>
                                <tr>
                                <th>Name</th>
                                <th>Version</th>
                                <th>.NET Framework</th>
                                <th>Site</th>
                                <th>Domain</th>
                                <th>Roles</th>
                                <th>Edition</th>
                                </tr>"

#Will include all servers, even those not tested by Exchange Analyzer
foreach ($Server in $ExchangeServersAll)
{
    #See Issue #62 in Github for why this ToString() is required for compatiblity with 2013/2016.
    $ServerADSite = ($ExchangeServersAll | Where-Object {$_.Name -ieq $($server.Name)}).Site.ToString() 
    
    $ExchangeServersSummaryHtml += "<tr>
                                    <td>$($Server.Name)</td>
                                    <td>$(Get-ExAServerProperty -Server $Server -Property "BuildDescription")</td>
                                    <td>$(Get-ExAServerProperty -Server $Server -Property ".NET Framework")</td>
                                    <td>$($ServerADSite.Split("/")[-1])</td>
                                    <td>$($Server.Domain)</td>
                                    <td>$($Server.ServerRole)</td>
                                    <td>$($Server.Edition)</td>
                                    </tr>"
}

$ExchangeServersSummaryHtml += "</table>
                                </p>"


#Build table of CAS URLs
$CASURLSummaryHtml = $null
$CASURLSummaryHtml += "<p>Summary of Client Access URLs/Namespaces:</p>"

foreach ($server in $CASURLs)
{
    #See Issue #62 in Github for why this ToString() is required for compatiblity with 2013/2016.
    $ServerADSite = ($ExchangeServers | Where-Object {$_.Name -ieq $($server.Name)}).Site.ToString() 

    $CASURLSummaryHtml += "<table>
                            <tr>
                            <th colspan=""3"">Server: $($server.Name), Site: $($ServerADSite.Split("/")[-1])</th>
                            </tr>
                            <tr>
                            <th>Service</th>
                            <th>Internal URL</th>
                            <th>External Url</th>
                            </tr>
                            <tr>
                            <td>Outlook AnyWhere-Object</td>
                            <td>$($server.OAInternal)</td>
                            <td>$($server.OAExternal)</td>
                            </tr>
                            <tr>
                            <td>MAPI/HTTP</td>
                            <td>$($server.MAPIInternal)</td>
                            <td>$($server.MAPIExternal)</td>
                            </tr>
                            <tr>
                            <td>Outlook on the web (OWA)</td>
                            <td>$($server.OWAInternal)</td>
                            <td>$($server.OWAExternal)</td>
                            </tr>
                            <tr>
                            <td>Exchange Control Panel</td>
                            <td>$($server.ECPInternal)</td>
                            <td>$($server.ECPExternal)</td>
                            </tr>
                            <tr>
                            <td>ActiveSync</td>
                            <td>$($server.EASInternal)</td>
                            <td>$($server.EASExternal)</td>
                            </tr>
                            <tr>
                            <td>Offline Address Book</td>
                            <td>$($server.OABInternal)</td>
                            <td>$($server.OABExternal)</td>
                            </tr>
                            <tr>
                            <td>Exchange Web Access</td>
                            <td>$($server.EWSInternal)</td>
                            <td>$($server.EWSExternal)</td>
                            </tr>
                            <tr>
                            <td>AutoDiscover (SCP)</td>
                            <td>$($server.AutoDSCP)</td>
                            <td>n/a</td>
                            </tr>
                            </table>
                            </p>"
}

#Build a table of mailbox databases
$DatabaseSummaryHtml = $null
$DatabaseSummaryHtml += "<p>Summary of mailbox databases:</p>
                        <p>
                        <table>
                        <tr>
                        <th>Database Name</th>
                        <th>Type</th>
                        <th>Server/DAG</th>
                        <th>Copies</th>
                        </tr>"

foreach ($Database in $ExchangeDatabases)
{
    $DatabaseSummaryHtml += "<tr>
                            <td>$($Database.Name)</td>
                            <td>$($Database.MasterType)</td>
                            <td>$($Database.MasterServerOrAvailabilityGroup)</td>
                            <td>$($Database.DatabaseCopies.Count)</td>
                            </tr>"
}

$DatabaseSummaryHtml += "</table>
                        </p>"


#Build a list of report categories
$reportcategories = $report | Group-Object -Property TestCategory | Select-Object Name

#Create report HTML for each category
foreach ($reportcategory in $reportcategories)
{
    $categoryHtmlTable = $null
    
    #Create HTML table headings
    if ($($reportcategory.Name) -eq "Exchange Servers")
    {
        $categoryHtmlHeader = "<h2>Category: $($reportcategory.Name)</h2>"
        $categoryHtmlHeader += $ExchangeServersSummaryHtml
        $categoryHtmlHeader += "<p>Results for $($reportcategory.Name) tests:</p>"
    }
    elseif ($($reportcategory.Name) -eq "Databases")
    {
        $categoryHtmlHeader = "<h2>Category: $($reportcategory.Name)</h2>"
        $categoryHtmlHeader += $DatabaseSummaryHtml
        $categoryHtmlHeader += "<p>Results for $($reportcategory.Name) tests:</p>"
    }
    elseif ($($reportcategory.Name) -eq "Client Access")
    {
        $categoryHtmlHeader = "<h2>Category: $($reportcategory.Name)</h2>"
        $categoryHtmlHeader += $CASURLSummaryHtml
        $categoryHtmlHeader += "<p>Results for $($reportcategory.Name) tests:</p>"
    }
    else
    {
        $categoryHtmlHeader = "<h2>Category: $($reportcategory.Name)</h2>
                                <p>Results for $($reportcategory.Name) tests:</p>"
    }
    $categoryHtmlHeader += "<p>
					        <table class=""testresults"">
					        <tr>
					        <th class=""test"">Test</th>
                            <th class=""description"">Description</th>
					        <th class=""outcome"">Outcome</th>
					        <th class=""comments"">Comments</th>
					        <th class=""details"">Details</th>
					        <th class=""reference"">Reference</th>
					        </tr>"

    $categoryHtmlTable += $categoryHtmlHeader

    #Generate each HTML table row
    foreach ($reportline in ($report | Where-Object {$_.TestCategory -eq $reportcategory.Name}))
    {
        $HtmlTableRow = "<tr>"
		$htmltablerow += "<td>$($reportline.TestName)</td>"
		$htmltablerow += "<td>$($reportline.TestDescription)</td>"    
        Switch ($reportline.TestOutcome)
        {	
            "Passed" {$htmltablerow += "<td class=""pass"">$($reportline.TestOutcome)</td>"}
            "Failed" {$htmltablerow += "<td class=""fail"">$($reportline.TestOutcome)</td>"}
            "Warning" {$HtmlTableRow += "<td class=""warn"">$($reportline.TestOutcome)</td>"}
            "Info" {$HtmlTableRow += "<td class=""info"">$($reportline.TestOutcome)</td>"}
            default {$htmltablerow += "<td>$($reportline.TestOutcome)</td>"}
		}

        $htmltablerow += "<td>$($reportline.Comments)</td>"
		
        #Build list of passed, warning, failed, and info objects for report details column
        $TestDetails = $null

        if ($($reportline.InfoObjects).Count -gt 0)
        {
            $TestDetails += "<p>Info items:</p><ul>"
            foreach ($object in $reportline.InfoObjects)
            {
                $TestDetails += "<li>$object</li>"
            }
            $TestDetails += "</ul>"
        }
        else
        {
            #$TestDetails += "<p>Info objects:</p><ul><li>n/a</li></ul>"
        }

        if ($($reportline.PassedObjects).Count -gt 0)
        {
            $TestDetails += "<p>Passed items:</p><ul>"
            foreach ($object in $reportline.PassedObjects)
            {
                $TestDetails += "<li>$object</li>"
            }
            $TestDetails += "</ul>"
        }
        else
        {
            #$TestDetails += "<p>Passed objects:</p><ul><li>n/a</li></ul>"
        }

        if ($($reportline.WarningObjects).Count -gt 0)
        {
            $TestDetails += "<p>Warning items:</p><ul>"
            foreach ($object in $reportline.WarningObjects)
            {
                $TestDetails += "<li>$object</li>"
            }
            $TestDetails += "</ul>"
        }
        else
        {
            #$TestDetails += "<p>Warning objects:</p><ul><li>n/a</li></ul>"
        }

        if ($($reportline.FailedObjects).Count -gt 0)
        {
            $TestDetails += "<p>Failed items:</p><ul>"
            foreach ($object in $reportline.FailedObjects)
            {
                $TestDetails += "<li>$object</li>"
            }
            $TestDetails += "</ul>"
        }
        else
        {
            #$TestDetails += "<p>Failed objects:</p><ul><li>n/a</li></ul>"
        }

        $htmltablerow += "<td>$TestDetails</td>"
				
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

$htmltail = "<p>Report created by <a href=""http://exchangeanalyzer.com"">Exchange Analyzer</a></p>
            </body>
			</html>"

#Roll the final HTML by assembling the head, body, and tail
$reportHtml = $htmlhead + $IntroHtml + $SummaryTableHtml + $bodyHtml + $htmltail
$reportHtml | Out-File $reportFile -Force

#endregion Generate Report


#endregion Main Script


$msgString = "Finished"
Write-Progress -Activity $ProgressActivity -Status $msgString -PercentComplete 100
Write-Verbose $msgString

#Escape any spaces in path prior to running Invoke-Expression
$command = $reportfile -replace ' ','` '
Invoke-Expression -Command $command
#...................................
# Finished
#...................................
