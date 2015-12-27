#These variables will be needed by most tests
#...................................
# Variables
#...................................

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$now = Get-Date

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


#...................................
# Functions
#...................................

#This function rolls up test results into an object
Function Get-TestResultObject($TestID, $PassedList, $FailedList, $ErrorList)
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

#This is your test
Function Run-TESTID()
{
    $TestID = "TESTID"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $ErrorList = @()

    #Your test logic goes here and populates the results
    $PassedList += "Foo"
    $FailedList += "Bar" #An empty failed list will result in a test passing.
    #Add any errors to $ErrorList

    #Roll the object to be returned to the results
    $ReportObj = Get-TestResultObject $TestID $PassedList $FailedList $ErrorList

    return $ReportObj
}


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

#Run your test and see the results
Run-TESTID

