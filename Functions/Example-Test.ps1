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

#This is your test
Function Run-TESTID()
{
    $TestID = "TESTID"
    Write-Verbose "Starting test $TestID"

    $PassedList = @()
    $FailedList = @()

    #Your test logic goes here
    
    #Populate the results    
    $PassedList += "Foo"
    $FailedList += "Bar" #An empty failed list will result in a test passing.

    #Roll the object to be returned to the results
    $ReportObj = Get-TestResultObject $TestID $PassedList $FailedList

    return $ReportObj
}