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