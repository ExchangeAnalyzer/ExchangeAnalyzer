$PassedList = @()
$FailedLIst = @()
foreach( $server in $servers)
{
    if( ( test-14 $server ) )
    {
        $PassedList += $server
    }
    else
    {
        $FailedList += $server
    }
}
 
$result = @{
    TestName = "Check Mailbox servers for sufficient RAM"
    TestNumber = $testNumber
    PassedServers = $PassedList
    FailedServers = $FailedList
    IfPassedComments = "Server SERVERNAME meets the minimum memory requirements for a Mailbox Role"
    IfFailedComments = "Server SERVERNAME does not have 8GB of memory, which is the minimum requirement for the Mailbox Role."
    UrlReference = "https://technet.microsoft.com/en-us/library/aa996719(v=exchg.160).aspx"
} 
 
return $result
